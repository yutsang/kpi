"""OpenAI-compatible LLM client with disk cache and strict retry.

Guarantees:
  - chat_json: retries with exponential backoff up to max_retries; raises if still failing.
  - map_parallel: round-based retry. Every input MUST produce a non-error result before
    returning, or RuntimeError is raised after max_round_retries rounds.
  - Disk cache: identical (model, system, user, schema_hint) tuples skip the API entirely
    and are resilient to crashes / Ctrl-C.
"""
from __future__ import annotations

import json
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, Iterable

import httpx
from tqdm import tqdm

from .cache import JsonlCache
from .io_setup import ProgressRefresher
from .rate_limiter import RateLimiter


class StopRequested(Exception):
    """Signalled when the run is aborted (fail-fast or KeyboardInterrupt)."""


def _parse_json_lenient(text: str) -> dict:
    """Parse JSON from a model response, tolerating ```json fences and surrounding prose."""
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z]*\s*", "", cleaned)
        cleaned = re.sub(r"\s*```$", "", cleaned)
        cleaned = cleaned.strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{[\s\S]*\}", cleaned)
        if not match:
            raise
        return json.loads(match.group(0))


class LLMClient:
    def __init__(self, cfg: dict, cache_path: Path | str | None = None):
        llm = cfg["llm"]
        # Accept one key (api_key: "...") OR several (api_keys: ["...", "..."]).
        _keys = llm.get("api_keys") or ([llm["api_key"]] if llm.get("api_key") else [])
        if not llm.get("api_base") or not _keys or not llm.get("chat_model"):
            raise ValueError("config.llm requires api_base, api_key (or api_keys), chat_model")
        verify_ssl = bool(llm.get("verify_ssl", False))
        timeout = llm.get("request_timeout", 30)
        self.api_base = str(llm["api_base"]).rstrip("/")
        # Each key gets its OWN rate limiter, and requests round-robin across keys, so the
        # company's per-key "1 req/s" cap is respected while N keys give ~N× throughput.
        # 1 key → identical to the old single-limiter behaviour.
        self._keys = [str(k).strip() for k in _keys if str(k).strip()]
        self.api_key = self._keys[0]   # back-compat: external refs / cache key model only
        self._http = httpx.Client(verify=verify_ssl, timeout=timeout)
        self.model = llm["chat_model"]
        self.use_json_mode = bool(llm.get("use_json_mode", True))
        self.max_retries = llm.get("max_retries", 3)
        self.concurrency = llm.get("concurrency", 16)
        self.max_round_retries = llm.get("max_round_retries", 2)
        self.round_backoff_seconds = llm.get("round_backoff_seconds", 10)
        self.fail_fast_threshold = int(llm.get("fail_fast_threshold", 0) or 0)
        _rps = float(llm.get("max_requests_per_second", 0) or 0)
        self._key_limiters = [RateLimiter(_rps) for _ in self._keys]
        self._key_idx = 0
        self._key_lock = threading.Lock()
        if len(self._keys) > 1:
            print(f"[llm] {len(self._keys)} API keys x {_rps or 'inf'} req/s each "
                  f"-> ~{(_rps * len(self._keys)) or 'inf'} req/s aggregate", flush=True)
        self.cache = JsonlCache(Path(cache_path)) if cache_path else None
        self._stop_event = threading.Event()
        self._error_log_path: Path | None = None
        if cache_path:
            self._error_log_path = Path(cache_path).with_suffix(".errors.log")
        self._err_log_lock = threading.Lock()
        self._heartbeat = None

    def _pick_key(self) -> int:
        """Round-robin the next key index (thread-safe). Spreading retries across keys
        also helps when one key is momentarily rate-limited."""
        with self._key_lock:
            i = self._key_idx
            self._key_idx = (self._key_idx + 1) % len(self._keys)
        return i

    def attach_heartbeat(self, heartbeat) -> None:
        self._heartbeat = heartbeat

    def _hb_update(self, done: int, ok: int, err: int) -> None:
        hb = self._heartbeat
        if hb is not None:
            try:
                hb.update(done=done, ok=ok, err=err)
            except Exception:
                pass

    def _log_error(self, label: str, err_msg: str, item_repr: str = "") -> None:
        if self._error_log_path is None:
            return
        with self._err_log_lock:
            with self._error_log_path.open("a", encoding="utf-8") as f:
                f.write(f"[{label}] {err_msg}\n")
                if item_repr:
                    f.write(f"    item: {item_repr[:300]}\n")

    def request_stop(self) -> None:
        """Signal all in-flight chat_json calls to abort. Also closes the HTTP client to
        unblock threads stuck inside an HTTP read (which would otherwise wait for the
        full request_timeout)."""
        self._stop_event.set()
        try:
            self._http.close()
        except Exception:
            pass

    def reset_stop(self) -> None:
        self._stop_event.clear()

    def chat_json(
        self,
        system: str,
        user: str,
        schema_hint: str = "",
        use_cache: bool = True,
    ) -> dict:
        cache_key = JsonlCache.make_key(self.model, system, user, schema_hint)
        if use_cache and self.cache is not None:
            hit = self.cache.get(cache_key)
            if hit is not None:
                return hit

        url = self.api_base + "/chat/completions"
        body: dict = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user + ("\n\n" + schema_hint if schema_hint else "")},
            ],
            "temperature": 0,
        }
        if self.use_json_mode:
            body["response_format"] = {"type": "json_object"}

        last_err: Exception | None = None
        for attempt in range(self.max_retries):
            if self._stop_event.is_set():
                raise StopRequested("aborted by client (stop event set)")
            ki = self._pick_key()
            headers = {
                "Authorization": f"Bearer {self._keys[ki]}",
                "Content-Type": "application/json",
            }
            self._key_limiters[ki].acquire(self._stop_event)
            if self._stop_event.is_set():
                raise StopRequested("aborted during rate-limit wait")
            try:
                resp = self._http.post(url, headers=headers, json=body)
                resp.raise_for_status()
                payload = resp.json()

                # Company rate-limit envelope: HTTP 200 but body says 'REQUEST TOO FREQUENT'.
                if isinstance(payload, dict) and (
                    payload.get("code") == 30001
                    or (payload.get("flag") is False and "choices" not in payload)
                ):
                    msg = payload.get("msg", "rate limited")
                    last_err = RuntimeError(f"server rate limit: {msg}")
                    backoff = min(5 * (2 ** attempt), 60)
                    if self._stop_event.wait(timeout=backoff):
                        raise StopRequested("aborted during rate-limit backoff")
                    continue

                choices = payload.get("choices") if isinstance(payload, dict) else None
                if not choices:
                    err = payload.get("error") if isinstance(payload, dict) else None
                    raise RuntimeError(
                        f"LLM returned no choices (error={err!r}, raw={str(payload)[:300]})"
                    )
                choice = choices[0] or {}
                msg_field = choice.get("message") or {}
                content = msg_field.get("content") or ""
                if not str(content).strip():
                    raise RuntimeError(f"LLM returned empty content: {str(payload)[:300]}")
                result = _parse_json_lenient(content)
                if use_cache and self.cache is not None:
                    self.cache.put(cache_key, result)
                return result
            except StopRequested:
                raise
            except Exception as e:
                last_err = e
                if self._stop_event.wait(timeout=min(2 ** attempt, 30)):
                    raise StopRequested("aborted by client during backoff")
        raise RuntimeError(f"LLM call failed after {self.max_retries} retries: {last_err}")

    def map_parallel(
        self,
        items: Iterable[Any],
        fn: Callable[[Any], Any],
        desc: str = "llm",
    ) -> list:
        items = list(items)
        results: list[Any] = [None] * len(items)
        pending = list(range(len(items)))

        def is_failure(r: Any) -> bool:
            if r is None:
                return True
            if isinstance(r, dict) and "_error" in r:
                return True
            return False

        for round_idx in range(self.max_round_retries):
            # Concurrency ladder: round 1 = full, then halve each retry until 1.
            # Lets us push hard initially and back off so the LLM endpoint stops 429-ing.
            cur_concurrency = max(1, self.concurrency // (2 ** round_idx))
            label = (
                f"{desc} round {round_idx + 1}/{self.max_round_retries} "
                f"(concurrency={cur_concurrency})"
            )
            ok_count = 0
            err_count = 0
            consecutive_fail = 0
            seen_errors: dict[str, int] = {}
            abort_reason: str | None = None
            ex = ThreadPoolExecutor(max_workers=cur_concurrency)
            try:
                futures = {ex.submit(fn, items[i]): i for i in pending}
                bar = tqdm(
                    as_completed(futures),
                    total=len(futures),
                    desc=label,
                    miniters=1,
                    mininterval=0.0,
                    smoothing=0.1,
                    file=sys.stderr,
                )
                refresher = ProgressRefresher(bar, interval=0.5).start()
                for fut in bar:
                    i = futures[fut]
                    try:
                        results[i] = fut.result()
                        ok_count += 1
                        consecutive_fail = 0
                    except Exception as e:
                        err_msg = str(e)
                        results[i] = {"_error": err_msg}
                        err_count += 1
                        consecutive_fail += 1
                        self._log_error(label, err_msg, repr(items[i]))
                        if err_msg not in seen_errors:
                            seen_errors[err_msg] = 1
                            tqdm.write(
                                f"[{label}] NEW ERROR ({len(seen_errors)} unique so far): {err_msg}"
                            )
                            sys.stderr.flush()
                        else:
                            seen_errors[err_msg] += 1
                        if (
                            self.fail_fast_threshold > 0
                            and consecutive_fail >= self.fail_fast_threshold
                        ):
                            abort_reason = (
                                f"fail-fast: {consecutive_fail} consecutive failures "
                                f"(threshold={self.fail_fast_threshold}). Last: {last_err}"
                            )
                            tqdm.write(f"[{label}] {abort_reason}")
                            sys.stderr.flush()
                            self.request_stop()
                            for f in list(futures.keys()):
                                f.cancel()
                            break
                    bar.set_postfix_str(
                        f"ok={ok_count} err={err_count} streak={consecutive_fail} uniq_errs={len(seen_errors)}"
                    )
                    bar.refresh()
                    sys.stderr.flush()
                    self._hb_update(done=ok_count + err_count, ok=ok_count, err=err_count)
            except KeyboardInterrupt:
                tqdm.write(f"[{label}] KeyboardInterrupt — cancelling pending futures")
                sys.stderr.flush()
                self.request_stop()
                for f in list(futures.keys()):
                    f.cancel()
                ex.shutdown(wait=False, cancel_futures=True)
                raise
            finally:
                try:
                    refresher.stop()
                except Exception:
                    pass
                ex.shutdown(wait=False, cancel_futures=True)
                self.reset_stop()
            if seen_errors:
                tqdm.write(
                    f"[{label}] Round summary: ok={ok_count} err={err_count} "
                    f"unique_errors={len(seen_errors)}"
                )
                for msg, cnt in sorted(seen_errors.items(), key=lambda kv: -kv[1])[:5]:
                    tqdm.write(f"    {cnt}x  {msg}")
                if self._error_log_path is not None:
                    tqdm.write(f"  Full errors logged to: {self._error_log_path}")
                sys.stderr.flush()
            if abort_reason:
                raise RuntimeError(abort_reason)

            pending = [i for i in range(len(items)) if is_failure(results[i])]
            if not pending:
                return results

            print(
                f"  {len(pending)} items still failing after round {round_idx + 1}; "
                f"sleeping {self.round_backoff_seconds}s before retry."
            )
            time.sleep(self.round_backoff_seconds)

        # Done with all rounds. Failed items keep their {"_error": ...} markers
        # and the caller decides how to handle them (typically: flag for manual review).
        return results

    def map_batched(
        self,
        items: list,
        cache_user_fn: Callable[[Any], str],
        cache_system: str,
        batch_user_fn: Callable[[list], str],
        batch_system: str,
        parse_batch_fn: Callable[[dict, list], list[dict]],
        desc: str = "llm batched",
        batch_size: int = 10,
    ) -> list:
        """Batched LLM tagging with per-item cache compatibility.

        cache_user_fn(item) returns the SINGLE-item user prompt; this is what's hashed
        for the cache key, so cache stays compatible with non-batched runs.
        batch_user_fn(batch) builds one combined user prompt for the batch.
        parse_batch_fn(response, batch_items) returns a list of N result dicts in order.

        Round-retry semantics: a whole failing batch is retried as a batch in the next
        round. After max_round_retries, raises.
        """
        n = len(items)
        results: list = [None] * n

        if self.cache is not None:
            hits = 0
            for i, item in enumerate(items):
                key = JsonlCache.make_key(self.model, cache_system, cache_user_fn(item))
                hit = self.cache.get(key)
                if hit is not None:
                    results[i] = hit
                    hits += 1
            print(f"  {desc}: cache hits {hits}/{n}")

        def needs_redo(idx: int) -> bool:
            r = results[idx]
            if r is None:
                return True
            if isinstance(r, dict) and r.get("_error"):
                return True
            return False

        pending = [i for i in range(n) if needs_redo(i)]
        if not pending:
            return results

        print(f"  {desc}: processing {len(pending)} items in batches of {batch_size}")

        for round_idx in range(self.max_round_retries):
            batches = [pending[s : s + batch_size] for s in range(0, len(pending), batch_size)]

            def process_batch(batch_idxs: list[int]) -> list[tuple[int, dict]]:
                batch_items = [items[i] for i in batch_idxs]
                try:
                    user_prompt = batch_user_fn(batch_items)
                    response = self.chat_json(
                        system=batch_system,
                        user=user_prompt,
                        use_cache=False,
                    )
                    batch_results = parse_batch_fn(response, batch_items)
                    if len(batch_results) != len(batch_items):
                        raise ValueError(
                            f"batch size mismatch: got {len(batch_results)} results for batch of {len(batch_items)}"
                        )
                    return list(zip(batch_idxs, batch_results))
                except Exception as e:
                    return [(idx, {"_error": str(e)}) for idx in batch_idxs]

            label = f"{desc} round {round_idx + 1}/{self.max_round_retries}"
            ok_count = 0
            err_count = 0
            consecutive_fail_batches = 0
            seen_errors: dict[str, int] = {}
            abort_reason: str | None = None
            ex = ThreadPoolExecutor(max_workers=self.concurrency)
            try:
                futures = [ex.submit(process_batch, b) for b in batches]
                with tqdm(
                    total=len(pending),
                    desc=label,
                    miniters=1,
                    mininterval=0.0,
                    smoothing=0.1,
                    file=sys.stderr,
                ) as bar, ProgressRefresher(bar, interval=0.5):
                    for fut in as_completed(futures):
                        pairs = fut.result()
                        batch_err_msg = None
                        for idx, r in pairs:
                            results[idx] = r
                            if isinstance(r, dict) and r.get("_error"):
                                err_count += 1
                                err_msg = str(r["_error"])
                                batch_err_msg = err_msg
                                if err_msg not in seen_errors:
                                    seen_errors[err_msg] = 1
                                    tqdm.write(
                                        f"[{label}] NEW ERROR ({len(seen_errors)} unique so far): {err_msg}"
                                    )
                                    sys.stderr.flush()
                                else:
                                    seen_errors[err_msg] += 1
                            else:
                                ok_count += 1
                                if self.cache is not None:
                                    key = JsonlCache.make_key(
                                        self.model, cache_system, cache_user_fn(items[idx])
                                    )
                                    self.cache.put(key, r)
                        if batch_err_msg is not None:
                            consecutive_fail_batches += 1
                            self._log_error(label, batch_err_msg)
                        else:
                            consecutive_fail_batches = 0
                        bar.update(len(pairs))
                        bar.set_postfix_str(
                            f"ok={ok_count} err={err_count} streak={consecutive_fail_batches} uniq_errs={len(seen_errors)} bsz={batch_size}"
                        )
                        bar.refresh()
                        self._hb_update(done=ok_count + err_count, ok=ok_count, err=err_count)
                        if (
                            self.fail_fast_threshold > 0
                            and consecutive_fail_batches >= self.fail_fast_threshold
                        ):
                            abort_reason = (
                                f"fail-fast: {consecutive_fail_batches} consecutive failed batches "
                                f"(threshold={self.fail_fast_threshold}). Last: {last_err}"
                            )
                            tqdm.write(f"[{label}] {abort_reason}")
                            sys.stderr.flush()
                            self.request_stop()
                            for f in list(futures):
                                f.cancel()
                            break
            except KeyboardInterrupt:
                tqdm.write(f"[{label}] KeyboardInterrupt — cancelling pending futures")
                sys.stderr.flush()
                self.request_stop()
                for f in list(futures):
                    f.cancel()
                ex.shutdown(wait=False, cancel_futures=True)
                raise
            finally:
                ex.shutdown(wait=False, cancel_futures=True)
                self.reset_stop()
            if seen_errors:
                tqdm.write(
                    f"[{label}] Round summary: ok={ok_count} err={err_count} "
                    f"unique_errors={len(seen_errors)}"
                )
                for msg, cnt in sorted(seen_errors.items(), key=lambda kv: -kv[1])[:5]:
                    tqdm.write(f"    {cnt}x  {msg}")
                if self._error_log_path is not None:
                    tqdm.write(f"  Full errors logged to: {self._error_log_path}")
                sys.stderr.flush()
            if abort_reason:
                raise RuntimeError(abort_reason)

            pending = [i for i in range(n) if needs_redo(i)]
            if not pending:
                return results

            print(
                f"  {len(pending)} items still failing after round {round_idx + 1}; "
                f"sleeping {self.round_backoff_seconds}s before retry."
            )
            time.sleep(self.round_backoff_seconds)

        return results
