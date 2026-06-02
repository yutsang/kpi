"""Full concurrency diagnostic — tries 7 different strategies to see which (if any)
gives speedup on your company LLM endpoint.

Each strategy runs N=10 identical "echo JSON" calls and reports:
  - total wall-clock time
  - per-call average
  - speedup vs sequential baseline

Strategies:
  A. Sequential baseline (concurrency=1)
  B. ThreadPoolExecutor + SHARED OpenAI client      (current production setup)
  C. ThreadPoolExecutor + per-thread httpx.Client   (separate connection pools)
  D. ThreadPoolExecutor + per-thread OpenAI client  (separate clients fully)
  E. ThreadPoolExecutor + raw httpx (no OpenAI SDK) (tests if SDK is the bottleneck)
  F. ThreadPoolExecutor + Connection: close header  (force new TCP per call)
  G. asyncio + httpx.AsyncClient                    (different concurrency model)

Usage:
    python -u src/diag_concurrency_full.py

If your LLM endpoint allows real concurrency, B/C/D/E should give ~Nx speedup.
If everything matches sequential timing, the endpoint serializes per API key.
If only G (asyncio) works, the bottleneck is in threading, not the endpoint.
"""
from __future__ import annotations

import asyncio
import json
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import httpx
from openai import OpenAI

from kpi.lib.conf import load_config  # noqa: E402
from kpi.lib.rate_limiter import RateLimiter  # noqa: E402

N = 10  # calls per strategy
SYSTEM = 'You are a JSON echo bot. Reply with exactly: {"ok": true}'
USER = "Reply now."


def make_openai_client(api_base: str, api_key: str, timeout: float, verify: bool):
    http = httpx.Client(verify=verify, timeout=timeout)
    return OpenAI(base_url=api_base, api_key=api_key, http_client=http)


def _safe_extract_openai(resp) -> str:
    if resp is None:
        raise RuntimeError("openai resp is None")
    choices = getattr(resp, "choices", None)
    if not choices:
        err = getattr(resp, "error", None)
        try:
            dump = resp.model_dump() if hasattr(resp, "model_dump") else repr(resp)
        except Exception:
            dump = repr(resp)
        raise RuntimeError(f"no choices (error={err!r}, dump={str(dump)[:300]})")
    msg = getattr(choices[0], "message", None)
    if msg is None:
        raise RuntimeError(f"choice without message: {repr(choices[0])[:200]}")
    content = getattr(msg, "content", None) or ""
    if not content.strip():
        raise RuntimeError("empty content")
    return content


def _safe_extract_httpx(resp) -> str:
    payload = resp.json()
    choices = payload.get("choices")
    if not choices:
        err = payload.get("error")
        raise RuntimeError(f"no choices (error={err!r}, raw={str(payload)[:300]})")
    msg = (choices[0] or {}).get("message") or {}
    content = msg.get("content") or ""
    if not content.strip():
        raise RuntimeError(f"empty content: {str(payload)[:300]}")
    return content


def call_via_openai(client: OpenAI, model: str) -> None:
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role": "system", "content": SYSTEM}, {"role": "user", "content": USER}],
        temperature=0,
        response_format={"type": "json_object"},
    )
    _ = _safe_extract_openai(resp)


def call_via_raw_httpx(httpx_client: httpx.Client, api_base: str, api_key: str, model: str, extra_headers: dict | None = None) -> None:
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    if extra_headers:
        headers.update(extra_headers)
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    url = api_base.rstrip("/") + "/chat/completions"
    resp = httpx_client.post(url, headers=headers, json=body)
    resp.raise_for_status()
    _ = _safe_extract_httpx(resp)


async def call_via_async_httpx(client: httpx.AsyncClient, api_base: str, api_key: str, model: str) -> None:
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": USER},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    url = api_base.rstrip("/") + "/chat/completions"
    resp = await client.post(url, headers=headers, json=body)
    resp.raise_for_status()
    _ = _safe_extract_httpx(resp)


def time_strategy(name: str, runner) -> dict:
    print(f"\n=== {name} ===", flush=True)
    t0 = time.perf_counter()
    stats = runner()  # runner returns dict {ok: int, err: int, first_err: str|None}
    elapsed = time.perf_counter() - t0
    ok = stats.get("ok", 0)
    err = stats.get("err", 0)
    first_err = stats.get("first_err")
    print(
        f"{name}: total={elapsed:.2f}s  per_call={elapsed/N:.2f}s  ok={ok}/{N} err={err}/{N}",
        flush=True,
    )
    if first_err:
        print(f"  first_error: {first_err[:200]}", flush=True)
    return {"name": name, "elapsed": elapsed, "ok": ok, "err": err, "first_err": first_err}


def main():
    cfg = load_config()
    llm = cfg["llm"]
    api_base = llm["api_base"]
    api_key = llm["api_key"]
    model = llm["chat_model"]
    timeout = float(llm.get("request_timeout", 30))
    verify = bool(llm.get("verify_ssl", False))

    print(f"Endpoint: {api_base}")
    print(f"Model:    {model}")
    print(f"N calls per strategy: {N}\n", flush=True)

    def collect_threaded(label: str, task_fn) -> dict:
        ok = 0
        err = 0
        first_err = None
        with ThreadPoolExecutor(max_workers=N) as ex:
            futures = [ex.submit(task_fn) for _ in range(N)]
            for i, f in enumerate(futures, 1):
                try:
                    f.result()
                    ok += 1
                    print(f"  {label}.{i}/{N} ok", flush=True)
                except Exception as e:
                    err += 1
                    if first_err is None:
                        first_err = str(e)
                    print(f"  {label}.{i}/{N} ERR  {str(e)[:100]}", flush=True)
        return {"ok": ok, "err": err, "first_err": first_err}

    summary = []

    # ======================================================================
    # A. Sequential baseline
    # ======================================================================
    def runner_a() -> dict:
        client = make_openai_client(api_base, api_key, timeout, verify)
        ok = 0
        err = 0
        first_err = None
        for i in range(N):
            try:
                call_via_openai(client, model)
                ok += 1
                print(f"  A.{i+1}/{N} ok", flush=True)
            except Exception as e:
                err += 1
                if first_err is None:
                    first_err = str(e)
                print(f"  A.{i+1}/{N} ERR  {str(e)[:100]}", flush=True)
        return {"ok": ok, "err": err, "first_err": first_err}
    summary.append(time_strategy("A. Sequential (concurrency=1)", runner_a))

    # ======================================================================
    # B. ThreadPool + shared OpenAI client (production setup)
    # ======================================================================
    def runner_b() -> dict:
        client = make_openai_client(api_base, api_key, timeout, verify)
        return collect_threaded("B", lambda: call_via_openai(client, model))
    summary.append(time_strategy("B. ThreadPool + shared OpenAI client", runner_b))

    # ======================================================================
    # C. ThreadPool + per-thread httpx.Client
    # ======================================================================
    def runner_c() -> dict:
        def task():
            http = httpx.Client(verify=verify, timeout=timeout)
            try:
                client = OpenAI(base_url=api_base, api_key=api_key, http_client=http)
                call_via_openai(client, model)
            finally:
                http.close()
        return collect_threaded("C", task)
    summary.append(time_strategy("C. ThreadPool + per-thread httpx.Client", runner_c))

    # ======================================================================
    # D. ThreadPool + per-thread fully fresh OpenAI client
    # ======================================================================
    def runner_d() -> dict:
        def task():
            client = make_openai_client(api_base, api_key, timeout, verify)
            call_via_openai(client, model)
        return collect_threaded("D", task)
    summary.append(time_strategy("D. ThreadPool + recreate OpenAI client per call", runner_d))

    # ======================================================================
    # E. ThreadPool + raw httpx (no OpenAI SDK)
    # ======================================================================
    def runner_e() -> dict:
        http = httpx.Client(verify=verify, timeout=timeout)
        try:
            return collect_threaded("E", lambda: call_via_raw_httpx(http, api_base, api_key, model))
        finally:
            http.close()
    summary.append(time_strategy("E. ThreadPool + raw httpx (no OpenAI SDK)", runner_e))

    # ======================================================================
    # F. ThreadPool + Connection: close (force new TCP per call)
    # ======================================================================
    def runner_f() -> dict:
        http = httpx.Client(verify=verify, timeout=timeout)
        try:
            return collect_threaded(
                "F",
                lambda: call_via_raw_httpx(http, api_base, api_key, model, {"Connection": "close"}),
            )
        finally:
            http.close()
    summary.append(time_strategy("F. ThreadPool + Connection:close (force new TCP)", runner_f))

    # ======================================================================
    # G. asyncio + httpx.AsyncClient
    # ======================================================================
    async def runner_g_async() -> dict:
        ok = 0
        err = 0
        first_err = None
        async with httpx.AsyncClient(verify=verify, timeout=timeout) as client:
            tasks = [call_via_async_httpx(client, api_base, api_key, model) for _ in range(N)]
            for i, coro in enumerate(asyncio.as_completed(tasks), 1):
                try:
                    await coro
                    ok += 1
                    print(f"  G.{i}/{N} ok", flush=True)
                except Exception as e:
                    err += 1
                    if first_err is None:
                        first_err = str(e)
                    print(f"  G.{i}/{N} ERR  {str(e)[:100]}", flush=True)
        return {"ok": ok, "err": err, "first_err": first_err}

    def runner_g() -> dict:
        return asyncio.run(runner_g_async())
    summary.append(time_strategy("G. asyncio + httpx.AsyncClient", runner_g))

    # ======================================================================
    # H/I/J. Rate-limited concurrent strategies — should all be 10/10 if rate is right
    # ======================================================================
    def make_rate_limited_runner(rate: float, label: str):
        def runner() -> dict:
            limiter = RateLimiter(rate)
            http = httpx.Client(verify=verify, timeout=timeout)
            def task():
                limiter.acquire()
                call_via_raw_httpx(http, api_base, api_key, model)
            try:
                return collect_threaded(label, task)
            finally:
                http.close()
        return runner

    summary.append(time_strategy("H. ThreadPool(8) + raw httpx + RateLimiter 3 req/s",
                                 make_rate_limited_runner(3.0, "H")))
    summary.append(time_strategy("I. ThreadPool(8) + raw httpx + RateLimiter 4 req/s",
                                 make_rate_limited_runner(4.0, "I")))
    summary.append(time_strategy("J. ThreadPool(8) + raw httpx + RateLimiter 5 req/s",
                                 make_rate_limited_runner(5.0, "J")))

    # ======================================================================
    # K/L/M. Conservative tiers — should be 10/10 ok (mimics how prod LLMClient
    # retries handle the occasional 30001, but here we test without retry to
    # find a config that's stable end-to-end without ever hitting limit).
    # ======================================================================
    def make_rate_limited_concurrency_runner(rate: float, conc: int, label: str):
        def runner() -> dict:
            limiter = RateLimiter(rate)
            http = httpx.Client(verify=verify, timeout=timeout)
            ok = 0
            err = 0
            first_err = None
            def task():
                limiter.acquire()
                call_via_raw_httpx(http, api_base, api_key, model)
            try:
                with ThreadPoolExecutor(max_workers=conc) as ex:
                    futures = [ex.submit(task) for _ in range(N)]
                    for i, f in enumerate(futures, 1):
                        try:
                            f.result()
                            ok += 1
                            print(f"  {label}.{i}/{N} ok", flush=True)
                        except Exception as e:
                            err += 1
                            if first_err is None:
                                first_err = str(e)
                            print(f"  {label}.{i}/{N} ERR  {str(e)[:100]}", flush=True)
                return {"ok": ok, "err": err, "first_err": first_err}
            finally:
                http.close()
        return runner

    summary.append(time_strategy("K. ThreadPool(2) + RateLimiter 1.5 req/s",
                                 make_rate_limited_concurrency_runner(1.5, 2, "K")))
    summary.append(time_strategy("L. ThreadPool(4) + RateLimiter 2 req/s",
                                 make_rate_limited_concurrency_runner(2.0, 4, "L")))
    summary.append(time_strategy("M. ThreadPool(2) + RateLimiter 2 req/s",
                                 make_rate_limited_concurrency_runner(2.0, 2, "M")))

    # ======================================================================
    # SUMMARY
    # ======================================================================
    print("\n\n========== SUMMARY ==========")
    baseline = next((r for r in summary if r["name"].startswith("A.") and r["ok"] == N), None)
    base_t = baseline["elapsed"] if baseline else None
    print(f"{'Strategy':<55}  {'Total':>8}  {'Per call':>10}  {'OK':>6}  {'ERR':>4}  {'Speedup':>8}")
    print("-" * 110)
    for r in summary:
        speedup_str = "—"
        if base_t and r["ok"] == N:
            speedup_str = f"{base_t / r['elapsed']:.2f}x"
        print(
            f"{r['name']:<55}  {r['elapsed']:>7.2f}s  {r['elapsed']/N:>9.2f}s  "
            f"{r['ok']:>4}/{N}  {r['err']:>4}  {speedup_str:>8}"
        )
        if r.get("first_err"):
            print(f"    └ first_err: {r['first_err'][:120]}")

    print("\nInterpretation:")
    print("  - A vs B-G: confirms whether endpoint serializes or rate-limits")
    print("  - H/I/J (rate-limited): the highest one with 10/10 ok is your safe production setting")
    print("  - Speedup should be roughly: rate_limit / sequential_rate_per_sec")
    print("    (e.g., if A is 1 req/s and I gives 4 req/s with 10/10 ok → 4x speedup)")


if __name__ == "__main__":
    main()
