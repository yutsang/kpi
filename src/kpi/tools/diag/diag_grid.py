"""Grid search over (concurrency, rate_limit) combos to find the best production config.

For each (threads, rate) combo, fires N calls and reports:
  - ok / N (success rate)
  - total wall time
  - per-call average
  - speedup vs sequential baseline

Sleeps between combos to let server quota recover.

Usage:
    python -u src/diag_grid.py
"""
from __future__ import annotations

import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import httpx

from kpi.lib.conf import load_config  # noqa: E402
from kpi.lib.rate_limiter import RateLimiter  # noqa: E402

N = 10                  # calls per combo
COOLDOWN_SECONDS = 8    # rest between combos so server quota recovers
SYSTEM = 'You are a JSON echo bot. Reply with exactly: {"ok": true}'
USER = "Reply now."

# (concurrency, rate_per_sec)  — rate=0 means no rate limit
GRID = [
    (1, 0),    # sequential baseline
    (2, 2),
    (2, 3),
    (2, 4),
    (4, 2),
    (4, 3),
    (4, 4),
    (4, 5),
    (4, 6),
    (8, 3),
    (8, 4),
    (8, 5),
    (8, 6),
    (8, 8),
    (16, 4),
    (16, 6),
    (16, 8),
    (16, 10),
]


def call_via_raw_httpx(http: httpx.Client, url: str, headers: dict, body: dict) -> None:
    resp = http.post(url, headers=headers, json=body)
    resp.raise_for_status()
    payload = resp.json()
    if not isinstance(payload, dict):
        raise RuntimeError(f"non-dict response: {str(payload)[:200]}")
    if payload.get("code") == 30001 or payload.get("flag") is False:
        raise RuntimeError(f"server rate limit: {payload.get('msg', '?')}")
    choices = payload.get("choices")
    if not choices:
        raise RuntimeError(f"no choices: {str(payload)[:200]}")
    content = (choices[0] or {}).get("message", {}).get("content") or ""
    if not str(content).strip():
        raise RuntimeError("empty content")


def run_combo(http: httpx.Client, url: str, headers: dict, body: dict, concurrency: int, rate: float, label: str) -> dict:
    limiter = RateLimiter(rate) if rate > 0 else None
    ok = 0
    err = 0
    first_err = None

    def task():
        if limiter is not None:
            limiter.acquire()
        call_via_raw_httpx(http, url, headers, body)

    t0 = time.perf_counter()
    with ThreadPoolExecutor(max_workers=concurrency) as ex:
        futures = [ex.submit(task) for _ in range(N)]
        for i, f in enumerate(futures, 1):
            try:
                f.result()
                ok += 1
            except Exception as e:
                err += 1
                if first_err is None:
                    first_err = str(e)
    elapsed = time.perf_counter() - t0
    print(
        f"  {label}: {elapsed:>5.2f}s  per_call={elapsed/N:>5.2f}s  ok={ok:>2}/{N}  err={err}"
        + (f"  (first_err: {first_err[:60]})" if first_err else ""),
        flush=True,
    )
    return {
        "concurrency": concurrency,
        "rate": rate,
        "elapsed": elapsed,
        "ok": ok,
        "err": err,
        "first_err": first_err,
    }


def main():
    cfg = load_config()
    llm = cfg["llm"]
    api_base = str(llm["api_base"]).rstrip("/")
    api_key = llm["api_key"]
    model = llm["chat_model"]
    timeout = float(llm.get("request_timeout", 30))
    verify = bool(llm.get("verify_ssl", False))

    url = api_base + "/chat/completions"
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

    print(f"Endpoint: {api_base}")
    print(f"Model:    {model}")
    print(f"Grid: {len(GRID)} combos × N={N} calls each, {COOLDOWN_SECONDS}s cooldown between\n", flush=True)

    http = httpx.Client(verify=verify, timeout=timeout)
    results = []
    for idx, (conc, rate) in enumerate(GRID, 1):
        rate_label = f"{rate} req/s" if rate > 0 else "no limit"
        label = f"[{idx:>2}/{len(GRID)}] conc={conc:>2}  rate={rate_label}"
        print(label, flush=True)
        r = run_combo(http, url, headers, body, conc, rate, label="    ")
        results.append(r)
        if idx < len(GRID):
            time.sleep(COOLDOWN_SECONDS)
    http.close()

    # ============ summary table ============
    baseline = next((r for r in results if r["concurrency"] == 1 and r["ok"] == N), None)
    base_t = baseline["elapsed"] if baseline else None

    print("\n\n========== GRID SUMMARY ==========")
    print(f"Sequential baseline: {base_t:.2f}s for {N} calls" if base_t else "Sequential baseline failed!")
    print()
    print(f"{'conc':>4}  {'rate':>10}  {'total':>7}  {'per_call':>9}  {'ok':>5}  {'speedup':>8}  {'safe':>4}")
    print("-" * 70)
    for r in results:
        rate_label = f"{r['rate']:.1f} req/s" if r['rate'] > 0 else "no limit"
        ok_str = f"{r['ok']}/{N}"
        speedup_str = f"{base_t / r['elapsed']:.2f}x" if base_t else "—"
        safe = "YES" if r["ok"] == N else "no"
        print(
            f"{r['concurrency']:>4}  {rate_label:>10}  {r['elapsed']:>6.2f}s  "
            f"{r['elapsed']/N:>8.2f}s  {ok_str:>5}  {speedup_str:>8}  {safe:>4}"
        )

    # ============ recommendations ============
    safe = [r for r in results if r["ok"] == N and r["concurrency"] > 1]
    if safe and base_t:
        fastest_safe = min(safe, key=lambda r: r["elapsed"])
        print(
            f"\n>>> Fastest 10/10 SAFE combo: concurrency={fastest_safe['concurrency']}, "
            f"rate={fastest_safe['rate']:.1f} req/s — "
            f"{base_t / fastest_safe['elapsed']:.2f}x faster than sequential."
        )
        print(f"    Recommended config:")
        print(f"      concurrency: {fastest_safe['concurrency']}")
        print(f"      max_requests_per_second: {fastest_safe['rate']}")
    else:
        print("\n>>> No combo achieved 10/10. Server quota highly variable; rely on retry+cache.")
        # Find the fastest with ≥80% success rate
        good = [r for r in results if r["ok"] >= int(N * 0.8) and r["concurrency"] > 1]
        if good and base_t:
            fastest = min(good, key=lambda r: r["elapsed"])
            print(
                f"    Fastest with ≥80% success: concurrency={fastest['concurrency']}, "
                f"rate={fastest['rate']:.1f} req/s — {fastest['ok']}/{N} ok, "
                f"{base_t / fastest['elapsed']:.2f}x speedup."
            )


if __name__ == "__main__":
    main()
