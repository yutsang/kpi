"""Multi-process rate-limit test.

Question: if you open 6 cmd windows and run step 3 in each (one per company),
will the company LLM endpoint serve all 6 in parallel, or does its rate-limit
quota apply per API key (in which case 6 procs share 1 quota, no speedup)?

This script spawns N child processes that all hammer the same endpoint at the
same time, each independently making M requests with no client-side rate limit.
We measure how many succeed per process and total wall-clock time.

Interpretation:
  - If 1-proc throughput == N-proc per-proc throughput: NO sharing penalty
    → opening N terminals genuinely scales (multi-process bypasses rate limit).
  - If N-proc per-proc throughput == (1-proc throughput / N): proc share quota
    → multi-terminal does NOT help.
  - In between: partial sharing.

Usage:
    python -u src/diag_multi_process.py
"""
from __future__ import annotations

import json
import os
import sys
import time
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import httpx

from kpi.lib.conf import load_config  # noqa: E402

M = 20  # requests per process
SYSTEM = 'You are a JSON echo bot. Reply with exactly: {"ok": true}'
USER = "Reply now."


def worker(args) -> dict:
    proc_id, api_base, api_key, model, timeout, verify, m = args
    http = httpx.Client(verify=verify, timeout=timeout)
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
    ok = 0
    err = 0
    rate_limit = 0
    t0 = time.perf_counter()
    for i in range(m):
        try:
            resp = http.post(url, headers=headers, json=body)
            payload = resp.json()
            if isinstance(payload, dict) and payload.get("code") == 30001:
                rate_limit += 1
                err += 1
            elif payload.get("choices"):
                ok += 1
            else:
                err += 1
        except Exception:
            err += 1
    elapsed = time.perf_counter() - t0
    http.close()
    return {
        "proc": proc_id,
        "ok": ok,
        "err": err,
        "rate_limit": rate_limit,
        "elapsed": elapsed,
        "ok_per_sec": ok / elapsed if elapsed > 0 else 0,
    }


def run_n_processes(n: int, cfg: dict) -> list[dict]:
    llm = cfg["llm"]
    args = [(
        i,
        llm["api_base"],
        llm["api_key"],
        llm["chat_model"],
        float(llm.get("request_timeout", 30)),
        bool(llm.get("verify_ssl", False)),
        M,
    ) for i in range(n)]
    print(f"\n=== Spawning {n} processes × {M} reqs each ===", flush=True)
    t0 = time.perf_counter()
    with ProcessPoolExecutor(max_workers=n) as ex:
        results = list(ex.map(worker, args))
    total_elapsed = time.perf_counter() - t0
    print(f"  Total wall-clock: {total_elapsed:.2f}s", flush=True)
    total_ok = sum(r["ok"] for r in results)
    total_err = sum(r["err"] for r in results)
    total_rl = sum(r["rate_limit"] for r in results)
    for r in results:
        print(
            f"  proc[{r['proc']}]: ok={r['ok']:>3}/{M}  err={r['err']}  "
            f"30001={r['rate_limit']}  elapsed={r['elapsed']:.1f}s  "
            f"({r['ok_per_sec']:.2f} ok/sec)",
            flush=True,
        )
    print(f"  TOTAL: ok={total_ok}/{n*M}  err={total_err}  30001={total_rl}", flush=True)
    print(f"  Combined throughput: {total_ok/total_elapsed:.2f} ok/sec", flush=True)
    return results


def main():
    cfg = load_config()
    print(f"Endpoint: {cfg['llm']['api_base']}")
    print(f"Model:    {cfg['llm']['chat_model']}\n", flush=True)

    summary = []
    for n in (1, 2, 4, 6):
        results = run_n_processes(n, cfg)
        total_ok = sum(r["ok"] for r in results)
        total_elapsed = max(r["elapsed"] for r in results)
        ok_per_sec = total_ok / total_elapsed if total_elapsed > 0 else 0
        summary.append({"n_proc": n, "total_ok": total_ok, "total_err": sum(r["err"] for r in results), "ok_per_sec": ok_per_sec})
        print(f"\n  Cool-down 10s between rounds...", flush=True)
        time.sleep(10)

    print("\n\n========== SUMMARY ==========")
    print(f"{'n_proc':>6}  {'total_ok':>8}  {'total_err':>9}  {'ok_per_sec':>10}  notes")
    print("-" * 70)
    base = next((s for s in summary if s["n_proc"] == 1), None)
    base_rate = base["ok_per_sec"] if base else 1.0
    for s in summary:
        ratio = s["ok_per_sec"] / base_rate if base_rate else 0
        note = ""
        if s["n_proc"] == 1:
            note = "baseline"
        elif ratio >= 0.9 * s["n_proc"]:
            note = "FULL SCALE — multi-proc bypasses rate limit"
        elif ratio >= 1.5:
            note = "partial scale"
        else:
            note = "no scale — proc share quota"
        print(f"{s['n_proc']:>6}  {s['total_ok']:>8}  {s['total_err']:>9}  {s['ok_per_sec']:>10.2f}  {note}")

    print("\nInterpretation:")
    print("  - If '1-proc' rate ≈ '6-proc' rate (per process):   key shares quota → no benefit from multi-cmd")
    print("  - If '6-proc' total_ok_per_sec ≈ 6× '1-proc':       quota is per-process/per-connection → MAJOR speedup")


if __name__ == "__main__":
    main()
