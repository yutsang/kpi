"""Diagnostic: time 10 LLM calls sequentially vs in parallel against your endpoint.

If parallel takes ~10x faster than sequential, concurrency works.
If parallel takes the same time as sequential, the endpoint is rate-limiting per key.

Usage:
    python src/diag_concurrency.py
"""
from __future__ import annotations

import sys
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

from kpi.lib.conf import load_config  # noqa: E402
from kpi.lib.llm import LLMClient  # noqa: E402

N = 10


def main():
    cfg = load_config()
    llm = LLMClient(cfg, cache_path=None)

    system = "You are a JSON echo bot. Reply with exactly: {\"ok\": true}"
    user = "Reply now."

    print(f"\n=== Sequential: {N} calls one-by-one ===")
    t0 = time.perf_counter()
    for i in range(N):
        llm.chat_json(system=system, user=user)
        print(f"  call {i+1}/{N} done  elapsed={time.perf_counter()-t0:.1f}s", flush=True)
    seq = time.perf_counter() - t0
    print(f"Sequential total: {seq:.2f}s  ({seq/N:.2f}s per call)")

    print(f"\n=== Parallel: {N} calls with {N} threads ===")
    t0 = time.perf_counter()
    with ThreadPoolExecutor(max_workers=N) as ex:
        futures = [ex.submit(llm.chat_json, system, user) for _ in range(N)]
        for i, f in enumerate(futures, 1):
            f.result()
            print(f"  call {i}/{N} done  elapsed={time.perf_counter()-t0:.1f}s", flush=True)
    par = time.perf_counter() - t0
    print(f"Parallel total:   {par:.2f}s  ({par/N:.2f}s per call)")

    speedup = seq / par if par > 0 else 0
    print(f"\n=== Speedup: {speedup:.2f}x ===")
    if speedup > 5:
        print("Concurrency works well. Push concurrency in step 3 to N+ threads.")
    elif speedup > 2:
        print("Partial concurrency. Endpoint allows some parallelism but not full.")
    else:
        print("Endpoint serializes per API key. Increase batch_size, decrease concurrency to 1.")


if __name__ == "__main__":
    main()
