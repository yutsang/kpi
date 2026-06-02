"""Diagnostic: probe the rerank endpoint with a real query and dump the raw
response so we can see if our parser matches the actual API shape.

Usage:
    python -u src/diag_rerank.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import httpx

from kpi.lib.conf import load_config  # noqa: E402
from kpi.lib.rerank import RerankClient  # noqa: E402


def main():
    cfg = load_config()
    rcfg = cfg.get("rerank") or {}
    if not rcfg.get("enabled"):
        print("ERROR: rerank.enabled is false in config.yaml")
        return

    api_base = rcfg["api_base"]
    api_key = rcfg.get("api_key") or cfg["llm"]["api_key"]
    model = rcfg.get("model", "bgem3v2_rerank")
    verify = bool(rcfg.get("verify_ssl", False))

    print(f"Endpoint: {api_base}")
    print(f"Model:    {model}\n", flush=True)

    test_cases = [
        # (query, expected_top_label_in_chinese)
        ("Comp Lodging", "酒店客房"),
        ("Comp Food", "餐飲"),
        ("Sponsorship Fee", "贊助費"),
        ("Venue License", "活動場地"),
        ("Payroll - Direct Event Investment", "人工成本"),
        ("AUC-System fit-out work", "建設與設施支出"),
        ("Marketing - Barter promotional", "其他"),
        ("Donations", "其他"),
    ]
    documents = [
        "數量/次數",
        "建設與設施支出",
        "人工成本",
        "酒店客房",
        "餐飲",
        "活動場地",
        "專業服務費",
        "贊助費",
        "其他",
    ]

    headers = {
        "authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    http = httpx.Client(verify=verify, timeout=30)

    for query, expected in test_cases:
        print(f"\n=== query: {query!r}  (expected ≈ {expected!r}) ===")
        body = {"documents": documents, "query": query, "model": model}
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        try:
            resp = http.post(api_base, headers=headers, content=data)
            print(f"  HTTP {resp.status_code}")
            payload = resp.json()
            print(f"  RAW response keys: {list(payload.keys()) if isinstance(payload, dict) else type(payload).__name__}")
            print(f"  RAW response (truncated): {json.dumps(payload, ensure_ascii=False)[:500]}")

            # try our parser
            scores = RerankClient._parse_scores(payload, len(documents))
            print(f"  parsed scores by our code: {[round(s, 3) for s in scores]}")
            top = max(range(len(scores)), key=lambda i: scores[i]) if scores else -1
            if top >= 0:
                print(f"  top: {documents[top]} = {scores[top]:.3f}")
            print(f"  zero-count: {sum(1 for s in scores if s == 0.0)}/{len(scores)}")
        except Exception as e:
            print(f"  EXCEPTION: {e}")
    http.close()


if __name__ == "__main__":
    main()
