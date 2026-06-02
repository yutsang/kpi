"""Diagnose results/mgm_golden_25.tsv — why build_mgm_golden output looks 太少.

Shows (a) the RAW layout (line count, per-line field count, first lines) so we can see if the
separator / column order differs from what build_mgm_golden.parse_golden() expects
(序號⇥名稱⇥payroll⇥capex⇥opex⇥total, amounts in 萬元), and (b) what the parser actually
extracts (N 項目 + totals + first rows). If (a) looks fine but (b) is tiny → parser mismatch.

Run on Windows:
  python scripts/inspect_mgm_golden.py
  python scripts/inspect_mgm_golden.py --golden results/mgm_golden_25.tsv
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass

ROOT = Path(__file__).resolve().parent.parent


def _num(x):
    x = str(x).replace(",", "").strip()
    try:
        return float(x) * 10000.0
    except Exception:
        return 0.0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--golden", default="results/mgm_golden_25.tsv")
    args = ap.parse_args()
    p = ROOT / args.golden
    if not p.exists():
        print(f"X {p} MISSING — build_mgm_golden 會 read 唔到 → output 空。檢查檔名/位置。")
        # list what IS in results/
        rd = ROOT / "results"
        if rd.exists():
            print("results/ 入面有:")
            for f in sorted(rd.glob("*mgm*")) + sorted(rd.glob("*golden*")):
                print(f"   {f.name}")
        return

    lines = p.read_text(encoding="utf-8-sig").splitlines()
    print(f"file: {p}")
    print(f"total lines: {len(lines)}")

    # separator sniff on the first 30 non-empty lines
    tabs = sum(l.count("\t") for l in lines[:30])
    commas = sum(l.count(",") for l in lines[:30])
    print(f"separator sniff (first 30 lines): tabs={tabs}  commas={commas}  "
          f"→ {'TAB' if tabs >= commas else 'COMMA?? (parser expects TAB!)'}")

    print("\n=== first 12 raw lines (│=tab boundary) ===")
    for i, l in enumerate(lines[:12]):
        parts = l.split("\t")
        digit0 = bool(re.fullmatch(r"\d+", parts[0].strip())) if parts else False
        print(f"  L{i:<2} fields={len(parts):<2} 序號?={'Y' if digit0 else '.'}  "
              f"{'│'.join(p.strip()[:14] for p in parts)[:110]}")

    # replicate parse_golden
    out = {}
    for l in lines:
        parts = [c.strip() for c in l.split("\t")]
        if len(parts) < 6:
            continue
        sn = parts[0].strip()
        if not re.fullmatch(r"\d+", sn):
            continue
        out[sn.zfill(3)] = {"name": parts[1], "payroll": _num(parts[2]),
                            "capex": _num(parts[3]), "opex": _num(parts[4]), "total": _num(parts[5])}
    cap = sum(g["capex"] for g in out.values())
    opx = sum(g["opex"] for g in out.values())
    pay = sum(g["payroll"] for g in out.values())
    print(f"\n=== parse_golden() result ===")
    print(f"  項目 parsed: {len(out)}")
    print(f"  Σcapex={cap:,.0f}  Σopex={opx:,.0f}  Σpayroll={pay:,.0f}  Σtotal={cap+opx+pay:,.0f}")
    print(f"  (期望: ~114 項目, total ~2,086,000,000)")
    print(f"\n  first 6 parsed 項目 (capex/opex/payroll in MOP):")
    for sn, g in list(out.items())[:6]:
        print(f"    {sn}  cap={g['capex']:>14,.0f}  opx={g['opex']:>13,.0f}  pay={g['payroll']:>12,.0f}  {g['name'][:24]}")
    if len(out) < 50:
        print("\n  ⚠ parsed < 50 項目 → layout 唔啱 parser。睇返上面 raw lines 嘅 fields= 同 序號?，"
              "話我知邊 column 係 序號/payroll/capex/opex/total，我改 parse_golden。")


if __name__ == "__main__":
    main()
