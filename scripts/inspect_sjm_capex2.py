"""inspect_sjm_capex2.py — sjm raw 全欄 + 2026年1-4月 Y rows 診斷
Run: python scripts\inspect_sjm_capex2.py
Out: results\inspect_sjm_capex2.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F    = ROOT / "data" / "sjm" / "raw" / "sjm_2025.xlsx"
OUT  = ROOT / "results" / "inspect_sjm_capex2.txt"


def main():
    L = ["# inspect_sjm_capex2 — SJM raw 全欄 + 2026年1-4月 Y rows", ""]
    if not F.exists():
        L.append(f"!! {F} 揾唔到"); _w(L); return

    rdf = pd.read_excel(F, sheet_name="data", dtype=str)
    rdf.columns = [str(c).strip() for c in rdf.columns]
    n = len(rdf)
    L.append(f"Loaded: {n:,} rows, {len(rdf.columns)} cols")

    # ── 1. ALL columns ──────────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 1. ALL columns"]
    for i, c in enumerate(rdf.columns):
        L.append(f"  [{i:>3}] {c}")

    # ── 2. All numeric columns (Σ) ──────────────────────────────────────────────
    L += ["", "=" * 70, "## 2. All numeric columns (Σ ≠ 0)"]
    L.append(f"  {'col':<55}  {'Σ':>20}  {'non-zero':>10}")
    for c in rdf.columns:
        num = pd.to_numeric(rdf[c].astype(str).str.replace(",","",regex=False), errors="coerce")
        if num.notna().sum() > n * 0.05 and abs(num.sum()) > 0:
            L.append(f"  {str(c):<55}  {num.sum():>20,.0f}  {int((num.fillna(0)!=0).sum()):>10,}")

    # ── 3. Capex rows: all numeric cols Σ ───────────────────────────────────────
    L += ["", "=" * 70, "## 3. 報告capex/opex=capex rows: numeric col Σ"]
    cap_mask = rdf.get("報告capex/opex", pd.Series("", index=rdf.index)).astype(str).str.strip().str.lower().eq("capex")
    L.append(f"  Capex rows: {int(cap_mask.sum()):,}/{n:,}")
    L.append(f"  {'col':<55}  {'Σ capex':>20}")
    for c in rdf.columns:
        num = pd.to_numeric(rdf[c].astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0)
        s = num[cap_mask].sum()
        if abs(s) > 0:
            L.append(f"  {str(c):<55}  {s:>20,.0f}")

    # ── 4. 2026年1-4月 Y rows ────────────────────────────────────────────────────
    L += ["", "=" * 70, "## 4. 2026年1-4月入賬 Y rows (全欄)"]
    ycol = None
    for c in rdf.columns:
        if "2026" in str(c) and ("1-4" in str(c) or "一" in str(c) or "月" in str(c)):
            ycol = c
            break
    if ycol is None:
        L.append("  !! 找不到 2026年1-4月欄，搜索所有包含 '2026' 嘅欄：")
        L.append("  " + str([c for c in rdf.columns if "2026" in str(c)]))
    else:
        L.append(f"  Column: [{ycol}]")
        y_mask = rdf[ycol].astype(str).str.strip().str.upper().eq("Y")
        L.append(f"  Y rows: {int(y_mask.sum()):,}")
        if y_mask.any():
            sub = rdf[y_mask].copy()
            L.append(f"\n  All {int(y_mask.sum())} Y rows — numeric col amounts:")
            # Show all numeric cols for Y rows
            L.append(f"  {'col':<55}  {'Σ Y rows':>20}  {'#non-zero':>10}")
            for c in rdf.columns:
                num = pd.to_numeric(sub[c].astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0)
                s = num.sum()
                nz = int((num != 0).sum())
                if abs(s) > 0 or nz > 0:
                    L.append(f"  {str(c):<55}  {s:>20,.0f}  {nz:>10,}")
            L.append("")
            # Show each Y row
            L.append("  Individual Y rows:")
            amt_cols = [c for c in rdf.columns if any(k in str(c) for k in ["跨年","調整","Val","金額"])]
            hdr = "  " + " | ".join(f"{str(c)[:20]:<20}" for c in ["報告capex/opex","項目期間"] + amt_cols)
            L.append(hdr)
            for _, row in sub.iterrows():
                vals = " | ".join(f"{str(row[c])[:20]:<20}" for c in ["報告capex/opex","項目期間"] + amt_cols)
                L.append("  " + vals)

    # ── 5. 各 bucket 欄 Σ for capex rows ────────────────────────────────────────
    L += ["", "=" * 70, "## 5. Per-bucket amount cols Σ for capex rows"]
    bucket_cols = {}
    for c in rdf.columns:
        cs = str(c)
        if any(k in cs for k in ["跨年", "調整金額", "Val/COArea"]):
            bucket_cols[cs] = c
    L.append(f"  {'col':<40}  {'Σ capex':>15}  {'Σ all':>15}")
    for cs, c in bucket_cols.items():
        num = pd.to_numeric(rdf[c].astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0)
        L.append(f"  {cs:<40}  {num[cap_mask].sum():>15,.0f}  {num.sum():>15,.0f}")

    # ── 6. Golden reconciliation: sum all bucket cols for capex ─────────────────
    L += ["", "=" * 70, "## 6. Attempt golden recon: capex rows, all bucket cols"]
    # Try: 25跨年+25調整金額 + 24跨年+24調整金額 + 23跨年+23調整金額
    combos = [
        ("25跨年", "25調整金額"),
        ("24跨年", "24調整金額"),
        ("23跨年", "23調整金額"),
    ]
    total = 0
    for amt_c, adj_c in combos:
        amt = pd.to_numeric(rdf.get(amt_c, 0), errors="coerce").fillna(0) if amt_c in rdf.columns else pd.Series(0.0, index=rdf.index)
        adj = pd.to_numeric(rdf.get(adj_c, 0), errors="coerce").fillna(0) if adj_c in rdf.columns else pd.Series(0.0, index=rdf.index)
        net_cap = (amt + adj)[cap_mask].sum()
        total += net_cap
        L.append(f"  {amt_c}+{adj_c} (capex rows): {net_cap/1e4:,.0f}萬")
    L.append(f"  TOTAL capex (3 buckets): {total/1e4:,.0f}萬  (golden=117,736  Δ={(total/1e4-117736):+,.0f})")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
