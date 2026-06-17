r"""Dump melco 仲 no-H 嘅行，clean 走 PO號/日期/數字後 group line_memo，等 Claude 逐個 type 分類。

冇 項目組 reference → 用呢個 build：剩低 no-H 全部係 ledger_account='0' 嘅 capex PO 長尾。
clean line_memo（去 PO\d+/PP\d+/LI.../日期/%/數字）→ group → 按 |Σ| 排 top ~150。
Claude 睇晒 pattern (UPS/Cabling/LED/Aristocrat/Game License/Kitchen/Air Con…) → 寫返 keyword batch。

Run:  python scripts/dump_melco_noH.py
出 results/dump_melco_noH.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
P = ROOT / "data" / "melco" / "output" / "company_5_kpi_report.parquet"
AMT = ["Amount - Amended", "ledger_amount", "amount_mop"]


def _clean(s):
    s = str(s)
    s = re.sub(r"\b(PO|PP|LI|GE|CE|MR|FA|RCL|IntmptyNo|AP|C|A|MSC|SA|WP|CoD|GLORY)[\d/\.\-A-Za-z]*\d[\d/\.\-A-Za-z]*", " ", s)
    s = re.sub(r"Dep\s*\d+%|\d+\s*%|\d+EA|\d+x|\bx\d+\b", " ", s)
    s = re.sub(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s*\d{0,4}", " ", s)
    s = re.sub(r"20\d\d|\b\d{3,}\b", " ", s)
    s = re.sub(r"[/|\\()（）]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:60] if s else "(空)"


def main():
    L = ["# dump_melco_noH — no-H 行 clean line_memo group (Claude 分類 capex PO 長尾)"]
    if not P.exists():
        L.append("!! melco parquet 揾唔到"); _w(L); return
    df = pd.read_parquet(P)
    amtc = next((c for c in AMT if c in df.columns), None)
    amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4
    Hid = df["horizontal_id"].astype(str).fillna("") if "horizontal_id" in df.columns else pd.Series("", index=df.index)
    lm = df["line_memo"].astype(str).fillna("") if "line_memo" in df.columns else pd.Series("", index=df.index)
    rp = df["report_period"].astype(str) if "report_period" in df.columns else pd.Series("", index=df.index)
    noH = Hid.isin(["", "nan", "None"])
    L.append(f"   no-H: {int(noH.sum()):,} 行  Σ={amt[noH].sum():,.0f}萬")

    clean = lm[noH].map(_clean)
    g = pd.DataFrame({"k": clean.values, "a": amt[noH].values, "b": rp[noH].values})
    agg = g.groupby("k").agg(Σ=("a", "sum"), n=("a", "size"),
                             bkt=("b", lambda s: ",".join(sorted(set(s))[:3]))).reset_index()
    agg = agg.reindex(agg["Σ"].abs().sort_values(ascending=False).index)
    cov = agg.head(150)["Σ"].abs().sum() / max(agg["Σ"].abs().sum(), 1) * 100
    L.append(f"   group 後 {len(agg):,} 個 distinct pattern；頭 150 覆蓋 {cov:.0f}% 金額")
    L.append("   ── clean line_memo | Σ萬 | 行數 | bucket ──")
    for _, r in agg.head(150).iterrows():
        L.append(f"      {str(r['k'])[:58]:<58} {r['Σ']:>9,.0f}萬  {int(r['n']):>4}行  {r['bkt']}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "dump_melco_noH.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
