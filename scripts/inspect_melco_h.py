r"""melco H 專查：點解 24 有 ~19億 冇 H。

melco H 係 **account_code driven**（predominant_rules 全部 account_code_equals/prefix）。
呢個 script 查 no-H 行：
  1. account_code (ledger_account_id) 空唔空？by bucket（證實 24 = blank-account block）
  2. no-H 行用咩 project-team 欄有值（支出性質 / 費用大類 / ledger_account 名）→ 用嚟砌 H map
  3. dump no-H top (account_desc × 支出性質) combo by Σ，等我直接寫 H 規則

Run:  python scripts/inspect_melco_h.py
出 results/inspect_melco_h.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
P = ROOT / "data" / "melco" / "output" / "company_5_kpi_report.parquet"
AMT = ["Amount - Amended", "ledger_amount", "本位幣金額", "amount_mop"]
ACODE = ["ledger_account_id", "account_code", "Account", "WD Ledger Account", "Ledger Account"]
ADESC = ["ledger_account", "account_desc", "支出性質"]
SPEND = ["支出性質", "支出性质-mapping", "支出性质 - 一级", "支出性質-一級", "費用大類", "费用大类",
         "Comp性質分類", "Comp性質-CN", "spend_category", "line_memo", "description"]


def _c(df, cands):
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if c in df.columns: return c
        if str(c).strip().lower() in low: return low[str(c).strip().lower()]
    return None


def _isblank(s):
    return s.astype(str).str.strip().isin(["", "0", "0.0", "nan", "None", "<NA>", "NaN"])


def _top(L, key, amt, n=15):
    g = pd.DataFrame({"k": key.astype(str).fillna(""), "a": amt}).groupby("k")["a"].agg(["sum", "size"])
    g = g.reindex(g["sum"].abs().sort_values(ascending=False).index)
    for k, r in g.head(n).iterrows():
        L.append(f"        {str(k)[:50]:<50} {r['sum']:>11,.0f}萬  ({int(r['size']):,}行)")


def main():
    L = ["# inspect_melco_h — 24 ~19億 冇 H 嘅 root cause (account_code driven)"]
    if not P.exists():
        L.append("!! melco parquet 揾唔到 (未跑?)"); _w(L); return
    df = pd.read_parquet(P)
    acode, adesc, amtc = _c(df, ACODE), _c(df, ADESC), _c(df, AMT)
    amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
    rp = df["report_period"].astype(str) if "report_period" in df.columns else pd.Series("", index=df.index)
    Hid = df["horizontal_id"].astype(str).fillna("") if "horizontal_id" in df.columns else pd.Series("", index=df.index)
    noH = Hid.isin(["", "nan", "None"])
    L.append(f"   cols: account_code='{acode}'  account_desc='{adesc}'  amount='{amtc}'")
    L.append(f"   no-H: {int(noH.sum()):,} 行  Σ={amt[noH].sum():,.0f}萬")

    # ── 1. account_code 空唔空 (overall vs no-H vs bucket) ──
    if acode:
        ab = _isblank(df[acode])
        L.append(f"\n   ① account_code 空白率: 全體 {ab.mean()*100:.0f}% | no-H 行 {ab[noH].mean()*100:.0f}% (證實 no-H = 冇 account_code?)")
        for b in ["25", "25_24SY", "24", "24_23SY", "23"]:
            m = rp == b
            if m.sum():
                L.append(f"      bucket {b:<8}: account_code 空 {ab[m].mean()*100:>3.0f}% ({int(ab[m].sum()):,}/{int(m.sum()):,}行)  no-H Σ={amt[m & noH].sum():,.0f}萬")

    # ── 2. no-H 行用咩欄有值 (邊個 project-team 欄可以砌 H) ──
    L.append(f"\n   ② no-H 行 各欄『有值』覆蓋 (揾邊個欄做 H source):")
    for c in [adesc] + SPEND:
        col = _c(df, [c]) if c else None
        if col and col in df.columns:
            nonblank = (~_isblank(df[col]))[noH]
            if nonblank.any():
                L.append(f"      '{col}': no-H 行有值 {nonblank.mean()*100:>3.0f}% ({int(nonblank.sum()):,}行)  Σ_有值={amt[noH & (~_isblank(df[col]))].sum():,.0f}萬")

    # ── 3. no-H top account_desc ──
    if adesc:
        L.append(f"\n   ③ no-H top account_desc (ledger_account 名) by Σ:")
        _top(L, df.loc[noH, adesc], amt[noH])

    # ── 4. no-H top 各 project-team spend 欄 ──
    for c in SPEND:
        col = _c(df, [c]) if c else None
        if col and col in df.columns and (~_isblank(df[col]))[noH].sum() > 0:
            L.append(f"\n   ④ no-H top '{col}' by Σ:")
            sub = noH & (~_isblank(df[col]))
            _top(L, df.loc[sub, col], amt[sub], n=20)
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_melco_h.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
