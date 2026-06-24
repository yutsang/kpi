"""inspect_mgm_adj.py — mgm tagged_rows 各 bucket 的 pre/post/adj 欄對應
Run: python scripts\inspect_mgm_adj.py
Out: results\inspect_mgm_adj.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
PQ   = ROOT / "data" / "mgm" / "interim" / "company_6_tagged_rows.parquet"
OUT  = ROOT / "results" / "inspect_mgm_adj.txt"

PRE_COLS  = ["25_Amt", "24_Amt", "23_Amt"]
ADJ_COLS  = ["調整金額", "24_調整金額", "23_調整金額"]
POST_COLS = ["調整後金額", "24_調整后", "23_調整后"]
LV_COLS   = ["Adj_lv1", "Adj_lv2", "調整項目名稱", "調整事項備註"]


def _num(s):
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False), errors="coerce").fillna(0.0)


def main():
    L = ["# inspect_mgm_adj — mgm per-bucket adjustment column mapping", ""]
    if not PQ.exists():
        L.append(f"!! {PQ} 揾唔到"); _w(L); return

    df = pd.read_parquet(PQ)
    n = len(df)
    L.append(f"Rows: {n:,}  year_bucket values: {sorted(df['year_bucket'].astype(str).unique())}")

    amt_mop = _num(df["amount_mop"]) if "amount_mop" in df.columns else pd.Series(0.0, index=df.index)

    L += ["", "=" * 70, "## 1. Per year_bucket: amount_mop + pre/adj/post cols Sigma (万)"]
    all_check_cols = PRE_COLS + ADJ_COLS + POST_COLS
    avail = [c for c in all_check_cols if c in df.columns]

    hdr = f"  {'bucket':<12}  {'rows':>7}  {'amount_mop':>12}" + "".join(f"  {c[:14]:>16}" for c in avail)
    L.append(hdr)

    for bk in sorted(df["year_bucket"].astype(str).unique()):
        m = df["year_bucket"].astype(str).eq(bk)
        row_str = f"  {bk:<12}  {int(m.sum()):>7,}  {amt_mop[m].sum()/1e4:>12,.0f}"
        for c in avail:
            v = _num(df[c])[m].sum() / 1e4
            row_str += f"  {v:>16,.0f}"
        L.append(row_str)

    L += ["", "=" * 70, "## 2. amount_mop vs pre/adj/post per bucket"]
    bucket_map = {
        "25":     ("25_Amt",  "調整金額",   "調整後金額"),
        "24":     ("24_Amt",  "24_調整金額", "24_調整后"),
        "23":     ("23_Amt",  "23_調整金額", "23_調整后"),
        "25_24SY":("24_Amt",  "24_調整金額", "24_調整后"),
        "25_23SY":("23_Amt",  "23_調整金額", "23_調整后"),
        "24_23SY":("23_Amt",  "23_調整金額", "23_調整后"),
    }
    for bk in sorted(df["year_bucket"].astype(str).unique()):
        m = df["year_bucket"].astype(str).eq(bk)
        if not m.any(): continue
        pre_c, adj_c, post_c = bucket_map.get(bk, ("","",""))
        mop_s  = amt_mop[m].sum() / 1e4
        pre_s  = _num(df[pre_c])[m].sum()  / 1e4 if pre_c  and pre_c  in df.columns else 0.0
        adj_s  = _num(df[adj_c])[m].sum()  / 1e4 if adj_c  and adj_c  in df.columns else 0.0
        post_s = _num(df[post_c])[m].sum() / 1e4 if post_c and post_c in df.columns else 0.0
        L.append(f"  [{bk}] rows={int(m.sum()):,}  amount_mop={mop_s:,.0f}  "
                 f"pre={pre_s:,.0f}  adj={adj_s:,.0f}  post={post_s:,.0f}  "
                 f"pre+adj={pre_s+adj_s:,.0f}  mop-post={mop_s-post_s:,.2f}")

    L += ["", "=" * 70, "## 3. lv1/lv2 per bucket non-null + top values"]
    for c in [c for c in LV_COLS if c in df.columns]:
        L.append(f"  [{c}]  total non-null={int(df[c].notna().sum()):,}")
        for bk in sorted(df["year_bucket"].astype(str).unique()):
            m = df["year_bucket"].astype(str).eq(bk)
            nn = int(df[c][m].notna().sum())
            if nn > 0:
                vc = df[c][m].dropna().astype(str).value_counts().head(3)
                L.append(f"    [{bk}] {nn}: " + " | ".join(f"{v}:{n}" for v,n in vc.items()))

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
