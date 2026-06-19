r"""揾 galaxy 23/24 嘅 opex staff cost account（25 已知：7926000/8000140/8000150/8000960 +
Casual Labor/Contract Services/Outsourced Contract Labor/Payroll-Direct Event）。

項目組 mark staff 嘅行（項目組H 含 'staff' / 'Staff cost'）= 真 staff。逐 bucket 列佢哋嘅
account_code + account_desc，等我哋砌 23/24 嘅規律（用 acct sum 大約對到 HQ 17,082 / 14,574）。

Run（Windows）:  python scripts\inspect_galaxy_staff_acct.py
出 results\inspect_galaxy_staff_acct.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
L: list[str] = ["# inspect_galaxy_staff_acct —— galaxy 23/24 opex staff account（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df = df[df["entity"].astype(str) == "galaxy"].copy()
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("項目組H", "horizontal_id", "final_capex_opex", "account_code", "account_desc"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    df["code"] = df["account_code"].str.replace(r"\.0$", "", regex=True)
    opex = ~df["final_capex_opex"].eq("Capex")
    # 項目組 mark 做 staff 嘅行（佢哋 mark 喺 項目組H = 'Staff cost...'）
    staff_mark = df["項目組H"].str.contains("staff", case=False, na=False)

    for yb in ("23", "24", "25"):
        sub = df[(df.yb == yb) & opex & staff_mark]
        L.append(f"\n{'='*64}\n## galaxy {yb}  項目組H=staff 嘅 opex 行 Σ={sub['post'].sum():,.0f}萬（{len(sub):,}行）")
        L.append(f"   ── by account_code（top）──")
        g = sub.groupby("code")["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(12).iterrows():
            descs = sub[sub.code == r["code"]]["account_desc"].value_counts().head(2).index.tolist()
            L.append(f"      {r['code']:<12} {r['post']:>9,.0f}萬  ({' / '.join(str(x)[:22] for x in descs)})")
        L.append(f"   ── by account_desc（top）──")
        gd = sub.groupby("account_desc")["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in gd.head(10).iterrows():
            L.append(f"      {r['account_desc'][:34]:<34} {r['post']:>9,.0f}萬")
    _w()


def _w():
    out = ROOT / "results" / "inspect_galaxy_staff_acct.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
