r"""inspect_h_other —— 每家 H_OTHER（其他）成分拆解，揾邊啲可以歸返真 H。

user 2026-06-23：其他比例偏大（galaxy 5.47%/vml 2.43%/mgm 3.12%）。
逐家 dump H_OTHER：
  1) by account_code top25：Σ萬 + n + account_desc 樣本 + 項目組H 最常見值
  2) by 項目組H（項目組自己嘅 H label）：睇我哋擺咗去其他、但項目組其實有真 H 嘅
  3) by year_bucket
→ 若大額落喺某 account 而項目組H = 真 H（e.g. 建設/器具/廣告/comp），就可以 reclass。

Run（Windows）:  python scripts\inspect_h_other.py
出 results\inspect_h_other.txt  ← paste 返嚟（太大就 head）
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_h_other.txt"
ENTS = ["galaxy", "vml", "mgm", "melco", "wynn", "sjm"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_h_other —— H_OTHER 成分拆解（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    hid = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    df["ent"] = df["entity"].astype(str)
    for c, n in [("account_code", "code"), ("account_desc", "ad"), ("description", "ds"),
                 ("項目組H", "pth"), ("subproject", "sp"), ("project", "pj"),
                 ("科目明細", "km"), ("final_capex_opex", "co"), ("ng_label", "ng")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))

    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        tot = e["amt"].sum()
        o = e[hid.loc[e.index].eq("H_OTHER")]
        so = o["amt"].sum()
        L.append(f"\n{'='*78}\n## {ent}  其他={so:,.0f}萬 / 總={tot:,.0f}萬 = {so/tot*100:.2f}%   行={len(o):,}")
        if len(o) == 0:
            continue
        # 1) by account_code
        L.append(f"   ── by account_code top 25（Σ萬｜n｜account_desc樣本｜項目組H最常見）──")
        g = o.groupby("code")["amt"].agg(["sum", "size"]).sort_values("sum", key=lambda s: s.abs(), ascending=False).head(25)
        for code, row in g.iterrows():
            sub = o[o["code"].eq(code)]
            ads = sub["ad"].value_counts().index[0] if (sub["ad"] != "").any() else ""
            pth = sub.loc[sub["pth"].ne(""), "pth"]
            pth_top = pth.value_counts().index[0] if len(pth) else "(空)"
            L.append(f"      {(code or '(空)')[:14]:<15}{row['sum']:>11,.0f}萬 {int(row['size']):>7,}行  "
                     f"{ads[:30]:<30} 項組H={pth_top[:10]}")
        # 2) by 項目組H（睇項目組有冇真 H）
        L.append(f"   ── by 項目組H（項目組自己 label；非空=可能可 reclass）──")
        for pth, v in o.groupby("pth")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(12).items():
            L.append(f"      {(pth or '(空)')[:24]:<26}{v:>11,.0f}萬")
        # 3) by year_bucket
        L.append(f"   ── by year_bucket ──")
        for yb, v in o.groupby(_s(o.get("year_bucket")))["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).items():
            L.append(f"      {yb:<12}{v:>11,.0f}萬")
        # 4) capex/opex split
        L.append(f"      [capex={o[o['co'].eq('Capex')]['amt'].sum():,.0f}萬  opex={o[~o['co'].eq('Capex')]['amt'].sum():,.0f}萬]")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
