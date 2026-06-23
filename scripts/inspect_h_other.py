r"""inspect_h_other —— 一次過 dump H_OTHER（橫向科目其他）+ V_OTHER（縱向 Vertical=其他）成分。

user 2026-06-23：其他比例偏大。兩個其他要分開：
  A) H_OTHER（horizontal_id=H_OTHER）：科目層面其他 —— 部分係特登（運輸/稅/FoodCost/博彩contra 非投資）。
  B) V_OTHER（vertical_id=V_OTHER / vertical_label=其他）：分類層面其他 —— 路演清理/NG11 leftover。

A 拆：account_code × 項目組H × year × capex/opex。
B 拆：NG × 項目組V × subproject × account_desc —— 睇邊啲有真 V/真活動可歸返。

Run（Windows）:  python scripts\inspect_h_other.py
出 results\inspect_h_other.txt  ← paste 返嚟（太大就分段貼 galaxy/vml/mgm）
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
    L = ["# inspect_h_other —— H_OTHER + V_OTHER 成分拆解（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    df["ent"] = df["entity"].astype(str)
    for c, n in [("horizontal_id", "hid"), ("vertical_id", "vid"), ("vertical_label", "vl"),
                 ("account_code", "code"), ("account_desc", "ad"), ("description", "ds"),
                 ("項目組H", "pth"), ("項目組V", "ptv"), ("subproject", "sp"), ("project", "pj"),
                 ("ng_code", "ngc"), ("ng_label", "ng"), ("final_capex_opex", "co")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))

    # ════════════ A) H_OTHER ════════════
    L.append(f"\n{'#'*78}\n# A) H_OTHER（橫向科目其他）")
    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        tot = e["amt"].sum()
        o = e[e["hid"].eq("H_OTHER")]
        so = o["amt"].sum()
        L.append(f"\n{'='*70}\n## {ent}  H_OTHER={so:,.0f} / 總={tot:,.0f} = {so/tot*100:.2f}%  行={len(o):,}")
        if len(o) == 0:
            continue
        L.append(f"   ── by account_code top 20（Σ萬｜n｜desc｜項組H）──")
        g = o.groupby("code")["amt"].agg(["sum", "size"]).sort_values("sum", key=lambda s: s.abs(), ascending=False).head(20)
        for code, row in g.iterrows():
            sub = o[o["code"].eq(code)]
            ads = sub.loc[sub["ad"].ne(""), "ad"]
            adv = ads.value_counts().index[0] if len(ads) else ""
            pths = sub.loc[sub["pth"].ne(""), "pth"]
            pth_top = pths.value_counts().index[0] if len(pths) else "(空)"
            L.append(f"      {(code or '(空)')[:14]:<15}{row['sum']:>10,.0f} {int(row['size']):>6,}行 {adv[:28]:<28} 項組H={pth_top[:10]}")
        L.append(f"   ── by 項目組H（非空=項目組有真 H，可 reclass）──")
        for pth, v in o.groupby("pth")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(10).items():
            L.append(f"      {(pth or '(空)')[:24]:<26}{v:>10,.0f}")
        L.append(f"      [capex={o[o['co'].eq('Capex')]['amt'].sum():,.0f}  opex={o[~o['co'].eq('Capex')]['amt'].sum():,.0f}]")

    # ════════════ B) V_OTHER ════════════
    L.append(f"\n\n{'#'*78}\n# B) V_OTHER（縱向 Vertical Label=其他）")
    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        tot = e["amt"].sum()
        o = e[e["vid"].eq("V_OTHER") | e["vl"].eq("其他")]
        so = o["amt"].sum()
        L.append(f"\n{'='*70}\n## {ent}  V_OTHER={so:,.0f} / 總={tot:,.0f} = {so/tot*100:.2f}%  行={len(o):,}")
        if len(o) == 0:
            continue
        L.append(f"   ── by NG ──")
        for (ngc, ng), v in o.groupby(["ngc", "ng"])["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(8).items():
            L.append(f"      {ngc:<5}{(ng or '')[:14]:<15}{v:>10,.0f}")
        L.append(f"   ── by 項目組V（非空=項目組有真 V，可 reclass）──")
        for ptv, v in o.groupby("ptv")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(10).items():
            L.append(f"      {(ptv or '(空)')[:30]:<32}{v:>10,.0f}")
        L.append(f"   ── by subproject top 20（睇實際活動，揾 keyword 歸返 V）──")
        for spv, v in o.groupby("sp")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(20).items():
            L.append(f"      {(spv or '(空)')[:48]:<50}{v:>9,.0f}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
