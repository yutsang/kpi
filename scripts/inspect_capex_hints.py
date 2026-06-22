r"""inspect_capex_hints —— 喺 capex 行揾「資本化人工」hints（唔注入，只為 re-label）。

user 2026-06-22：唔注入任何嘢，capex 資本化人工項目組唔會比，只能喺現有 row 揾 hints。
逐家 capex 但唔係 H_LABOR 嘅行，喺 4 個維度掃 labor 訊號：
  A. 項目組H staff-ish（最硬，項目組自己標）—— 出精確 項目組H 值
  B. account_desc labor keyword（AUC-Design/salary/labour/設計…）
  C. vendor labor（payroll/manpower/staffing/人力/勞務…）
  D. description(JE memo) labor keyword
每段量化 per entity 可 re-label 幾多（pre/調整前 = HQ Table1 口徑）。

Run（Windows）:  python scripts\inspect_capex_hints.py [ent ...]  (default galaxy vml sjm mgm)
出 results\inspect_capex_hints.txt  ← paste 返嚟
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_capex_hints.txt"
ENTS = sys.argv[1:] or ["galaxy", "vml", "sjm", "mgm"]

PTH_STAFF = r"staff|人工|薪|payroll|labou?r|salary|wage|工資|headcount"
AD_LABOR  = r"salary|salaries|wage|payroll|labou?r|\bstaff\b|design\s*&|design and|headcount|人工|薪|勞務|設計|工資"
VEN_LABOR = r"payroll|manpower|staffing|labou?r|recruit|人力|勞務|outsourc|secondment|派遣"
DS_LABOR  = AD_LABOR

# HQ Table 1（capex+opex 調整前，24-prefix）vs 而家 cap+opex H_LABOR（pre）—— 量缺口
HQ_T1_24 = {"galaxy": 34926 + 16872, "wynn": 10219 + 9, "vml": 12163 + 12163,
            "melco": 14831 + 9, "mgm": 9270 + 3463, "sjm": 4115 + 2485}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_capex_hints —— capex 行資本化人工 hints（萬，調整前 pre，24-prefix；只為 re-label 唔注入）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    df["e"] = df["entity"].astype(str)
    for c, n in [("horizontal_id", "hid"), ("final_capex_opex", "co"), ("account_code", "code"),
                 ("account_desc", "ad"), ("項目組H", "pth"), ("vendor", "ven"), ("description", "ds")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["code"] = df["code"].str.replace(r"\.0$", "", regex=True)
    df["cap"] = df["co"].eq("Capex")
    df["islab"] = df["hid"].eq("H_LABOR")

    for ent in ENTS:
        e = df[df.e.eq(ent) & df.y2.eq("24")]
        cap_nl = e[e.cap & ~e.islab]                       # capex 但唔係人工
        cur_lab = e[e.islab]["pre"].sum()
        L.append("\n" + "█" * 80)
        L.append(f"█ {ent}  capex非人工 Σ(pre)={cap_nl['pre'].sum():,.0f}  | 現cap+opex人工={cur_lab:,.0f} vs HQ-Tab1={HQ_T1_24.get(ent,0):,} (Δ={cur_lab-HQ_T1_24.get(ent,0):+,.0f})")
        L.append("█" * 80)
        if cap_nl.empty:
            L.append("  (冇 capex 非人工行)"); continue

        def seg(letter, name, mask, showcol):
            sub = cap_nl[mask]
            L.append(f"\n  {letter}. {name}：{len(sub):,}行 Σ(pre)={sub['pre'].sum():,.0f}")
            if len(sub):
                g = sub.groupby(showcol)["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
                for _, r in g.head(8).iterrows():
                    lbl = " | ".join(str(r[c])[:30] for c in (showcol if isinstance(showcol, list) else [showcol]))
                    L.append(f"       {lbl[:56]:<56}{r['pre']:>10,.0f}")

        seg("A", "項目組H staff-ish（最硬）",
            cap_nl["pth"].str.contains(PTH_STAFF, case=False, regex=True, na=False), ["pth", "ad"])
        seg("B", "account_desc labor keyword（要眼睇）",
            cap_nl["ad"].str.contains(AD_LABOR, case=False, regex=True, na=False), ["code", "ad"])
        seg("C", "vendor labor（資本化外判勞務）",
            cap_nl["ven"].str.contains(VEN_LABOR, case=False, regex=True, na=False), ["ven", "ad"])
        seg("D", "description(JE memo) labor",
            cap_nl["ds"].str.contains(DS_LABOR, case=False, regex=True, na=False), ["ds"])

        # union（唔重複）—— 全部 hint 命中可 re-label 上限
        any_hint = (cap_nl["pth"].str.contains(PTH_STAFF, case=False, regex=True, na=False)
                    | cap_nl["ad"].str.contains(AD_LABOR, case=False, regex=True, na=False)
                    | cap_nl["ven"].str.contains(VEN_LABOR, case=False, regex=True, na=False)
                    | cap_nl["ds"].str.contains(DS_LABOR, case=False, regex=True, na=False))
        hard = cap_nl["pth"].str.contains(PTH_STAFF, case=False, regex=True, na=False)
        L.append(f"\n  → 可 re-label 上限：硬(項目組H)={cap_nl[hard]['pre'].sum():,.0f}萬 ｜ 連 desc/vendor/memo={cap_nl[any_hint]['pre'].sum():,.0f}萬")
        L.append(f"     （galaxy 估 ~7,420 硬；vml/sjm 若全段≈0 = 確認 capex 真係冇 labor，cap+opex 只能接受 under）")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
