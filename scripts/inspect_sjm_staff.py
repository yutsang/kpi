r"""inspect_sjm_staff —— sjm opex staff：user 要 24=0、25=142（而家 24=28、25=0）。
揾 24 嗰 28 係咩（剷）+ 25 缺 142 喺邊（撈）。
Run（Windows）:  python scripts\inspect_sjm_staff.py
出 results\inspect_sjm_staff.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_sjm_staff.txt"
KW = r"salary|salaries|wage|payroll|staff|薪|工資|工资|員工|员工|人工|花紅|花红|bonus|gratuity|公積金|公积金|福利|津貼|津贴|獎金|奖金|prov\s*fund|MPF"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_sjm_staff —— sjm opex staff 24→0 / 25→142"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("sjm")].copy()
    e["pre"] = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    e["post"] = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["y2"] = e["year_bucket"].astype(str).str[:2]
    e["hid"] = _s(e.get("horizontal_id", pd.Series("", index=e.index)))
    e["cap"] = _s(e.get("final_capex_opex", pd.Series("", index=e.index))).eq("Capex")
    e["ad"] = _s(e.get("account_desc", pd.Series("", index=e.index)))
    e["code"] = _s(e.get("account_code", pd.Series("", index=e.index)))
    e["il"] = _s(e.get("is_labor", pd.Series("", index=e.index)))
    e["kw"] = e["ad"].str.contains(KW, case=False, regex=True, na=False)

    # A. 現 opex H_LABOR（= 24 嗰 28 / 25 嗰 0）
    L.append("\n## A. sjm opex H_LABOR by year × account（= 而家計入 opex staff）")
    lab = e[e.hid.eq("H_LABOR") & ~e.cap]
    for yr in ("23", "24", "25"):
        s = lab[lab.y2.eq(yr)]
        L.append(f"\n  [{yr}] opex staff Σ後={s['post'].sum():,.0f} Σ前={s['pre'].sum():,.0f}")
        for _, r in s.groupby(["code", "ad"])["post"].sum().reset_index().sort_values("post", key=lambda x: x.abs(), ascending=False).head(8).iterrows():
            L.append(f"     {(r['code'] or '')[:10]:<10}{(r['ad'] or '(空)')[:34]:<34}{r['post']:>8,.0f}萬")

    # B. opex 似人工 keyword 但唔係 H_LABOR（= 25 缺 142 可能喺度）
    L.append("\n## B. sjm opex『似人工』(keyword) 但唔係 H_LABOR by year（揾 25 嗰 142）")
    miss = e[~e.cap & e.kw & ~e.hid.eq("H_LABOR")]
    for yr in ("23", "24", "25"):
        s = miss[miss.y2.eq(yr)]
        if s.empty:
            L.append(f"\n  [{yr}] 冇"); continue
        L.append(f"\n  [{yr}] Σ後={s['post'].sum():,.0f} Σ前={s['pre'].sum():,.0f}")
        for _, r in s.groupby(["code", "ad", "hid"])["post"].sum().reset_index().sort_values("post", key=lambda x: x.abs(), ascending=False).head(8).iterrows():
            L.append(f"     {(r['code'] or '')[:10]:<10}{(r['ad'] or '(空)')[:30]:<30}H={r['hid']:<12}{r['post']:>8,.0f}萬")

    # C. opex is_labor flag（如果 sjm opex 有 is_labor）
    _ilon = ~e["il"].isin(["", "0", "0.0", "nan", "None", "N", "n", "False", "false", "FALSE", "No", "no"])
    op_il = e[~e.cap & _ilon]
    L.append(f"\n## C. sjm opex is_labor non-0：{len(op_il)}行 Σ後={op_il['post'].sum():,.0f}（by year）")
    for yr in ("23", "24", "25"):
        s = op_il[op_il.y2.eq(yr)]
        if len(s):
            L.append(f"  [{yr}] Σ後={s['post'].sum():,.0f}  現H: " + " | ".join(f"{k}={v:,.0f}" for k, v in s.groupby('hid')['post'].sum().items()))
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
