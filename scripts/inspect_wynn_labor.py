r"""驗 wynn 人工規律：人工成本 ⟺ Nature of Expenses(account_desc)='Staff Cost'。

  §A wynn opex 人工成本 by account_desc（Nature of Expenses）→ Σ調整後，標★邊啲係 Staff Cost
      → 睇而家幾多人工係「非 Staff Cost」(要剷走) 同確認 exact string
  §B account_desc='Staff Cost' 但而家唔係人工嘅（要補入）
  §C 「非 Staff Cost 但被當人工」嘅行 → 而家 account_desc + 應該係咩 H（睇佢哋本身性質）

Run（Windows）:  python scripts\inspect_wynn_labor.py
出 results\inspect_wynn_labor.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
L: list[str] = ["# inspect_wynn_labor —— wynn 人工 ⟺ Nature of Expenses='Staff Cost'?（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df = df[df["entity"].astype(str) == "wynn"].copy()
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("horizontal_id", "horizontal_label", "final_capex_opex", "account_desc"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    opex = ~df["final_capex_opex"].eq("Capex")
    acc = df["account_desc"].astype(str)
    is_sc = acc.str.contains("Staff Cost", case=False, na=False)

    lab = df[(df.horizontal_id == "H_LABOR") & opex]
    L.append(f"\n  wynn opex 人工成本 共 {lab['post'].sum():,.0f}萬（{len(lab):,}行）")
    L.append(f"\n  ── §A 人工 by account_desc（★=含 'Staff Cost'）──")
    g = lab.groupby("account_desc")["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(20).iterrows():
        star = "★" if "staff cost" in str(r["account_desc"]).lower() else " "
        L.append(f"   {star} {r['account_desc'][:38]:<38} {r['post']:>9,.0f}萬")
    sc_in = lab[is_sc.loc[lab.index]]["post"].sum()
    nonsc_in = lab[~is_sc.loc[lab.index]]["post"].sum()
    L.append(f"\n   人工內：Staff Cost={sc_in:,.0f}萬  /  非Staff Cost={nonsc_in:,.0f}萬（要剷走）")

    L.append(f"\n  ── §B account_desc='Staff Cost' 但而家唔係人工（要補入）──")
    miss = df[is_sc & opex & (df.horizontal_id != "H_LABOR")]
    if len(miss):
        for k, v in miss.groupby("horizontal_label")["post"].sum().sort_values(key=lambda s: s.abs(), ascending=False).items():
            L.append(f"     現H={k or '(空)':<12} {v:,.0f}萬")
    else:
        L.append("     （冇 → 所有 Staff Cost 已係人工）")

    L.append(f"\n  ── §C 非Staff Cost 但被當人工（要剷走，睇佢哋應落咩 H）by account_desc × bucket ──")
    bad = lab[~is_sc.loc[lab.index]]
    gb = bad.groupby(["account_desc", "yb"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in gb.head(15).iterrows():
        L.append(f"     {r['account_desc'][:30]:<30} {r['yb']:<9} {r['post']:>8,.0f}萬")

    L.append(f"\n  ── 人工 by bucket（對 HQ wynn staff 6,928/8,765/10,958）──")
    for yb, v in lab.groupby("yb")["post"].sum().items():
        L.append(f"     {yb:<9} {v:>9,.0f}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_wynn_labor.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
