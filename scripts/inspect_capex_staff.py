r"""inspect capex staff cost —— user：capex staff 數字非常細，肯定有錯。

每 entity：H_LABOR Σ by (capex/opex × bucket)，睇 capex 人工 vs opex 人工。
若 capex 人工 ≈ 0 但理應有（staff 攤入 capex 項目）→ 即係 capex enforcement / capex勞工 tag 咗去建設。
再 dump capex 入面 account_desc 含 staff/salar/payroll/labor 但 tag 咗 建設/設施 嘅金額（= 被搶走嘅 capex staff）。

Run（Windows）:  python scripts\inspect_capex_staff.py
出 results\inspect_capex_staff.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
SKW = ["staff", "Staff", "STAFF", "salar", "Salar", "SALAR", "payroll", "Payroll", "PAYROLL",
       "人工", "薪", "工資", "casual", "Casual", "wage", "Wage", "Basic Salary"]
L: list[str] = ["# inspect_capex_staff —— capex 人工 vs opex 人工（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("entity", "horizontal_id", "horizontal_label", "final_capex_opex", "account_desc"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    cap = df["final_capex_opex"].eq("Capex")
    lab = df["horizontal_id"].eq("H_LABOR")
    staff_kw = df["account_desc"].apply(lambda s: any(k in str(s) for k in SKW))

    L.append("\n## §A H_LABOR Σ by (entity × capex/opex × bucket)")
    L.append(f"   {'entity':<8}{'co':<7}{'23':>9}{'24':>9}{'25':>9}")
    for ent in ENTS:
        for co in ("Capex", "Opex"):
            sub = df[(df.entity == ent) & lab & (df.final_capex_opex == co)]
            row = {yb: sub[sub.yb == yb]["post"].sum() for yb in ("23", "24", "25")}
            if sum(abs(v) for v in row.values()) < 1:
                continue
            L.append(f"   {ent:<8}{co:<7}{row['23']:>9,.0f}{row['24']:>9,.0f}{row['25']:>9,.0f}")

    L.append("\n## §B capex 入面 account_desc 似 staff 但 tag 咗 建設/設施（= 被搶嘅 capex staff）")
    cs = df[cap & staff_kw & df["horizontal_id"].isin(["H_CONSTRUCTION", "H_EQUIP"])]
    L.append(f"   共 {cs['post'].sum():,.0f}萬（{len(cs):,}行）")
    for ent in ENTS:
        e = cs[cs.entity == ent]
        if not len(e):
            continue
        L.append(f"   ▸ {ent}: {e['post'].sum():,.0f}萬  top account_desc:")
        for k, v in e.groupby("account_desc")["post"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(6).items():
            L.append(f"       {str(k)[:34]:<34} {v:,.0f}萬")
    _w()


def _w():
    out = ROOT / "results" / "inspect_capex_staff.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
