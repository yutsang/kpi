r"""staff gap 工具 —— 逐家逐年揾 row-level staff account 補/剷。
 opex(Table2 調整後)：need=golden−ours；ADD=opex 非人工但似 staff 嘅 account；現 H_LABOR=剷候選。
 capex(Table1 調整前 capex+opex)：capex staff-like account（補入 Table1 gap）。
keyword + account_code 雙信號；showing current H so 知搬邊個。

Run（Windows）:  python scripts\inspect_staff_gap.py galaxy mgm
                 python scripts\inspect_staff_gap.py wynn vml melco sjm
出 results\inspect_staff_gap_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy", "mgm"]
# HQ Table 2（opex 調整後）23/24/25
T2 = {"galaxy": {"23": 17082, "24": 14574, "25": 23950}, "wynn": {"23": 6928, "24": 8765, "25": 10958},
      "vml": {"23": 790, "24": 1457, "25": 1253}, "melco": {"23": 7218, "24": 13912, "25": 27806},
      "mgm": {"23": 6357, "24": 6829, "25": 21467}, "sjm": {"23": 0, "24": 0, "25": 142}}
# HQ Table 1（capex+opex 調整前）23/24/24_23SY
T1 = {"galaxy": {"23": 17082, "24": 34926, "24_23SY": 16872}, "sjm": {"23": 5554, "24": 4115, "24_23SY": 2485},
      "wynn": {"23": 10219, "24": 10219, "24_23SY": 9}, "vml": {"23": 791, "24": 12163, "24_23SY": 12163},
      "melco": {"23": 15989, "24": 14831, "24_23SY": 9}, "mgm": {"23": 7231, "24": 9270, "24_23SY": 3463}}
SDESC = ["Salar", "salar", "SALAR", "Payroll", "payroll", "PAYROLL", "Casual", "casual", "Staff", "staff", "STAFF",
         "Wage", "wage", "Contract Labo", "CONTRACT LABO", "Bonus", "Provident", "Overtime", "Allowance",
         "Pension", "人工", "薪", "工資", "Labour", "Labor", "LABOR", "Employee", "MPF", "Gratuity",
         "Severance", "Personnel", "Headcount", "Temp Staff", "Agency", "Recruit", "Outsourced", "Direct Event Invstm",
         "員工", "勞", "Manning", "Compensation"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def run(ent, df):
    L = [f"# inspect_staff_gap — {ent}（萬；opex 23/24=調整後,25=調整前；ADD=似staff非人工,現H_LABOR=剷候選）"]
    e = df[df["entity"].astype(str) == ent].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    yb = e["year_bucket"].astype(str)
    e["m"] = post.where(yb.isin(["23", "24"]), pre)
    e["pre"] = pre
    e["yb"] = yb
    for c in ("horizontal_id", "horizontal_label", "final_capex_opex", "account_code", "account_desc", "項目組H"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    e["blob"] = e["account_desc"] + " " + e["項目組H"] + " " + e["code"]
    e["sl"] = e["blob"].apply(lambda s: any(k in str(s) for k in SDESC))
    cap = e["final_capex_opex"].eq("Capex")
    lab = e["horizontal_id"].eq("H_LABOR")
    # ── opex (Table 2) ──
    for yb_v in ("23", "24", "25"):
        sub = e[(e.yb == yb_v) & ~cap]
        if not len(sub):
            continue
        ours = sub.loc[sub.horizontal_id.eq("H_LABOR"), "m"].sum()
        g = T2.get(ent, {}).get(yb_v, 0)
        L.append(f"\n{'='*64}\n## {ent} {yb_v} OPEX staff: golden={g:,.0f} 我={ours:,.0f} need={g-ours:+,.0f}")
        add = sub[~sub.horizontal_id.eq("H_LABOR") & sub.sl]
        ga = add.groupby(["code", "account_desc", "horizontal_label"])["m"].sum().reset_index()
        ga = ga[ga["m"].abs() >= 50].sort_values("m", key=lambda s: s.abs(), ascending=False)
        L.append(f"  ── ADD候選（opex 非人工但似 staff）Σ={add['m'].sum():,.0f} ──")
        for _, r in ga.head(12).iterrows():
            L.append(f"     {r['code']:<9}{r['account_desc'][:30]:<30} 現H={r['horizontal_label'][:8]:<8} {r['m']:>8,.0f}")
        cur = sub[sub.horizontal_id.eq("H_LABOR")]
        gc = cur.groupby(["code", "account_desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        L.append(f"  ── 現 H_LABOR（剷候選睇有冇非 staff）top ──")
        for _, r in gc.head(8).iterrows():
            sl = "" if any(k in str(r["account_desc"]) + str(r["code"]) for k in SDESC) else " ←非staff?"
            L.append(f"     {r['code']:<9}{r['account_desc'][:30]:<30} {r['m']:>8,.0f}{sl}")
    # ── capex staff-like (Table 1) ──
    L.append(f"\n{'='*64}\n## {ent} CAPEX staff-like（Table 1 補；萬=調整前）")
    for yb_v in ("23", "24", "24_23SY"):
        sub = e[(e.yb == yb_v) & cap]
        if not len(sub):
            continue
        g = T1.get(ent, {}).get(yb_v, 0)
        opex_lab = e.loc[(e.yb == yb_v) & ~cap & lab, "pre"].sum()
        cap_lab = sub.loc[sub.horizontal_id.eq("H_LABOR"), "pre"].sum()
        sl_cap = sub[sub.sl & ~sub.horizontal_id.eq("H_LABOR")]
        L.append(f"  {yb_v}: T1golden={g:,.0f} | opex人工(前)={opex_lab:,.0f} cap人工(前)={cap_lab:,.0f} | capex似staff非人工 Σ={sl_cap['pre'].sum():,.0f}")
        gs = sl_cap.groupby(["code", "account_desc", "horizontal_label"])["pre"].sum().reset_index()
        gs = gs[gs["pre"].abs() >= 50].sort_values("pre", key=lambda s: s.abs(), ascending=False)
        for _, r in gs.head(8).iterrows():
            L.append(f"       {r['code']:<9}{r['account_desc'][:28]:<28} 現H={r['horizontal_label'][:8]:<8} {r['pre']:>8,.0f}")
    p = ROOT / "results" / f"inspect_staff_gap_{ent}.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {p.relative_to(ROOT)}\n")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        run(ent, df)


if __name__ == "__main__":
    main()
