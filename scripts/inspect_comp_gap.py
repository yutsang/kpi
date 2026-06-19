r"""per-component gap finder — 對 internal_resource_golden（23/24，調整後）逐 component 揾 row-level 來源。
  under（golden>ours）→ scan 其他 H 入面 keyword 命中嘅行（mis-tag 嘅 comp/staff，應搬入）
  over （ours>golden）→ 列現 component 行 by account_desc（應搬走/收）
錢全部 tie（capex/opex Δ≈0）→ gap = 純 H 派錯位，唔係 top-down。

Run（Windows）:  python scripts\inspect_comp_gap.py galaxy
                 python scripts\inspect_comp_gap.py wynn vml melco mgm sjm
出 results\inspect_comp_gap_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
GOLD = ROOT / "internal_resource_golden.xlsx"
ENTS = sys.argv[1:] or ["galaxy"]

# component → (golden col, our horizontal_id, keyword regex 去其他 H 搵 mis-tag)
COMP = [
    ("人工",   "internal_staff_cost", "H_LABOR",
     r"Salar|Payroll|Casual|Staff|Wage|Contract Labo|Bonus|Provident|Overtime|Allowance|Pension|人工|薪|工資|Labour|Labor"),
    ("客房",   "客房支出", "H_HOTEL_ROOM",
     r"Lodging|Room|Hotel|客房|房晚|Accommodation|Suite|Front Desk|Ritz|Marriott|Okura|StRegis|St\. Regis|Banyan|住宿|Expense in Kind"),
    ("會場",   "會場支出", "H_VENUE",
     r"Venue|Hall|Meeting|Ballroom|Function|場地|會場|Event Space|Special Event|Event Service|Rental|租|Other Operating|Expense in Kind"),
    ("食飲",   "食品飲料支出", "H_FNB",
     r"Food|Beverage|F&B|餐飲|食品|Banquet|Catering|Dining|Restaurant|Service Charge|Buffet|Bar |飲|宴|Expense in Kind"),
    ("贈票",   "贈票支出", "H_COMP_TICKET",
     r"Ticket|門票|Admission|Club Use|Spa|Show|Concert|演唱|表演|入場|贈票|Salon"),
    ("comp其他", "comp_其他", "H_COMP_OTHER",
     r"Comp|Barter|Transport|Tourism|Joint Vent|External|Contra|In Kind|Outside Comp"),
]
NG_MAP = {"NG0": "博彩娛樂", "NG1": "吸引外國客源", "NG2": "會議展覽", "NG3": "娛樂表演",
          "NG4": "體育盛事", "NG5": "文化藝術", "NG6": "健康養生", "NG7": "主題遊樂",
          "NG8": "美食之都", "NG9": "社區旅遊", "NG10": "海上旅遊", "NG11": "其他"}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def load_gold():
    g = pd.read_excel(GOLD)
    g.columns = [str(c).strip() for c in g.columns]
    g["entity"] = g["entity"].astype(str).str.strip()
    g["year"] = g["year"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    for _, col, _, _ in COMP:
        g[col] = pd.to_numeric(g.get(col, 0), errors="coerce").fillna(0.0)
    return g


def run(ent, df, gold):
    L = [f"# inspect_comp_gap — {ent}（golden 23/24 調整後；under→其他H揾,over→現行收）"]
    e = df[df["entity"].astype(str) == ent].copy()
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    e["m"] = post
    e["yb"] = e["year_bucket"].astype(str)
    for c in ("horizontal_id", "horizontal_label", "comp_type", "account_desc", "account_code", "項目組H", "ng_code", "final_capex_opex"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    e["blob"] = (e["account_desc"] + " " + e["comp_type"] + " " + e["項目組H"])
    ge = gold[gold["entity"] == ent]
    for yr in ("23", "24"):
        sub = e[e.yb == yr]
        gy = ge[ge["year"] == yr]
        if not len(gy):
            continue
        L.append(f"\n{'='*64}\n## {ent} {yr}")
        for lab, gcol, hid, kw in COMP:
            gv = gy[gcol].sum()
            ov = sub.loc[sub.horizontal_id == hid, "m"].sum()
            d = ov - gv
            flag = "✓" if abs(d) < 500 else ("OVER" if d > 0 else "UNDER")
            L.append(f"\n  ── {lab}: golden={gv:,.0f} 我={ov:,.0f} Δ={d:,.0f} [{flag}] ──")
            if abs(d) < 500:
                continue
            if d < 0:  # under → 去其他 H scan keyword（capex 排除 staff 以外；comp 唔理 capex/opex）
                cand = sub[(sub.horizontal_id != hid) & sub["blob"].str.contains(kw, case=False, regex=True, na=False)]
                if hid != "H_LABOR":  # comp 通常 opex；但唔強制
                    pass
                g = cand.groupby(["horizontal_label", "account_desc"])["m"].sum().reset_index()
                g = g[g["m"].abs() >= 100].sort_values("m", key=lambda s: s.abs(), ascending=False)
                L.append(f"     需 +{-d:,.0f}萬。其他 H 命中 keyword（搬入候選）top:")
                for _, r in g.head(12).iterrows():
                    L.append(f"       現H={r['horizontal_label'][:9]:<9} {r['account_desc'][:30]:<30} {r['m']:>8,.0f}萬")
            else:  # over → 列現有行
                cur = sub[sub.horizontal_id == hid]
                g = cur.groupby(["account_desc", "comp_type"])["m"].sum().reset_index()
                g = g.sort_values("m", key=lambda s: s.abs(), ascending=False)
                L.append(f"     超 {d:,.0f}萬。現 {lab} 行（搬走候選）top:")
                for _, r in g.head(10).iterrows():
                    L.append(f"       {r['account_desc'][:30]:<30} ct={r['comp_type'][:12]:<12} {r['m']:>8,.0f}萬")
    p = ROOT / "results" / f"inspect_comp_gap_{ent}.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {p.relative_to(ROOT)}\n")


def main():
    if not CSV.exists() or not GOLD.exists():
        print("!! csv/golden 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    gold = load_gold()
    for ent in ENTS:
        run(ent, df, gold)


if __name__ == "__main__":
    main()
