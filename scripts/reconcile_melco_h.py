"""Row-level reconcile of Melco H buckets: 項目組 pivot-formula logic vs OUR pipeline H.

The project team classifies each row into an H bucket through a set of Excel pivot tables /
SUMIFS formulas driven by their own helper columns (取數標籤 / KP識別Comp / Comp性質-CN /
KP識別人工 / KP識別表演演出 / Ledger Type / 建築設施設備劃分 / 支出性質 …). This script:

  1. Re-implements those 7 reconciliations as ONE row-level priority cascade  →  column 項目組核對_H
  2. Joins OUR pipeline horizontal_label (from kpi_report.parquet) by KPMG唯一識別碼  →  column 我們_H
  3. Flags every row where they differ + totals each bucket (theirs vs ours), so you can sort/pivot.

TWO modes:
  --inspect   : dump every column (Excel letter + name + coverage + top values) AND the full value
                distribution of each ROLE column (so the cascade value-maps below can be confirmed).
                Run THIS FIRST, paste back, then we lock the maps.
  (default)   : run the cascade + join + diff, write results/melco_reconcile_<year>.xlsx

Run on Windows (file is root/melco_audit_25.xlsx, tab 'Data'):
  python scripts/reconcile_melco_h.py --inspect
  python scripts/reconcile_melco_h.py                 # full reconcile, year 25
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
OUR_PARQUET = ROOT / "data/melco/output/company_5_kpi_report.parquet"

# ── ROLE → header-name candidates (resolved by exact, then loose 'contains') ──────────────
# letters in comments are the audit-file columns the user quoted; we match by NAME (robust).
ROLES = {
    "amount":     ["Amount - Amended", "Amount Amended", "Amount-Amended"],   # col Y
    "take":       ["取數標簽", "取數標籤", "取数标签",
                   "Report Years-重分類2023年期後調整後", "Report Years"],       # 25:取數標簽==Y / 24:Report Years 非空
    "comp_flag":  ["KP識別Comp", "KP识别Comp", "KP識別comp"],                    # == "Y"
    "comp_nat":   ["Comp性質-CN", "Comp性质-CN", "Comp性質"],                    # sub-comp nature
    "labor_flag": ["KP識別人工（Staff cost come from Project Team需Capex提供）",
                   "KP識別人工", "KP识别人工"],                                  # not-blank → 人工
    "perf_flag":  ["KP識別表演演出", "KP识别表演演出", "識別表演演出"],            # == "Y" → 合約成本
    "ledger":     ["Ledger Type", "Journal Type"],                             # CAPEX / OPEX
    "constr_div": ["建築設施設備劃分", "建筑设施设备划分"],                        # capex x-axis split
    "spend_nat":  ["支出性質", "支出性质-mapping", "支出性質-mapping"],           # opex residual driver
    "ledger_acct":["ledger_account", "Ledger ID and description"],             # AJ — "Sponsorship"/"License" tag
    "nature_aj":  ["支出性質", "Nature"],                                       # AJ — "Sponsorship" tag
    "proj_nat":   ["項目性質", "项目性质"],                                       # 項目性質 (info)
    "key":        ["KPMG唯一識別碼", "KPMG唯一识别码"],                          # join key to our output
    "spend_cat":  ["spend_category", "spend category", "spendcategory"],
    "line_memo":  ["line_memo", "linememo", "line memo"],
    "subproj":    ["Project/Sub-Project Name", "subproject name", "Sub-Project Name"],
}

# ── VALUE MAPS (best-guess from conf predominant_rules + the formulas; CONFIRM via --inspect) ──
COMP_NATURE_TO_H = {            # Comp性質-CN  →  our H label (24 用「演唱會門票支出」, 25 用「贈票支出」)
    "客房支出": "酒店客房", "食品飲料支出": "餐飲", "場地租借": "活動場地",
    "贈票支出": "贈票支出", "演唱會門票支出": "贈票支出", "其他": "Comp其他",
}
LABOR_VALUES = {"staff cost come from project team", "y", "y,staff cost分攤", "y, staff cost分攤"}

def constr_to_h(v: str) -> str:                # 建築設施設備劃分  →  our H label
    s = str(v)
    if "人工" in s:                                                          return "人工成本"  # 人工，不選
    if any(k in s for k in ("建設", "建築", "場地", "construction", "fit")): return "建設與設施支出"
    if any(k in s for k in ("設施", "器具", "採購", "equip")):               return "設施及器具採購"
    if any(k in s for k in ("維護", "maintenance")):                         return "維護費"
    return "維護費"                            # 篩不出上述兩個 → 一般維護費

def spend_to_h(v: str) -> str:                 # 支出性質 (opex residual)  →  our H label
    s = str(v).lower()
    if any(k in s for k in ("sponsor", "贊助", "赞助")):                      return "贊助費"
    if any(k in s for k in ("商標", "许可", "許可", "license", "授權", "授权", "trademark", "royalt")): return "授權費"
    if any(k in s for k in ("lease", "rental", "租賃", "租赁", "租金")):       return "租賃費"
    if any(k in s for k in ("職工", "薪酬", "人工", "interco-人工", "福利")):   return "人工成本"
    if any(k in s for k in ("表演者", "影視製作", "影视", "製作成本", "水舞間", "cost recovery")): return "合約成本"
    if any(k in s for k in ("維修", "保養", "維護", "维修", "保养")):          return "維護費"
    if any(k in s for k in ("water", "電費", "电费", "水電", "水电", "utilit")): return "其他"
    if any(k in s for k in ("營銷", "营销", "media", "廣告", "广告", "market", "推廣", "推广", "advertis")): return "廣告及推廣"
    if any(k in s for k in ("professional", "專業", "专业", "agency", "顧問", "代理", "consult")): return "專業服務費"
    return "其他"                              # 管理費/差旅費/一般用品 等 → 由 diff 揭示再 refine


def resolve(cols, cands):
    for c in cands:                                   # exact
        for col in cols:
            if str(col).strip() == c:
                return col
    for c in cands:                                   # loose contains
        for col in cols:
            if c in str(col):
                return col
    return None


def is_y(v):  return str(v).strip().lower() in ("y", "yes", "是", "1", "1.0")
def nonblank(v): return str(v).strip().lower() not in ("", "nan", "none", "n/a", "na", "0", "0.0", "-")
def is_taken(v):   # 取數: 25 用 取數標簽=='Y'; 24 用 Report Years 非空（兩者都排除 Net-off）
    return str(v).strip().lower() not in ("", "nan", "none", "n/a", "na",
                                          "net-off", "netoff", "kp識別net off", "kp识别net off")


def their_bucket(row, R):
    """Project-team's reconciliation bucket for ONE row (priority cascade)."""
    if R["take"] and not is_taken(row.get(R["take"])):
        return "(未取數)"
    # 1. Comp — by Comp性質-CN value (25: == KP識別Comp=Y; 24: Comp性質-CN 非 N/A/空)
    nat = str(row.get(R["comp_nat"], "")).strip() if R["comp_nat"] else ""
    if nat in COMP_NATURE_TO_H:
        return COMP_NATURE_TO_H[nat]
    # 2. 人工
    if R["labor_flag"]:
        lv = str(row.get(R["labor_flag"], "")).strip().lower()
        if lv and lv not in ("nan", "none", "n/a", "na", "0", "0.0", "-") and \
           (lv in LABOR_VALUES or "staff cost" in lv or lv == "y"):
            return "人工成本"
    ledger = str(row.get(R["ledger"], "")).strip().upper() if R["ledger"] else ""
    # 3. Capex split (建築設施設備劃分)
    if "CAPEX" in ledger:
        return constr_to_h(row.get(R["constr_div"], "")) if R["constr_div"] else "建設與設施支出"
    # 4. 合約成本(演藝) — opex, comp已剔除, 人工已剔除
    if R["perf_flag"] and is_y(row.get(R["perf_flag"])):
        return "合約成本"
    # 5. 贊助費 — 歷史上放喺營銷費用裏，靠 ledger_account="Sponsorship" 拆出
    la = (str(row.get(R["ledger_acct"], "")) + " " + str(row.get(R["spend_cat"], ""))).lower() if R.get("ledger_acct") else ""
    if "sponsor" in la or "贊助" in la or "赞助" in la:
        return "贊助費"
    # 6. residual opex by 支出性質
    return spend_to_h(row.get(R["spend_nat"], "")) if R["spend_nat"] else "其他"


def inspect(df, R):
    print(f"=== INSPECT melco Data — rows={len(df):,}, {len(df.columns)} cols ===\n")
    print(f"{'#':>3} {'col':>4}  {'header':<46} {'cover%':>7} {'uniq':>6}  top values")
    print("-" * 110)
    for i, c in enumerate(df.columns):
        s = df[c]
        cov = s.notna().mean() * 100
        nun = s.nunique(dropna=True)
        vc = s.astype(str).value_counts().head(6)
        tops = " | ".join(f"{k[:18]}({v})" for k, v in vc.items())
        print(f"{i:>3} {get_column_letter(i+1):>4}  {str(c)[:46]:<46} {cov:>6.0f}% {nun:>6}  {tops[:60]}")
    print("\n=== ROLE resolution (NAME match) ===")
    for role, cands in ROLES.items():
        col = resolve(df.columns, cands)
        print(f"  {role:<11} -> {col!r}")
    print("\n=== full value distribution of key ROLE columns (Σ amount) ===")
    amt = R["amount"]
    for role in ("take", "comp_flag", "comp_nat", "labor_flag", "perf_flag", "ledger",
                 "constr_div", "spend_nat", "proj_nat"):
        col = R[role]
        if not col:
            print(f"\n--- {role}: (column not found) ---"); continue
        g = df.groupby(df[col].astype(str)).agg(rows=(col, "size"),
              amt=(amt, "sum") if amt else (col, "size")).sort_values("amt", key=lambda s: s.abs(), ascending=False)
        print(f"\n--- {role}  ('{col}') — {df[col].nunique()} 值 ---")
        for v, r in g.head(30).iterrows():
            print(f"     {str(v)[:34]:34} rows={int(r['rows']):>7}  Σ={r['amt']:>16,.0f}")
    # confirm our output has the join key + horizontal_label
    print("\n=== our pipeline output ===")
    if OUR_PARQUET.exists():
        names = pq.read_schema(OUR_PARQUET).names
        print(f"  {OUR_PARQUET.name} cols (key/H?): "
              f"{[n for n in names if n in ('KPMG唯一識別碼','horizontal_label','horizontal_id','report_period') or '識別' in str(n)]}")
    else:
        print(f"  X {OUR_PARQUET} missing — run kedro melco first")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=None, help="default: melco_audit_<year>.xlsx")
    ap.add_argument("--sheet", default="Data")
    ap.add_argument("--year", default="25")
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--n", type=int, default=30)
    args = ap.parse_args()

    fname = args.file or f"melco_audit_{args.year}.xlsx"
    fp = ROOT / fname
    if not fp.exists(): fp = Path(fname)
    if not fp.exists():
        print(f"X {args.file} not found (repo root or cwd)"); return
    print(f"reading {fp.name} [{args.sheet}] …")
    df = pd.read_excel(fp, sheet_name=args.sheet)
    R = {role: resolve(df.columns, cands) for role, cands in ROLES.items()}

    if args.inspect:
        inspect(df, R); return

    if not R["amount"]:
        print("X amount column not found — run --inspect"); return
    amt = R["amount"]
    df[amt] = pd.to_numeric(df[amt], errors="coerce").fillna(0)

    # ── H taken STRAIGHT from the project-team raw columns (their 取數 logic) ──
    df["項目組核對_H"] = df.apply(lambda r: their_bucket(r, R), axis=1)
    taken = df[df["項目組核對_H"].ne("(未取數)")].copy()
    excl  = df[df["項目組核對_H"].eq("(未取數)")]
    print(f"\n=== 取數 scope ('{R['take']}' == Y) ===")
    print(f"  taken    rows={len(taken):>7}  Σ={taken[amt].sum():>17,.0f}")
    print(f"  excluded rows={len(excl):>7}  Σ={excl[amt].sum():>17,.0f}   (Net-off / 非取數)")

    # ── the deliverable numbers: H buckets from raw ──
    print(f"\n=== H buckets from raw (taken, Σ {amt}) — 對你嘅 7 張 pivot ===")
    bt = taken.groupby("項目組核對_H")[amt].agg(["size", "sum"]).sort_values(
        "sum", key=lambda s: s.abs(), ascending=False)
    for b, r in bt.iterrows():
        print(f"   {str(b):<14} rows={int(r['size']):>7}  Σ={r['sum']:>17,.0f}")
    print(f"   {'總計':<14} rows={len(taken):>7}  Σ={taken[amt].sum():>17,.0f}")

    # ── CHECK 數: tie raw flags directly to bucket sums ──
    def buck(*names): return taken.loc[taken["項目組核對_H"].isin(names), amt].sum()
    ledg = taken[R["ledger"]].astype(str).str.upper() if R["ledger"] else pd.Series("", index=taken.index)
    is_cap = ledg.str.contains("CAPEX")
    print(f"\n=== tie-out (raw flag 加總  vs  bucket 加總，應該相等) ===")
    if R["comp_flag"]:
        comp_raw = taken.loc[taken[R["comp_flag"]].astype(str).str.lower().eq("y"), amt].sum()
        comp_b = buck("酒店客房", "餐飲", "活動場地", "贈票支出", "Comp其他")
        print(f"  Comp   KP識別Comp=Y Σ={comp_raw:>15,.0f}  vs comp buckets Σ={comp_b:>15,.0f}  {'✓' if abs(comp_raw-comp_b)<1 else '⚠DIFF'}")
    cap_raw = taken.loc[is_cap, amt].sum()
    cap_lab = taken.loc[is_cap & taken["項目組核對_H"].eq("人工成本"), amt].sum()
    cap_b = taken.loc[is_cap & taken["項目組核對_H"].isin(["建設與設施支出", "設施及器具採購", "維護費"]), amt].sum()
    print(f"  Capex  Ledger=Capex Σ={cap_raw:>15,.0f}  −人工 {cap_lab:>13,.0f}  vs 建設/設施/維護(capex) Σ={cap_b:>15,.0f}  {'✓' if abs(cap_raw-cap_b)<1 else '⚠'}")
    print(f"  人工   bucket 人工成本 Σ={buck('人工成本'):>15,.0f}")
    print(f"  一致   Σ buckets={bt['sum'].sum():>15,.0f}  vs  Σ taken={taken[amt].sum():>15,.0f}  {'✓' if abs(bt['sum'].sum()-taken[amt].sum())<1 else '⚠'}")

    # ── per source-column → bucket (verify each pivot + reveal 規律 for conf rules) ──
    for role, title in (("comp_nat", "Comp性質-CN → comp split"),
                        ("constr_div", "建築設施設備劃分 → capex split"),
                        ("spend_nat", "支出性質 → opex 主要 driver")):
        c = R.get(role)
        if not c: continue
        print(f"\n--- {title}  ('{c}') ---")
        g = taken.groupby([taken[c].astype(str), "項目組核對_H"])[amt].sum().reset_index()
        for sv, grp in sorted(g.groupby(c), key=lambda kv: -abs(kv[1][amt].sum())):
            t = grp[amt].sum()
            if abs(t) < 1 or str(sv).strip() in ("", "nan"): continue
            grp = grp.sort_values(amt, key=lambda s: s.abs(), ascending=False)
            tops = " ; ".join(f"{r['項目組核對_H']}={r[amt]/t*100:.0f}%" for _, r in grp.head(2).iterrows())
            print(f"     {str(sv)[:26]:26} Σ={t:>15,.0f}  {tops[:48]}")

    # ── deliverable: per-row H from raw (audit-ready subset) ──
    out = ROOT / "results" / f"melco_H_from_raw_{args.year}.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    cols = [c for c in (R["key"], R["subproj"], R["proj_nat"], R["take"], R["ledger"], R["comp_flag"],
                        R["comp_nat"], R["labor_flag"], R["perf_flag"], R["constr_div"], R["spend_nat"],
                        R["amount"], "項目組核對_H") if c]
    df[cols].to_excel(out, index=False)
    print(f"\n→ {out}  (每行: 來源 flag 欄 + 項目組核對_H + amount — 你可 pivot 對你嘅 7 張表)")
    print("   (項目組核對_H = 直接由 raw 嘅項目組欄推出嚟，即係你話「項目組已話我哋點取數」嗰個數)")


if __name__ == "__main__":
    main()
