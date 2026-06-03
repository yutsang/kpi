"""Melco 橫向(H) 對數:用項目組嘅取數邏輯(pivot formula)算出每行嘅 their_H,join 我哋
kpi_report 嘅 our_H,出 row-level 差異 + summary。

項目組取數 cascade (first-match-wins per 取數=Y row):
  1 comp        : KP識別Comp=Y                                  -> by Comp性質-CN
  2 labor       : KP識別人工 非空非0                              -> H_LABOR
  3 performer   : KP识别表演演出=Y & Comp=N/A/空 & Opex & 人工空  -> H_PERFORMER
  4 capex       : Ledger Type=Capex & 人工空                      -> by 建築設施設備劃分
  5 sponsorship : ledger_account contains Sponsorship            -> H_SPONSORSHIP
  6 marketing   : 支出性質=營銷費用                               -> H_ADVERTISING
  7 professional: 支出性質=專業服務費                             -> H_PROFESSIONAL
  8 license     : ledger_account contains License/Trademark      -> H_LICENSE
  9 residual    : 支出性質 map

Run (Windows):
  python scripts/reconcile_melco.py --year 25
  python scripts/reconcile_melco.py --year 24 --file melco_2024.xlsx
Output: results/melco_<year>_reconcile.tsv (row-level) + summary printed.
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq

ROOT = Path(__file__).resolve().parent.parent

COMP_TO_H = {"場地租借": "H_VENUE", "客房支出": "H_HOTEL_ROOM", "食品飲料支出": "H_FNB",
             "贈票支出": "H_COMP_TICKET", "演唱會門票支出": "H_COMP_TICKET", "其他": "H_COMP_OTHER"}
CAPEX_TO_H = {"建設與設施支出": "H_CONSTRUCTION", "建設與設施支出_水舞間": "H_CONSTRUCTION",
              "設施及器具採購": "H_EQUIP", "設施及器具採購_水舞間": "H_EQUIP", "一般維護費": "H_MAINTENANCE"}
SPEND_TO_H = {"營銷費用": "H_ADVERTISING", "表演者費用": "H_PERFORMER", "職工薪酬": "H_LABOR",
              "維修與保養費用": "H_MAINTENANCE", "水電費": "H_OTHER", "專業服務費": "H_PROFESSIONAL",
              "管理費": "H_OTHER", "Cost Recovery": "H_PERFORMER", "內部資源（Comp）": "H_COMP_OTHER",
              "外部資源（Comp）": "H_COMP_OTHER", "影視製作及租賃成本": "H_LEASE"}
BLANKS = {"", "nan", "none", "<na>", "0", "0.0", "n/a", "na"}


def col(df, *names):
    low = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n in df.columns: return n
        if str(n).strip().lower() in low: return low[str(n).strip().lower()]
    # substring fallback
    for n in names:
        for c in df.columns:
            if str(n).strip().lower() in str(c).strip().lower(): return c
    return None


def blank(v):
    return str(v).strip().lower() in BLANKS


# 24 performer 取數 whitelist (支出性质-mapping ∈ 呢啲 or blank);25 唔需要
PERF_WL24 = {"cost recovery", "保險費用", "表演者費用", "差旅費", "第三方公司支出",
             "管理費", "其他", "影視製作及租賃成本", "折扣費用", "製作費用", ""}


def their_h(r, C, yr):
    g = lambda k: r.get(C.get(k)) if C.get(k) else None
    comp_ind, comp_nat = g("comp_ind"), str(g("comp_nat") or "").strip()
    payroll = g("payroll"); perf = str(g("perf") or "").strip().upper()
    cap = str(g("capex_opex") or ""); spend = str(g("spend") or "").strip()
    ledger = str(g("ledger") or ""); capx = str(g("capex_split") or "").strip()
    # 1 comp — 25: KP識別Comp=Y ; 24: Comp性質-CN 非 N/A 非空 (項目組 24 用 Comp性質分類)
    comp_is = (str(comp_ind or "").strip().upper() == "Y") if yr == "25" \
        else (not blank(comp_nat) and "N/A" not in comp_nat.upper())
    if comp_is:
        return COMP_TO_H.get(comp_nat, "H_COMP_OTHER")
    # 2 labor — 只係特定值至算人工 (Y / Y,staff cost分攤 / Staff cost come from Project Team),
    #            唔係「非空就算」(個欄好多行有非空非人工嘅值)
    pay = str(payroll or "").strip()
    if "staff cost" in pay.lower() or pay.upper() == "Y" or pay.upper().startswith("Y,"):
        return "H_LABOR"
    # 3 performer — 25: comp=N/A/空 & Opex ; 24: comp_nat=N/A/空 & 支出性质∈whitelist
    if perf == "Y" and blank(payroll):
        if yr == "25":
            ok = (blank(comp_ind) or "N/A" in str(comp_ind).upper()) and "opex" in cap.lower()
        else:
            ok = ("N/A" in comp_nat.upper() or blank(comp_nat)) and spend.lower() in PERF_WL24
        if ok:
            return "H_PERFORMER"
    # 4 capex (建設/設施/維護) — 劃分空白/篩唔出 → 一般維護費 (項目組 note)
    if "capex" in cap.lower() and blank(payroll):
        return CAPEX_TO_H.get(capx, "H_MAINTENANCE")
    # 5 sponsorship (within 營銷)
    if "sponsorship" in ledger.lower():
        return "H_SPONSORSHIP"
    # 6 marketing / 7 professional / 8 license / 9 residual
    if spend == "營銷費用":
        return "H_ADVERTISING"
    if spend == "專業服務費":
        return "H_PROFESSIONAL"
    if "license" in ledger.lower() or "trademark" in ledger.lower():
        return "H_LICENSE"
    return SPEND_TO_H.get(spend, "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="25")
    ap.add_argument("--file", default=None)
    ap.add_argument("--sheet", default="Data")
    args = ap.parse_args()
    yr = str(args.year)

    fn = args.file or (f"melco_2025.xlsx" if yr == "25" else "melco_2024.xlsx")
    fp = next((p for p in [Path(fn), ROOT / "data/melco/raw" / fn, ROOT / fn,
                           ROOT / f"melco_audit_{yr}.xlsx"] if p.exists()), None)
    if not fp:
        print(f"X audit file not found (tried {fn}, data/melco/raw/{fn}, melco_audit_{yr}.xlsx)"); return
    print(f"reading {fp}  sheet={args.sheet}")
    a = pd.read_excel(fp, sheet_name=args.sheet, dtype=object)
    C = {"take": col(a, "取數標籤", "取數標簽", "取數标签"),
         "comp_ind": col(a, "KP識別Comp", "KP识别Comp"),
         "comp_nat": col(a, "Comp性質-CN（N/A為Net off及不適用），待確認kp識別，客戶未識別部分）", "Comp性質-CN"),
         "payroll": col(a, "KP識別人工（Staff cost come from Project Team需Capex提供）", "KP識別人工"),
         "perf": col(a, "KP识别表演演出", "KP識別表演演出"),
         "capex_opex": col(a, "Ledger Type", "Journal Type"),
         "spend": col(a, "支出性質", "支出性质-mapping"),
         "ledger": col(a, "ledger_account"),
         "capex_split": col(a, "建築設施設備劃分"),
         "amount": col(a, "Amount - Amended", "Amount Amended", "Amount-Amended"),
         "uid": col(a, "KPMG唯一識別碼")}
    print("  detected:", {k: v for k, v in C.items() if v})
    miss = [k for k in ("take", "comp_nat", "payroll", "spend", "ledger", "uid", "amount") if not C[k]]
    if miss:
        print(f"  ! missing cols: {miss}  — columns are: {list(a.columns)[:30]}")

    # 取數 gate — 25 用 取數標籤=Y;24 嘅 pivot 用 Report Years=全部,冇 取數標籤 gate → 唔 filter
    if yr == "25" and C["take"]:
        take = a[C["take"]].astype(str).str.strip().str.upper().isin(["Y"])
        a = a[take].copy()
        print(f"  25: 取數標籤=Y gate -> {len(a):,} rows")
    else:
        print(f"  24: no 取數標籤 gate (Report Years=all) -> {len(a):,} rows")
    a["their_H"] = a.apply(lambda r: their_h(r, C, yr), axis=1)
    a["_amt"] = pd.to_numeric(a[C["amount"]], errors="coerce").fillna(0) if C["amount"] else 0.0
    a["_uid"] = a[C["uid"]].astype(str) if C["uid"] else ""

    # join our H from kpi_report
    pqf = ROOT / "data/melco/output/company_5_kpi_report.parquet"
    our = None
    if pqf.exists():
        k = pq.read_table(pqf).replace_schema_metadata(None).to_pandas()
        ucol = col(k, "KPMG唯一識別碼", "unique_id")
        if ucol and "horizontal_label" in k.columns:
            our = dict(zip(k[ucol].astype(str), k["horizontal_label"].astype(str)))
    a["our_H"] = a["_uid"].map(our) if our else ""
    HLBL = {"H_VENUE": "活動場地", "H_HOTEL_ROOM": "酒店客房", "H_FNB": "餐飲", "H_COMP_TICKET": "贈票支出",
            "H_COMP_OTHER": "Comp其他", "H_LABOR": "人工成本", "H_PERFORMER": "合約成本", "H_CONSTRUCTION": "建設與設施支出",
            "H_EQUIP": "設施及器具採購", "H_MAINTENANCE": "維護費", "H_SPONSORSHIP": "贊助費", "H_ADVERTISING": "廣告及推廣",
            "H_PROFESSIONAL": "專業服務費", "H_LICENSE": "授權費", "H_LEASE": "租賃費", "H_OTHER": "其他"}
    a["their_H_label"] = a["their_H"].map(HLBL).fillna(a["their_H"])
    a["agree"] = (a["their_H_label"].fillna("") == a["our_H"].fillna("")) | (a["our_H"].fillna("") == "")

    # 診斷:仲有邊啲行 their_H 空白(formula 覆蓋唔到) — 睇 Ledger Type x 支出性質
    blanks = a[a["their_H_label"].fillna("") == ""]
    if len(blanks):
        print(f"\n  ⚠ their_H 空白: {len(blanks):,} 行, Σ {blanks['_amt'].sum():,.0f} — 頭 12 個 (Ledger Type | 支出性質):")
        bd = blanks.groupby([blanks[C["capex_opex"]].astype(str) if C["capex_opex"] else "",
                             blanks[C["spend"]].astype(str) if C["spend"] else ""])["_amt"].agg(["size", "sum"])
        for (lt, sp), r in bd.sort_values("sum", ascending=False).head(12).iterrows():
            print(f"      {str(lt)[:14]:14} | {str(sp)[:24]:24} {int(r['size']):>6}  {r['sum']:>14,.0f}")

    out = ROOT / "results" / f"melco_{yr}_reconcile_diff.tsv"
    out.parent.mkdir(parents=True, exist_ok=True)
    keep = [c for c in [C["uid"], C["spend"], C["comp_nat"], C["capex_split"],
                        "their_H_label", "our_H", "_amt"] if c]
    diff = a[~a["agree"]]  # 只寫唔 agree 嘅行 (慳走一致嗰啲)
    diff[keep].sort_values("_amt", ascending=False).to_csv(out, sep="\t", index=False)
    print(f"\nwrote {out.relative_to(ROOT)}  (只列 their≠our: {len(diff):,} / {len(a):,} 行)")
    # summary: their_H total vs our_H total
    print("\n  their_H bucket Σ (萬):")
    g = a.groupby("their_H_label")["_amt"].agg(["size", "sum"]).sort_values("sum", ascending=False)
    for lbl, r in g.iterrows():
        print(f"    {str(lbl):14} {int(r['size']):>6}  {r['sum']:>16,.0f}")
    if our:
        dis = a[~a["agree"]]
        print(f"\n  唔 agree (their≠our, 兩邊都有值): {len(dis):,} 行, Σ {dis['_amt'].sum():,.0f}")
        print("  最大分歧 (their_H -> our_H):")
        for (th, ou), sub in sorted(dis.groupby(["their_H_label", "our_H"]),
                                    key=lambda kv: -kv[1]["_amt"].sum())[:15]:
            print(f"    {str(th):12} -> {str(ou):12}  {len(sub):>5}行  {sub['_amt'].sum():>14,.0f}")


if __name__ == "__main__":
    main()
