#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_summary_tables.py — 由 prep_tableau feed 做「其他信息」嘅兩張數字表（報告 slide 42-45，
而家係 Tableau 截圖）。全部由底層數據 aggregate（報告只係 ref，永不用原報告做底）。

  4.1 金額匯總（slide 42）：於 2025 年發生嘅投資金額，by 範疇 × {2025計劃 / 2024計劃期後 /
       2023計劃期後}，每組 報告投資金額(調整前) + 潛在調整後投資金額(調整後)，+ 合計。
  4.2 設施vs活動（slide 43-45）：每個 plan-bucket 一張，by 範疇 × {設施建設(capex調整後) /
       活動舉辦(opex調整後) / 合計}。

「於2025發生」= 報告年(支出年)==25 = year_bucket ∈ {25, 25_24SY, 25_23SY}。
分組：博彩(vertical_label：娛樂場優化→設施設備優化) → 非博彩(ng_label，ng_code 序)，+ 小計 + 合計。

用法（Windows）：
    python scripts\\report\\build_summary_tables.py "tableau_combined_25.csv" --entity mgm
  出 mgm_金額匯總.xlsx（sheet：金額匯總 / 設施vs活動-2025計劃 / -2024期後 / -2023期後）+ console。
"""
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)

pd.set_option("display.max_columns", 30); pd.set_option("display.width", 200)

BUCKET = {"25": "2025年度投資計劃", "25_24SY": "2024年度計劃期後投資", "25_23SY": "2023年度計劃期後投資"}
BUCKET_ORDER = ["2025年度投資計劃", "2024年度計劃期後投資", "2023年度計劃期後投資"]
GORDER = {"博彩娛樂場優化": 0, "博彩娛樂場場地的優化": 0, "博彩設施設備優化": 1, "博彩設施及設備的優化": 1}


def _ngn(s):
    m = re.search(r"(\d+)", str(s)); return int(m.group(1)) if m else 99


def _load(feed: Path, entity: str) -> pd.DataFrame:
    df = pd.read_csv(feed, low_memory=False)
    if entity and "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity]
    df = df[df["dicj code"].astype(str).str.match(r"^項目\s*\d")]     # 丟 pseudo 碼
    df["_yb"] = df["year_bucket"].astype(str).str.strip()
    df["_bucket"] = df["_yb"].map(BUCKET)
    df = df[df["_bucket"].notna()].copy()                             # 只留「於2025發生」3 bucket
    df["_sub"] = df.apply(lambda r: r["vertical_label"] if r["ng_scope"] == "gaming" else r["ng_label"], axis=1)
    df["_scope"] = (df["ng_scope"] != "gaming").astype(int)           # 博彩=0 先
    df["_go"] = df["_sub"].map(lambda s: GORDER.get(s, 5))
    df["_ngn"] = df["ng_code"].map(_ngn)
    return df


def _order(agg):
    return agg.sort_values(["_scope", "_go", "_ngn", "_sub"])


def _emit(agg, valcols):
    """agg：每 (ng_scope,_sub) 一行 + valcols。→ 加 博彩/非博彩 小計 + 總計。"""
    ALL = ["範疇"] + valcols
    rows = []

    def agg_row(sub, label):
        r = {c: "" for c in ALL}; r["範疇"] = label
        for c in valcols:
            r[c] = round(pd.to_numeric(sub[c], errors="coerce").sum(), 1)
        return r
    for scope in [0, 1]:
        sc = _order(agg[agg["_scope"] == scope])
        if sc.empty:
            continue
        nm = "博彩項目" if scope == 0 else "非博彩項目"
        for _, row in sc.iterrows():
            r = {"範疇": row["_sub"]}
            for c in valcols:
                r[c] = round(float(row[c]), 1)
            rows.append(r)
        rows.append(agg_row(sc, f"{nm}小計"))
    rows.append(agg_row(_order(agg), "總計"))
    return pd.DataFrame(rows, columns=ALL)


def summary_amount(df) -> pd.DataFrame:
    """4.1 金額匯總：範疇 × bucket → 報告投資金額 / 潛在調整後投資金額 + 合計。"""
    g = df.groupby(["_scope", "_go", "_ngn", "_sub", "_bucket"], dropna=False).agg(
        報告=("調整前_萬", "sum"), 調整後=("調整後_萬", "sum")).reset_index()
    # pivot bucket → 兩個 measure
    base = g.groupby(["_scope", "_go", "_ngn", "_sub"], dropna=False)
    idx = base.size().reset_index()[["_scope", "_go", "_ngn", "_sub"]]
    out = idx.copy()
    valcols = []
    for bk in BUCKET_ORDER:
        sub = g[g["_bucket"] == bk].set_index(["_scope", "_go", "_ngn", "_sub"])
        for meas, lab in [("報告", "報告投資金額"), ("調整後", "潛在調整後投資金額")]:
            col = f"{bk}·{lab}"
            out[col] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
                lambda k: sub[meas].get(k, 0.0)).astype(float).round(1).values
            valcols.append(col)
    out["合計·報告投資金額"] = out[[f"{b}·報告投資金額" for b in BUCKET_ORDER]].sum(axis=1).round(1)
    out["合計·潛在調整後投資金額"] = out[[f"{b}·潛在調整後投資金額" for b in BUCKET_ORDER]].sum(axis=1).round(1)
    valcols += ["合計·報告投資金額", "合計·潛在調整後投資金額"]
    return _emit(out, valcols)


def facility_activity(df, bucket_label) -> pd.DataFrame:
    """4.2 設施vs活動（一個 bucket）：範疇 × 設施建設(capex調整後)/活動舉辦(opex調整後)/合計。"""
    d = df[df["_bucket"] == bucket_label]
    if d.empty:
        return pd.DataFrame()
    cap = d[d["final_capex_opex"] == "Capex"].groupby(["_scope", "_go", "_ngn", "_sub"])["調整後_萬"].sum()
    ope = d[d["final_capex_opex"] == "Opex"].groupby(["_scope", "_go", "_ngn", "_sub"])["調整後_萬"].sum()
    idx = d.groupby(["_scope", "_go", "_ngn", "_sub"]).size().reset_index()[["_scope", "_go", "_ngn", "_sub"]]
    out = idx.copy()
    out["設施建設/資本性支出"] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
        lambda k: cap.get(k, 0.0)).astype(float).round(1).values
    out["活動舉辦/營運性支出"] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
        lambda k: ope.get(k, 0.0)).astype(float).round(1).values
    out["合計"] = (out["設施建設/資本性支出"] + out["活動舉辦/營運性支出"]).round(1)
    return _emit(out, ["設施建設/資本性支出", "活動舉辦/營運性支出", "合計"])


def main():
    args = sys.argv[1:]
    entity = None
    if "--entity" in args:
        i = args.index("--entity"); entity = args[i + 1].lower(); del args[i:i + 2]
    if not args:
        print("俾 tableau feed csv 路徑"); return
    df = _load(Path(args[0]), entity)
    print(f"（{entity}：於2025發生 {len(df)} 行；bucket 分佈：{df['_bucket'].value_counts().to_dict()}）")

    out = Path(f"{entity or 'all'}_金額匯總.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        amt = summary_amount(df)
        amt.to_excel(xw, sheet_name="金額匯總", index=False)
        print(f"\n{'='*80}\n# 4.1 金額匯總（於2025發生 by 範疇）\n{amt.to_string(index=False)}")
        for bk in BUCKET_ORDER:
            fa = facility_activity(df, bk)
            if fa.empty:
                continue
            sn = "設施vs活動-" + bk.replace("年度", "").replace("投資", "")[:12]
            fa.to_excel(xw, sheet_name=sn[:31], index=False)
            print(f"\n{'-'*70}\n# 4.2 設施vs活動 — {bk}\n{fa.to_string(index=False)}")
    print(f"\n✓ 寫入 {out.resolve()}（開嚟同報告 slide 42-45 對）")


if __name__ == "__main__":
    main()
