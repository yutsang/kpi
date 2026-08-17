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

sys.path.insert(0, str(Path(__file__).resolve().parent))
import feed_schema as FS

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
    FS.add_dims(df)                    # plan_year / spend_year / 範疇（一處派生）
    df["_sub"] = df["範疇"]
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
        rows.append({c: "" for c in ALL} | {"範疇": nm})   # section 標題行（對 scan p42）
        for _, row in sc.iterrows():
            r = {"範疇": row["_sub"]}
            for c in valcols:
                r[c] = round(float(row[c]), 1)
            rows.append(r)
        rows.append(agg_row(sc, f"{nm}小計"))
    rows.append(agg_row(_order(agg), "合計"))
    return pd.DataFrame(rows, columns=ALL)


def summary_amount(df) -> pd.DataFrame:
    """4.1 金額匯總：範疇 × bucket → 報告投資金額 / 潛在調整後投資金額 + 合計。"""
    # ★ 對 scan p42（項目組 2026-08-17）：每個年度出【報告 / 投資金額的潛在調整事項 / 潛在調整後】
    #   三欄；唔要「合計」欄組；最尾多一組「潛在調整後投資金額」拆設施建設／活動舉辦。
    g = df.groupby(["_scope", "_go", "_ngn", "_sub", "_bucket"], dropna=False).agg(
        報告=("調整前_萬", "sum"), 調整=("調整_萬", "sum"), 調整後=("調整後_萬", "sum")).reset_index()
    cap = df[df["final_capex_opex"] == "Capex"].groupby(
        ["_scope", "_go", "_ngn", "_sub"], dropna=False)["調整後_萬"].sum()
    ope = df[df["final_capex_opex"] == "Opex"].groupby(
        ["_scope", "_go", "_ngn", "_sub"], dropna=False)["調整後_萬"].sum()
    # pivot bucket → 兩個 measure
    base = g.groupby(["_scope", "_go", "_ngn", "_sub"], dropna=False)
    idx = base.size().reset_index()[["_scope", "_go", "_ngn", "_sub"]]
    out = idx.copy()
    valcols = []
    for bk in BUCKET_ORDER:
        sub = g[g["_bucket"] == bk].set_index(["_scope", "_go", "_ngn", "_sub"])
        for meas, lab in [("報告", "報告投資金額"), ("調整", "投資金額的潛在調整事項"),
                          ("調整後", "潛在調整後投資金額")]:
            col = f"{bk}·{lab}"
            out[col] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
                lambda k: sub[meas].get(k, 0.0)).astype(float).round(1).values
            valcols.append(col)
    _k = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index
    for lab, ser in (("設施建設/資本性支出", cap), ("活動舉辦/營運性支出", ope)):
        col = f"潛在調整後投資金額·{lab}"
        out[col] = [round(float(ser.get(k, 0.0)), 1) for k in _k]
        valcols.append(col)
    return _emit(out, valcols)          # 冇「合計」欄組（項目組 2026-08-17）


# 4.2（scan p43-45）欄組／欄名（逐字對報告）
FA_G1, FA_G2, FA_G3 = "設施建設/資本性支出", "活動舉辦/營運性支出", "合計"
FA_LEG = ["項目數量", "獲批的計劃投資金額", "報告投資金額",
          "設施建設及活動舉辦金額分攤", "投資金額小計", "投資金額的潛在調整事項",
          "潛在調整後投資金額"]
FA_TOT = ["項目數量", "獲批的計劃投資金額", "報告投資金額",
          "投資金額的潛在調整事項", "潛在調整後投資金額"]
# 報告投資金額點拆 capex/opex：優先用【項目組自己申報】嗰條欄（feed 帶落嚟），
#   搵唔到就當佢同我哋分類一樣（→「分攤」欄全 0，唔會影響 tie）。
DECLARED_COLS = ("declared_capex_opex", "capex_opex", "Capex_Opex", "Capex/Opex")


def _cap_key(v):
    return "Capex" if str(v).strip().lower().startswith("cap") else "Opex"


def facility_activity(df, bucket_label, plan_split=None) -> pd.DataFrame:
    """4.2 區分設施建設/活動舉辦（一個 bucket）→ 對 scan p43：3 個欄組 × 逐範疇。

      設施建設/資本性支出：項目數量 | 獲批的計劃投資金額 a¹ | 報告投資金額 b¹
                          | 設施建設及活動舉辦金額分攤 c¹ | 投資金額小計 d¹=b¹+c¹
                          | 投資金額的潛在調整事項 e¹ | 潛在調整後投資金額 f¹=d¹+e¹
      活動舉辦/營運性支出：同上（²）
      合計                ：項目數量 a=a¹+a² | 獲批 | 報告投資金額 b=b¹+b²
                          | 潛在調整事項 e=e¹+e² | 潛在調整後 f=b+e   ← 冇「分攤」欄（c¹+c²=0）

    ⚠ d/e/f 全部用【我哋嘅 final_capex_opex】→ 同 4.1 最後嗰組 tie 得返。
      b 用【項目組申報】嗰條 capex/opex 欄；c = d − b（即我哋重分類搬咗幾多）。
      plan_split = {(gaming, 碼): (計劃capex, 計劃opex)}，冇就 a 欄留空。"""
    d = df[df["_bucket"] == bucket_label].copy()
    if d.empty:
        return pd.DataFrame()
    idx = ["_scope", "_go", "_ngn", "_sub"]
    dec = next((c for c in DECLARED_COLS if c in d.columns), None)
    d["_ours"] = d["final_capex_opex"].map(_cap_key)
    d["_decl"] = d[dec].map(_cap_key) if dec else d["_ours"]
    if not dec:
        print("    ⚠ 4.2：feed 冇【項目組申報】嘅 capex/opex 欄（試過 "
              + "／".join(DECLARED_COLS) + "）→「設施建設及活動舉辦金額分攤」欄全 0")

    def _sum(col, leg, by):
        return d[d[by] == leg].groupby(idx, dropna=False)[col].sum()

    keys = d.groupby(idx, dropna=False).size().reset_index()[idx]
    out = keys.copy()
    ki = out.set_index(idx).index

    def _col(ser):
        return [round(float(ser.get(k, 0.0)), 1) for k in ki]
    npj = d.groupby(idx + ["_ours"], dropna=False)["dicj code"].nunique()
    valcols = []
    for leg, gname in (("Capex", FA_G1), ("Opex", FA_G2)):
        b = _sum("調整前_萬", leg, "_decl")
        dd = _sum("調整前_萬", leg, "_ours")
        e = _sum("調整_萬", leg, "_ours")
        cells = {
            "項目數量": [int(npj.get(k + (leg,), 0)) for k in ki],
            "獲批的計劃投資金額": _col(plan_split.get(leg, {}) if plan_split else {}),
            "報告投資金額": _col(b),
            "設施建設及活動舉辦金額分攤": [round(x - y, 1) for x, y in zip(_col(dd), _col(b))],
            "投資金額小計": _col(dd),
            "投資金額的潛在調整事項": _col(e),
        }
        cells["潛在調整後投資金額"] = [round(x + y, 1) for x, y in
                                     zip(cells["投資金額小計"], cells["投資金額的潛在調整事項"])]
        for c in FA_LEG:
            col = f"{gname}·{c}"
            out[col] = cells[c]; valcols.append(col)
    tot = {
        # scan p43 公式：合計項目數量 = a¹+a²（一個項目兩邊都有就數兩次），唔係 distinct
        "項目數量": [int(x + y) for x, y in
                    zip(out[f"{FA_G1}·項目數量"], out[f"{FA_G2}·項目數量"])],
        "獲批的計劃投資金額": [round(x + y, 1) for x, y in
                             zip(out[f"{FA_G1}·獲批的計劃投資金額"], out[f"{FA_G2}·獲批的計劃投資金額"])],
        "報告投資金額": _col(d.groupby(idx, dropna=False)["調整前_萬"].sum()),
        "投資金額的潛在調整事項": _col(d.groupby(idx, dropna=False)["調整_萬"].sum()),
    }
    tot["潛在調整後投資金額"] = [round(x + y, 1) for x, y in
                               zip(tot["報告投資金額"], tot["投資金額的潛在調整事項"])]
    for c in FA_TOT:
        col = f"{FA_G3}·{c}"
        out[col] = tot[c]; valcols.append(col)
    return _emit(out, valcols)


def fa_formula_row(cols):
    """4.2 表頭下面嗰行斜體公式（對 scan p43）。"""
    m = {f"{FA_G1}·獲批的計劃投資金額": "a¹", f"{FA_G1}·報告投資金額": "b¹",
         f"{FA_G1}·設施建設及活動舉辦金額分攤": "c¹", f"{FA_G1}·投資金額小計": "d¹=b¹+c¹",
         f"{FA_G1}·投資金額的潛在調整事項": "e¹", f"{FA_G1}·潛在調整後投資金額": "f¹=d¹+e¹",
         f"{FA_G2}·獲批的計劃投資金額": "a²", f"{FA_G2}·報告投資金額": "b²",
         f"{FA_G2}·設施建設及活動舉辦金額分攤": "c²", f"{FA_G2}·投資金額小計": "d²=b²+c²",
         f"{FA_G2}·投資金額的潛在調整事項": "e²", f"{FA_G2}·潛在調整後投資金額": "f²=d²+e²",
         f"{FA_G3}·獲批的計劃投資金額": "a=a¹+a²", f"{FA_G3}·報告投資金額": "b=b¹+b²",
         f"{FA_G3}·投資金額的潛在調整事項": "e=e¹+e²", f"{FA_G3}·潛在調整後投資金額": "f=b+e"}
    return [m.get(str(c), "") for c in cols]


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
