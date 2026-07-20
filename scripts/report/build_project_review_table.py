#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_project_review_table.py — 由 prep_tableau feed 生成「單個項目審查結果匯總表」數據
（報告 slide 46-63 而家係 Tableau 截圖，呢個 native 砌返 + 對數）。

每個 報告年(25/24/23) 出一張：
  行 = 項目（按 dicj code = 項目N 合併，-OPEX/-製作期 等細拆加埋）
  欄 = 項目編號 | 項目名稱 | 範疇 | 博彩/非博彩 | 申報投資金額(調整前_萬)
       | <各 canonical 調整類型…> | 潛在調整合計(調整_萬) | 潛在調整後投資金額(調整後_萬)
  （完成率：feed 冇計劃金額 → 暫缺，待另接 plan 源）
排序：非博彩 → 博彩，範疇(ng_code)，項目編號。加合計行。

用法（Windows，kpi-main 底下）：
    python scripts\\report\\build_project_review_table.py "tableau_combined_25.csv" --entity mgm
    # 輸出 mgm_項目審查匯總.xlsx（3 個 sheet：25/24/23）+ console 預覽
"""
import re
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("✗ 未裝 pandas → pip install pandas openpyxl"); sys.exit(1)

pd.set_option("display.max_columns", 40)
pd.set_option("display.width", 220)

# 調整一級 → canonical 報告類別（合併變體）。之後對到 Tableau 表再微調。
CANON = {
    "其他日常營運支出調整": "其他日常營運支出調整",
    "其他日常營運支出調整調整": "其他日常營運支出調整",
    "酒店客房改造支出": "酒店客房改造支出",
    "超過可計入範圍的内部資源支出": "超出可計入範圍的內部資源支出",
    "超出可計入範圍的內部資源支出": "超出可計入範圍的內部資源支出",
    "不符合“吸引外國客源”定義的相關投資支出": "不符合“吸引外國客源”定義的相關投資支出",
    "不符合吸引外國客源": "不符合“吸引外國客源”定義的相關投資支出",
    "未完全實現投資目的的投資支出": "未完全實現投資目的的投資支出",
    "未能實現投資目的的投資支出": "未完全實現投資目的的投資支出",
    "計入報告投資金額的高管及一般支持人員人工成本": "高管及一般支持人員人工成本",
    "一般性支持部門的人工成本": "高管及一般支持人員人工成本",
    "一般支持性部門的人工成本": "高管及一般支持人員人工成本",
    "計劃獲批前的投資": "計劃獲批前的投資",
    "將2024年的支出計入2025年報告投資金額": "將往年支出計入本年報告投資金額",
    "2024年度計劃與2023年度計劃期後調整之間的報告投資金額跨期調整": "年度計劃期後調整之間的跨期調整",
}
# canonical 欄順序（報告版式；缺嘅自動略）
CANON_ORDER = [
    "高管及一般支持人員人工成本", "其他日常營運支出調整", "超出可計入範圍的內部資源支出",
    "酒店客房改造支出", "不符合“吸引外國客源”定義的相關投資支出", "未完全實現投資目的的投資支出",
    "計劃獲批前的投資", "將往年支出計入本年報告投資金額", "年度計劃期後調整之間的跨期調整",
]


def _norm(c) -> str:
    """項目編號正規化：去空白、去『項目』、去前導零 → 對 清單 承批公司項目序號。"""
    s = re.sub(r"\s+", "", str(c if c is not None else ""))
    s = re.sub(r"^項目", "", s)
    m = re.match(r"^0*(\d+(?:\.\d+)?)$", s)
    return (m.group(1) if m else s).lower()


# 每個報告年 → 清單「計劃投資金額」欄 header 嘅正則（database sheet；萬澳門元）
_PLAN_RE = {
    25: re.compile(r"2025.*預計投資金額.*合計|2025.*預計投資金額（萬"),
    24: re.compile(r"2024.*預計投資金額.*合計|2024.*預計投資金額（萬"),
    23: re.compile(r"2023.*預計投資金額"),
}


def load_plan(path: Path, log=print) -> dict:
    """{報告年: {正規化項目編號: 計劃投資金額_萬}} — 由投資項目清單 database sheet 抽。"""
    import openpyxl
    out = {25: {}, 24: {}, 23: {}}
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        log(f"  ⚠ 清單開唔到 {path}: {e}"); return out
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            rows.append(r)
            if i > 2500:
                break
        # 揾表頭行（含『承批公司項目序號』）
        hdr_r = code_c = None
        for ri in range(min(14, len(rows))):
            for ci, v in enumerate(rows[ri] or []):
                if "承批公司項目序號" in ("" if v is None else str(v)):
                    hdr_r, code_c = ri, ci; break
            if hdr_r is not None:
                break
        if hdr_r is None:
            continue
        hdr = [("" if v is None else str(v).replace("\n", "")) for v in rows[hdr_r]]
        # 每年計劃欄
        plan_c = {}
        for yr, rgx in _PLAN_RE.items():
            for ci, h in enumerate(hdr):
                if rgx.search(h):
                    plan_c[yr] = ci; break
        if not plan_c:
            continue
        log(f"  清單 sheet {sn!r}: 計劃欄 " +
            ", ".join(f"{yr}→{hdr[ci][:20]!r}" for yr, ci in plan_c.items()))
        for ri in range(hdr_r + 1, len(rows)):
            row = rows[ri]
            code = _norm(row[code_c]) if code_c < len(row) else ""
            if not code:
                continue
            for yr, ci in plan_c.items():
                if ci < len(row):
                    try:
                        val = float(row[ci])
                        out[yr].setdefault(code, val)
                    except (TypeError, ValueError):
                        pass
        break   # database sheet 夠，唔使掃其他
    return out


def build_year(df: pd.DataFrame, year: int, plan: dict | None = None) -> pd.DataFrame:
    d = df[df["報告年"] == year].copy()
    if d.empty:
        return pd.DataFrame()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    # 逐項目 base（申報/調整/調整後）
    base = d.groupby("dicj code").agg(
        項目名稱=("project", "first"),
        範疇=("ng_label", "first"),
        類型=("ng_scope", "first"),
        _ngcode=("ng_code", "first"),
        申報投資金額=("調整前_萬", "sum"),
        潛在調整合計=("調整_萬", "sum"),
        潛在調整後投資金額=("調整後_萬", "sum"),
    )
    # 各 canonical 調整類型 breakdown（只用有調整一級嘅行）
    adj = d[d["調整一級"].notna()]
    pv = adj.pivot_table(index="dicj code", columns="_adj",
                         values="調整_萬", aggfunc="sum", fill_value=0)
    tab = base.join(pv).fillna(0)
    # 欄順序：識別欄 + canonical 調整欄（存在嘅）+ 合計/後
    adj_cols = [c for c in CANON_ORDER if c in tab.columns]
    other = [c for c in pv.columns if c not in CANON_ORDER]   # 未 map 到嘅（要通知）
    cols = (["項目名稱", "範疇", "類型", "申報投資金額"] + adj_cols + other
            + ["潛在調整合計", "潛在調整後投資金額"])
    tab = tab.reset_index().rename(columns={"dicj code": "項目編號"})
    # 排序：非博彩→博彩，ng_code，項目編號（項目N 數字序）
    def _pn(s):
        m = re.search(r"(\d+(?:\.\d+)?)", str(s)); return float(m.group(1)) if m else 9e9
    tab["_scope_ord"] = (tab["類型"] == "gaming").astype(int)
    tab["_pn"] = tab["項目編號"].map(_pn)
    tab = tab.sort_values(["_scope_ord", "_ngcode", "_pn"]).drop(columns=["_scope_ord", "_ngcode", "_pn"])
    num = ["申報投資金額"] + adj_cols + other + ["潛在調整合計", "潛在調整後投資金額"]
    for c in num:
        tab[c] = tab[c].round(1)
    plan_cols = []
    if plan:
        tab["計劃投資金額"] = pd.to_numeric(
            tab["項目編號"].map(lambda x: plan.get(_norm(x))), errors="coerce").round(1)
        plan_cols = ["計劃投資金額"]
    tab = tab[["項目編號", "項目名稱", "範疇", "類型"] + num + plan_cols]
    # 合計行（num + 計劃 加總）
    total = {c: "" for c in tab.columns}
    total["項目編號"] = "合計"
    for c in num + plan_cols:
        total[c] = round(pd.to_numeric(tab[c], errors="coerce").sum(), 1)
    tab = pd.concat([tab, pd.DataFrame([total])], ignore_index=True)
    # 完成率 = 潛在調整後投資金額 ÷ 計劃投資金額（逐行 + 合計自動）
    if plan:
        rate = (pd.to_numeric(tab["潛在調整後投資金額"], errors="coerce")
                / pd.to_numeric(tab["計劃投資金額"], errors="coerce"))
        tab["完成率"] = rate.replace([float("inf"), float("-inf")], pd.NA).round(4)
    return tab, other


def main():
    args = sys.argv[1:]
    entity = qingdan = None
    if "--entity" in args:
        i = args.index("--entity"); entity = args[i + 1].lower(); del args[i:i + 2]
    if "--qingdan" in args:
        i = args.index("--qingdan"); qingdan = args[i + 1]; del args[i:i + 2]
    if not args:
        print("俾 tableau feed csv 路徑（--qingdan <清單.xlsx> 加計劃金額+完成率）"); return
    src = Path(args[0])
    df = pd.read_csv(src, low_memory=False)
    if entity and "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity]
    df["報告年"] = pd.to_numeric(df["報告年"], errors="coerce")

    plan = None
    if qingdan:
        print(f"── 讀清單計劃金額: {qingdan}")
        plan = load_plan(Path(qingdan))
        for yr in (25, 24, 23):
            print(f"    報告年{yr}: {len(plan.get(yr, {}))} 個項目有計劃金額")

    out = Path(f"{entity or 'all'}_項目審查匯總.xlsx")
    unmapped_all = set()
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for yr in (25, 24, 23):
            res = build_year(df, yr, plan.get(yr) if plan else None)
            if isinstance(res, tuple):
                tab, other = res
            else:
                continue
            if tab.empty:
                continue
            unmapped_all |= set(other)
            tab.to_excel(xw, sheet_name=f"報告年{yr}", index=False)
            print(f"\n{'='*80}\n# 報告年 {yr}：{len(tab)-1} 項目")
            # 預覽（去長項目名，慳位）
            prev = tab.copy()
            prev["項目名稱"] = prev["項目名稱"].astype(str).str.slice(0, 14)
            print(prev.head(28).to_string(index=False))
            if len(tab) > 29:
                print(f"    …（另 {len(tab)-29} 行，見 xlsx）")
    if unmapped_all:
        print(f"\n⚠ 未 map 到 canonical 嘅 調整一級（我要加入 CANON）: {sorted(unmapped_all)}")
    print(f"\n✓ 寫入 {out.resolve()}（開嚟同 Tableau 嗰張對，話我要改咩）")


if __name__ == "__main__":
    main()
