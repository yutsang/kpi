#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_overview_tables.py — 概述數字（報告 ①②③ slide 10-40，而家係 Tableau chart 圖）由 feed rollup。
全部由底層數據 aggregate（報告只係 ref）。三個 cut：

 1) 投資概況總覽 overview_by_bucket（S10-14 / 19-26）：每 plan-bucket 一張，
    博彩/非博彩/總計 × 計劃投資金額 / 報告投資金額 / 潛在調整合計 / 調整後投資金額 / 完成率 / 調整後完成率。
 2) 潛在調整事項匯總 adjustment_bridge（S15-17）：7 canonical 調整類型 × {2025計劃/2024期後/2023期後/合計}。
 3) 主要發現摘要 finding_summary（S28-40）：7 類型 × 調整額合計 / 涉及項目數 / 主要涉及項目。

「於2025發生」slice + canonical 映射同 build_summary_tables / build_project_review_table 一致（互相 tie）。

用法（Windows）：
    python scripts\\report\\build_overview_tables.py "tableau_combined_25.csv" --entity mgm ^
        --qingdan "data\\投資項目清單\\3. MGM.…投资项目清单.xlsx"
  出 mgm_概述數字.xlsx + console。（一鍵見 make_report.py）
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)

import build_summary_tables as S
import build_project_review_table as B

BUCKET_PLANYR = {"2025年度投資計劃": 25, "2024年度計劃期後投資": 24, "2023年度計劃期後投資": 23}


def _plan_tot(plan, yr, gaming=None):
    d = (plan or {}).get(yr, {})
    return round(sum(v for (g, c), v in d.items() if gaming is None or g == gaming), 1)


def _rate(a, b):
    try:
        return round(a / b, 4) if b else None
    except Exception:
        return None


def overview_by_bucket(df, bucket, plan):
    """S10-14 / 19-26：博彩/非博彩/總計 × 計劃/報告/潛在調整/調整後/完成率/調整後完成率。"""
    d = df[df["_bucket"] == bucket]
    yr = BUCKET_PLANYR[bucket]
    if d.empty and not (plan or {}).get(yr):
        return pd.DataFrame()
    cols = ["範疇", "計劃投資金額", "報告投資金額", "潛在調整合計", "調整後投資金額",
            "投資計劃完成率", "潛在調整後投資計劃完成率"]
    rows = []
    tot = dict(計劃=0.0, 報告=0.0, 調整=0.0, 後=0.0)
    for scope, gaming, name in [(0, True, "博彩項目"), (1, False, "非博彩項目")]:
        sc = d[d["_scope"] == scope]
        pl = _plan_tot(plan, yr, gaming)
        if sc.empty and pl == 0:
            continue
        rep = round(sc["調整前_萬"].sum(), 1)
        adj = round(sc["調整_萬"].sum(), 1)
        aft = round(sc["調整後_萬"].sum(), 1)
        rows.append({"範疇": name, "計劃投資金額": pl, "報告投資金額": rep, "潛在調整合計": adj,
                     "調整後投資金額": aft, "投資計劃完成率": _rate(rep, pl),
                     "潛在調整後投資計劃完成率": _rate(aft, pl)})
        tot["計劃"] += pl; tot["報告"] += rep; tot["調整"] += adj; tot["後"] += aft
    rows.append({"範疇": "總計", "計劃投資金額": round(tot["計劃"], 1), "報告投資金額": round(tot["報告"], 1),
                 "潛在調整合計": round(tot["調整"], 1), "調整後投資金額": round(tot["後"], 1),
                 "投資計劃完成率": _rate(tot["報告"], tot["計劃"]),
                 "潛在調整後投資計劃完成率": _rate(tot["後"], tot["計劃"])})
    return pd.DataFrame(rows, columns=cols)


def adjustment_bridge(df):
    """S15-17：7 canonical 調整類型 × {2025計劃/2024期後/2023期後/合計}。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
    rows = []
    for adj in B.ADJ7:
        r = {"潛在調整事項": adj}
        for bk in S.BUCKET_ORDER:
            r[bk] = round(d[(d["_bucket"] == bk) & (d["_adj"] == adj)]["調整_萬"].sum(), 1)
        r["合計"] = round(sum(r[bk] for bk in S.BUCKET_ORDER), 1)
        rows.append(r)
    # 跨年及其他調整（唔喺報告 7 類，多數係期後嘅跨期/將往年計入本年）→ 令 bucket 合計對返概況潛在調整
    other = {"潛在調整事項": "跨年及其他調整"}
    for bk in S.BUCKET_ORDER:
        tot_bk = round(d[d["_bucket"] == bk]["調整_萬"].sum(), 1)
        other[bk] = round(tot_bk - sum(x[bk] for x in rows), 1)
    other["合計"] = round(sum(other[bk] for bk in S.BUCKET_ORDER), 1)
    if any(abs(other[bk]) > 0.05 for bk in S.BUCKET_ORDER):
        rows.append(other)
    tot = {"潛在調整事項": "合計"}
    for bk in S.BUCKET_ORDER + ["合計"]:
        tot[bk] = round(sum(x[bk] for x in rows), 1)
    rows.append(tot)
    return pd.DataFrame(rows, columns=["潛在調整事項"] + S.BUCKET_ORDER + ["合計"])


def finding_summary(df):
    """S28-40：每個 canonical 調整類型 → 調整額合計 / 涉及項目數 / 主要涉及項目(top3)。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
    rows = []
    for adj in B.ADJ7:
        sub = d[(d["_adj"] == adj) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        projs = sub.groupby("project")["調整_萬"].sum().abs().sort_values(ascending=False)
        rows.append({"潛在調整事項": adj, "調整額合計": round(sub["調整_萬"].sum(), 1),
                     "涉及項目數": int(sub["dicj code"].nunique()),
                     "主要涉及項目": "、".join(str(p) for p in projs.index[:3])})
    return pd.DataFrame(rows, columns=["潛在調整事項", "調整額合計", "涉及項目數", "主要涉及項目"])


def main():
    args = sys.argv[1:]
    entity = qingdan = None
    if "--entity" in args:
        i = args.index("--entity"); entity = args[i + 1].lower(); del args[i:i + 2]
    if "--qingdan" in args:
        i = args.index("--qingdan"); qingdan = args[i + 1]; del args[i:i + 2]
    if not args:
        print("俾 tableau feed csv 路徑（--qingdan 加計劃/完成率）"); return
    df = S._load(Path(args[0]), entity)
    plan = B.load_plan(Path(qingdan)) if qingdan else None
    print(f"（{entity}：於2025發生 {len(df)} 行）")

    out = Path(f"{entity or 'all'}_概述數字.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        for bk in S.BUCKET_ORDER:
            ov = overview_by_bucket(df, bk, plan)
            if ov.empty:
                continue
            ov.to_excel(xw, sheet_name=("概況-" + bk.replace("年度", "").replace("投資", ""))[:31], index=False)
            print(f"\n{'='*76}\n# 投資概況總覽 — {bk}\n{ov.to_string(index=False)}")
        ab = adjustment_bridge(df)
        ab.to_excel(xw, sheet_name="潛在調整事項匯總", index=False)
        print(f"\n{'='*76}\n# 潛在調整事項匯總（7 類型 × bucket）\n{ab.to_string(index=False)}")
        fs = finding_summary(df)
        fs.to_excel(xw, sheet_name="主要發現摘要", index=False)
        print(f"\n{'='*76}\n# 主要發現摘要\n{fs.to_string(index=False)}")
    print(f"\n✓ 寫入 {out.resolve()}（開嚟同報告 slide 10-40 對）")


if __name__ == "__main__":
    main()
