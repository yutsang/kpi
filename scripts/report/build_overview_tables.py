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


def overview_by_bucket(df, bucket, plan, category=None):
    """整體投資概況（S10-14 = 2025計劃有計劃/完成率；S19-26 = 期後冇計劃、多潛在調整金額欄）。
    逐範疇 + 博彩/非博彩小計 + 總計。欄跟報告 IMG_0104/0105：
      2025計劃：項目數量 | 獲批的計劃投資金額 | 報告投資金額 | 完成率 | 潛在調整後投資金額 | 完成率 | 設施建設 | 活動舉辦
      期後    ：項目數量 | 報告投資金額 | 潛在調整金額 | 潛在調整後投資金額 | 設施建設 | 活動舉辦
    ⚠ 項目數量/逐範疇計劃 = feed 出現嘅碼；零申報項目未計入（報告項目數量含零申報）→ 小計/總計計劃用清單準數。"""
    d = df[df["_bucket"] == bucket]
    if d.empty:
        return pd.DataFrame()
    yr = BUCKET_PLANYR[bucket]
    is_py = (bucket == "2025年度投資計劃")
    idx = ["_scope", "_go", "_ngn", "_sub"]
    g = d.groupby(idx, dropna=False).agg(
        項目數量=("dicj code", "nunique"), 報告=("調整前_萬", "sum"),
        調整=("調整_萬", "sum"), 後=("調整後_萬", "sum")).reset_index()
    cap = d[d["final_capex_opex"] == "Capex"].groupby(idx)["調整後_萬"].sum().rename("設施")
    ope = d[d["final_capex_opex"] == "Opex"].groupby(idx)["調整後_萬"].sum().rename("活動")
    g = g.merge(cap.reset_index(), on=idx, how="left").merge(ope.reset_index(), on=idx, how="left")
    g[["設施", "活動"]] = g[["設施", "活動"]].fillna(0.0)
    g = g.sort_values(idx)

    # 項目數量：2025計劃表要數【計劃項目】（清單），唔係 feed 出現嘅碼 —— 否則同「已實施/未實施」
    #   兩行對唔上（項目組 2026-08-15：總計 84 但 79+10≠84）。期後表冇計劃概念，照用 feed。
    plan_by_sub, n_by_sub, n_by_scope = {}, {}, {}
    if is_py and plan:
        cat = category or {}
        sub_of = {}       # (gaming,碼)→範疇；博彩/非博彩共用項目N 都要留（唔可以 drop 淨 code）
        d2sub = {}        # 項目性質(D)→set(範疇)：由 feed 有嘅項目學，用嚟派零投資項目計劃
        for _, r in d.drop_duplicates(["ng_scope", "dicj code"]).iterrows():
            key = (r["ng_scope"] == "gaming", B._norm(r["dicj code"]))
            sub_of[key] = r["_sub"]
            dv = cat.get(key)
            if dv:
                d2sub.setdefault(str(dv), set()).add(r["_sub"])
        d2sub1 = {k: next(iter(v)) for k, v in d2sub.items() if len(v) == 1}  # 唯一對應先用
        miss = 0
        for (gm, code), v in (plan.get(yr, {}) or {}).items():
            sub = sub_of.get((gm, code))
            if sub is None:      # 零投資項目：feed 冇行 → 用 D→範疇 學到嘅 map 派返
                sub = d2sub1.get(str(cat.get((gm, code), "")))
            if sub is not None:
                plan_by_sub[sub] = plan_by_sub.get(sub, 0.0) + v
                if v > 0:      # 計劃金額 0 嘅唔算「獲批開展」（清單有大量 0 行）
                    n_by_sub[sub] = n_by_sub.get(sub, 0) + 1
                    n_by_scope[0 if gm else 1] = n_by_scope.get(0 if gm else 1, 0) + 1
            elif v:
                miss += 1
        if miss:
            print(f"    ⚠ overview {bucket}：{miss} 個有計劃項目派唔到範疇（缺項目性質對應）")

    def mk(name, cnt, pl, rep, adj, aft, fac, act):
        r = {"範疇": name, "項目數量": int(cnt), "報告投資金額": round(rep, 1),
             "潛在調整金額": round(adj, 1), "潛在調整後投資金額": round(aft, 1),
             "設施建設/資本性支出": round(fac, 1), "活動舉辦/營運性支出": round(act, 1)}
        if is_py:
            r["獲批的計劃投資金額"] = round(pl, 1)
            r["投資計劃完成率"] = _rate(rep, pl)
            r["潛在調整後投資計劃完成率"] = _rate(aft, pl)
        return r

    rows, n_tot = [], 0
    for scope in [0, 1]:
        sc = g[g["_scope"] == scope]
        name = "博彩項目" if scope == 0 else "非博彩項目"
        if sc.empty:
            # ⚠ 報告就算該 scope 全 0 都會出「博彩項目」section + 小計行（全部「-」），
            #   唔會成組唔見（項目組 2026-08-17）。
            rows.append({"範疇": name})
            rows.append(mk(f"{name}小計", 0, _plan_tot(plan, yr, scope == 0), 0, 0, 0, 0, 0))
            continue
        rows.append({"範疇": name})     # section 標題行（跟報告 IMG_0105：博彩項目 / 非博彩項目）
        for _, row in sc.iterrows():
            rows.append(mk(row["_sub"], n_by_sub.get(row["_sub"], row["項目數量"]),
                           plan_by_sub.get(row["_sub"], 0.0),
                           row["報告"], row["調整"], row["後"], row["設施"], row["活動"]))
        # ⚠ 項目數量：小計/總計要用【去重】distinct，唔可以逐範疇加總 ——
        #   一個項目跨兩個範疇會被計兩次（user 2026-08-12 對數揭到）
        n_sc = n_by_scope.get(scope) or d[d["_scope"] == scope]["dicj code"].nunique()
        n_tot += n_sc                # ⚠ 總計唔可以全表 nunique：博彩「項目19」同非博彩「項目19」
                                     #   係兩個唔同項目（撞號），全表去重會少計 → 總計 = 兩個 scope 之和
        rows.append(mk(f"{name}小計", n_sc, _plan_tot(plan, yr, scope == 0),
                       sc["報告"].sum(), sc["調整"].sum(), sc["後"].sum(), sc["設施"].sum(), sc["活動"].sum()))
    rows.append(mk("總計", n_tot, _plan_tot(plan, yr, None),
                   g["報告"].sum(), g["調整"].sum(), g["後"].sum(), g["設施"].sum(), g["活動"].sum()))
    if is_py:
        cols = ["範疇", "項目數量", "獲批的計劃投資金額", "報告投資金額", "投資計劃完成率",
                "潛在調整後投資金額", "潛在調整後投資計劃完成率", "設施建設/資本性支出", "活動舉辦/營運性支出"]
    else:
        cols = ["範疇", "項目數量", "報告投資金額", "潛在調整金額", "潛在調整後投資金額",
                "設施建設/資本性支出", "活動舉辦/營運性支出"]
    return pd.DataFrame(rows)[cols]


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


def zero_investment_summary(df, plan, cat, narr, ent_up="MGM"):
    """報告概述尾段：2025計劃申報投資為零嘅非博彩項目（原計劃有金額、2025 實際=0）分類。
    → (n, 總計劃_萬, groups)；groups={跨年/內部研究/取消: [範疇…]}。分類靠 期後有支出=跨年、
    項目狀況含取消=取消、其餘=內部研究（best-effort，項目組口徑為準）。"""
    if not plan or not plan.get(25):
        return None
    from collections import Counter

    def _key(r):
        return (str(r["ng_scope"]) == "gaming", B._norm(r["dicj code"]))

    def _has(sub):
        return {_key(r) for _, r in sub.iterrows()
                if abs(pd.to_numeric(r.get("調整前_萬", 0), errors="coerce") or 0) > 0.05}

    spent25 = _has(df[df["_bucket"] == "2025年度投資計劃"])
    post = _has(df[df["_bucket"].isin(["2024年度計劃期後投資", "2023年度計劃期後投資"])])
    groups = {"跨年": [], "內部研究": [], "取消": []}
    tot = 0.0
    for (gm, code), ev in (plan.get(25) or {}).items():
        if gm or (ev or 0) <= 0.05 or (gm, code) in spent25:
            continue        # 只計非博彩、原計劃>0、2025計劃無實際支出
        rec = (narr or {}).get((gm, code)) or (narr or {}).get((not gm, code)) or {}
        status = str(rec.get("項目狀況", ""))
        sub = (cat or {}).get((gm, code), "") or "其他"
        kind = "跨年" if (gm, code) in post else ("取消" if "取消" in status else "內部研究")
        groups[kind].append(sub)
        tot += ev
    n = sum(len(v) for v in groups.values())
    if n == 0:
        return None
    return n, round(tot, 1), groups


def zero_investment_text(zi, ent_up="MGM"):
    """zero_investment_summary → 報告式段落文字。"""
    if not zi:
        return None
    from collections import Counter
    n, tot, groups = zi
    lines = [f"{ent_up}有{n}個非博彩項目，原計劃項目於2025年度投資執行報告中申報的投資支出為零"
             f"（原計劃金額約{tot:,.0f}萬澳門元），主要包括："]
    lab = {"跨年": "個項目的2025年支出被申報為2023／2024年度計劃在2025年的期後投資金額（跨年項目）",
           "內部研究": "個項目仍處於內部研究、重新規劃或選址階段，未發生實際支出",
           "取消": "個項目已取消"}
    for kind in ("跨年", "內部研究", "取消"):
        g = groups.get(kind) or []
        if not g:
            continue
        c = Counter(g)
        detail = "、".join(f"{k}{v}個" for k, v in c.items())
        lines.append(f"{len(g)}{lab[kind]}（涉及{detail}）")
    return lines


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
