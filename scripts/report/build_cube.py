#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_cube.py — 由 JE 明細 + 清單，出【報告立方】同【計劃盤】兩張細檔。

點解（user 2026-08-17 睇完 databook 之後定嘅方向）：
  ① tableau_combined_25.csv = JE 明細（~620k 行）→ 留住做 drill-down / audit trail
  ② plan_by_project.csv     = 計劃盤（每 entity 每年幾百行）← 完成率個【分母】而家喺清單，
                              每次現場 parse，所以「項目數量 89 vs 報告 95」「完成率唔啱」不停出現
  ③ report_cube.csv         = 報告立方（幾千行）← 報告每一格 = 立方嘅一個 sum 或兩個 sum 相除

之後 pptx 只需要讀 ②③，唔使再 parse year_bucket 字串／join 清單／反推已認可。

用法（Windows）：
    python scripts\\report\\build_cube.py "tableau_combined_25.csv" --entity mgm ^
        --qingdan "data\\投資項目清單\\3. MGM.…投资项目清单.xlsx" [--out .]
    # 唔指定 --entity 就做晒全部 entity
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)

import feed_schema as FS
import build_project_review_table as B

# 立方維度（報告任何一張表都係喺呢啲維度上切）
DIMS = ["entity", "plan_year", "spend_year", "ng_scope", "範疇", "ng_code", "ng_label",
        "final_capex_opex", "調整一級"]
# 立方度量（canonical，唔用上游三代重複欄）
MEAS = {"報告投資金額_萬": "調整前_萬", "潛在調整金額_萬": "調整_萬", "潛在調整後投資金額_萬": "調整後_萬"}


def load_feed(path, entity=None):
    df = pd.read_csv(path, low_memory=False)
    if entity and "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity.lower()].copy()
    df = df[df["dicj code"].astype(str).str.match(r"^項目\s*\d")].copy()   # 丟 pseudo 碼
    FS.add_dims(df)
    for c in DIMS:
        if c not in df.columns:
            df[c] = ""
        df[c] = df[c].fillna("")
    return df


def build_report_cube(df):
    """→ 立方：DIMS × MEAS + 項目數（distinct dicj code）。"""
    d = df.copy()
    for tgt, src in MEAS.items():
        d[tgt] = pd.to_numeric(d[src], errors="coerce").fillna(0.0) if src in d.columns else 0.0
    g = d.groupby(DIMS, dropna=False).agg(
        **{k: (k, "sum") for k in MEAS},
        項目數=("dicj code", "nunique")).reset_index()
    for k in MEAS:
        g[k] = g[k].round(1)
    return g.sort_values(DIMS)


def build_plan_table(df, qingdan, entity):
    """→ 計劃盤：(entity, plan_year, ng_scope, dicj code) × 計劃投資金額 + 範疇 + 有冇實際支出。
    ⚠ 「是否申報為零」而家係由 feed 有冇支出【推】出嚟；正解係項目組喺清單畀一條明碼欄，
       到時直接讀，唔使推（95 vs 89 嗰條數就會自動啱）。"""
    plan = B.load_plan(Path(qingdan)) if qingdan else {}
    cat = {}
    try:
        cat = B.load_category(Path(qingdan)) if qingdan else {}
    except Exception:
        cat = {}
    if not plan:
        return pd.DataFrame()
    # feed 有支出嘅碼（逐計劃年）
    spent = {}
    for py, sub in df.groupby("plan_year"):
        nz = sub[pd.to_numeric(sub["調整前_萬"], errors="coerce").fillna(0) != 0]
        spent[py] = {(str(r["ng_scope"]) == "gaming", B._norm(r["dicj code"]))
                     for _, r in nz.drop_duplicates(["ng_scope", "dicj code"]).iterrows()}
    # feed 學到嘅 碼→範疇
    sub_of = {(str(r["ng_scope"]) == "gaming", B._norm(r["dicj code"])): r["範疇"]
              for _, r in df.drop_duplicates(["ng_scope", "dicj code"]).iterrows()}
    rows = []
    for yr, d in (plan or {}).items():
        for (gm, code), amt in (d or {}).items():
            rows.append({
                "entity": entity, "plan_year": int(yr), "ng_scope": "gaming" if gm else "non_gaming",
                "dicj code": code, "範疇": sub_of.get((gm, code), ""),
                "項目性質": cat.get((gm, code), ""),
                "計劃投資金額_萬": round(float(amt or 0), 1),
                "有實際支出": (gm, code) in spent.get(yr, set()),
            })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out["獲批開展"] = out["計劃投資金額_萬"] > 0      # 計劃金額 0 唔算獲批開展（清單有大量 0 行）
    return out.sort_values(["plan_year", "ng_scope", "dicj code"])


def main():
    av = sys.argv[1:]
    if not av:
        print(__doc__); return
    feed = av[0]

    def opt(f, d=None):
        return av[av.index(f) + 1] if f in av else d
    entity = (opt("--entity") or "").lower() or None
    qingdan = opt("--qingdan")
    out = Path(opt("--out", "."))
    if not Path(feed).exists():
        print(f"✗ 揾唔到 feed {feed}"); return
    df = load_feed(feed, entity)
    print(f"feed {feed}：{len(df):,} 行"
          f"｜plan_year {sorted(set(df['plan_year'].dropna()))}"
          f"｜spend_year {sorted(set(df['spend_year'].dropna()))}")
    tag = entity or "all"

    cube = build_report_cube(df)
    p1 = out / f"{tag}_report_cube.csv"
    cube.to_csv(p1, index=False, encoding="utf-8-sig")
    print(f"✓ {p1.resolve()}  {len(cube):,} 行 × {len(cube.columns)} 欄")
    for py, sub in cube.groupby("plan_year"):
        print(f"    計劃年 {py}：報告 {sub['報告投資金額_萬'].sum():,.0f} 萬"
              f"／調整後 {sub['潛在調整後投資金額_萬'].sum():,.0f} 萬")

    if qingdan:
        plan = build_plan_table(df, qingdan, tag)
        if plan.empty:
            print("    ⚠ 清單讀唔到計劃金額 → 冇出計劃盤")
        else:
            p2 = out / f"{tag}_plan_by_project.csv"
            plan.to_csv(p2, index=False, encoding="utf-8-sig")
            print(f"✓ {p2.resolve()}  {len(plan):,} 行")
            for py, sub in plan.groupby("plan_year"):
                ok = sub[sub["獲批開展"]]
                print(f"    計劃年 {py}：獲批開展 {len(ok)} 個"
                      f"（有支出 {int(ok['有實際支出'].sum())}／零申報 {int((~ok['有實際支出']).sum())}）"
                      f"，計劃金額 {ok['計劃投資金額_萬'].sum():,.0f} 萬")


if __name__ == "__main__":
    main()
