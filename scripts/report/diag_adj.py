#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
diag_adj.py — 由 feed 直接打印【計劃年 bucket × 調整類型】嘅金額同行數（原始 `調整一級`
兼 canonical 8 類都出），對照原報告 1.4／2.2／2.4 三張調整表。

點解要：報告嗰三張表邊一欄有數，完全由 feed 嘅 `調整一級` 決定，報告層睇唔到來源。
diff_report 試過對出我哋成類調整喺某個 bucket 完全缺席（原報告有、我哋冇）。要分清係
feed 根本冇行派到呢類，定係派咗去第二類，就要睇 feed 原始分佈。

用法：
    python scripts\\report\\diag_adj.py                        # mgm + root tableau_combined_25.csv
    python scripts\\report\\diag_adj.py mgm --feed xxx.csv
    python scripts\\report\\diag_adj.py mgm --detail 日常營運   # 該類逐 bucket／項目列出
    python scripts\\report\\diag_adj.py mgm --raw              # 連未 canonical 化嘅原字串一齊出

睇法：每個 bucket 之下嘅 ⚠「呢個 bucket 冇、其他 bucket 有」就係要跟嘅線索 ——
該類 rule 喺呢個計劃年冇 fire，同原報告逐欄對就知係咪應該有。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from build_project_review_table import CANON, ADJ_ALL
except Exception:                                   # 單檔攞出去跑都唔想炸
    CANON, ADJ_ALL = {}, []

try:
    import pandas as pd
except ImportError:
    print("✗ 未裝 pandas → pip install pandas"); sys.exit(1)

# 報告年 25 之下，year_bucket → 計劃年 bucket（同 make_report 一致）
BUCKET = {"25": "2025年度投資計劃（1.4 表）",
          "25_24SY": "2024年度計劃期後投資（2.2 表）",
          "25_23SY": "2023年度計劃期後投資（2.4 表）"}
FEEDS = ["tableau_combined_25.csv", "data/tableau/tableau_combined_25.csv"]
AMT = ["調整_萬", "潛在調整金額", "調整金額", "adjustment_amount"]


def _pick(df, names):
    return next((c for c in names if c in df.columns), None)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    av = sys.argv[1:]
    ent = next((a for a in av if not a.startswith("--")), "mgm")
    feed = av[av.index("--feed") + 1] if "--feed" in av else \
        next((f for f in FEEDS if Path(f).exists()), None)
    detail = av[av.index("--detail") + 1] if "--detail" in av else None
    raw = "--raw" in av
    if not feed:
        print(f"✗ 搵唔到 feed（試過 {', '.join(FEEDS)}）；用 --feed 指明"); return

    df = pd.read_csv(feed, low_memory=False)
    amt = _pick(df, AMT)
    if "調整一級" not in df.columns or amt is None or "year_bucket" not in df.columns:
        print(f"✗ feed 欠欄（要 year_bucket／調整一級／{'／'.join(AMT)}）")
        print(f"   feed 現有：{list(df.columns)}"); return

    d = df[df["entity"].astype(str).str.lower() == ent.lower()].copy() \
        if "entity" in df.columns else df.copy()
    if d.empty:
        d = df[df["entity"].astype(str).str.contains(ent, case=False, na=False)].copy()
    d[amt] = pd.to_numeric(d[amt], errors="coerce").fillna(0.0)
    d["_yb"] = d["year_bucket"].astype(str).str.strip()
    d = d[d["_yb"].isin(BUCKET)]
    d["_raw"] = d["調整一級"].fillna("").astype(str).str.strip()
    d["_adj"] = d["_raw"].map(CANON).fillna(d["_raw"]).replace("", "（無調整）")
    print(f"### {feed}｜entity={ent}｜金額欄 {amt}｜{len(d):,} 行（報告年 25 三個 bucket）\n")

    key = "_raw" if raw else "_adj"
    live, tables = {}, {}                       # 邊幾類全份有數（用嚟捉「淨係呢個 bucket 冇」）
    for yb in BUCKET:
        sub = d[d["_yb"] == yb]
        g = (sub.groupby(sub[key].replace("", "（無調整）"))
                .agg(金額=(amt, "sum"), 行數=(amt, "size")).sort_values("金額"))
        tables[yb] = (sub, g[g["金額"].abs() > 0.4])
        live.update({k: 1 for k in tables[yb][1].index})

    for yb, label in BUCKET.items():
        sub, g = tables[yb]
        if sub.empty:
            print(f"── {label}：冇行\n"); continue
        print(f"── {label}　調整合計 {sub[amt].sum():,.0f} 萬")
        for k, r in g.iterrows():
            no = f"[{ADJ_ALL.index(k) + 1}] " if k in ADJ_ALL else "    "
            print(f"     {r['金額']:>11,.0f} 萬　{int(r['行數']):>6,} 行　{no}{k}")
        # 只報「呢個 bucket 冇、但第二個 bucket 有」嘅類 —— 淨係嗰啲先算異常
        gone = [t for t in live if t not in set(g.index) and t != "（無調整）"]
        if gone:
            print(f"     ⚠ 呢個 bucket 冇、其他 bucket 有：{'、'.join(gone)}")
        print()

    if detail:
        print(f"══ 逐 bucket／項目明細：調整類型 含「{detail}」")
        m = d[d["_adj"].str.contains(detail, na=False) | d["_raw"].str.contains(detail, na=False)]
        if m.empty:
            print("   （三個 bucket 都冇呢一類 → rule 根本冇派到，唔係派錯去第二類）"); return
        pcol = _pick(m, ["project", "項目名稱", "dicj_code", "DICJ Code"])
        g = (m.groupby(["_yb"] + ([pcol] if pcol else []))
              .agg(金額=(amt, "sum"), 行數=(amt, "size")).sort_values("金額"))
        for k, r in g.iterrows():
            print(f"   {r['金額']:>10,.0f} 萬　{int(r['行數']):>5,} 行　{k}")


if __name__ == "__main__":
    main()
