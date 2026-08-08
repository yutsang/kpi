#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_project.py — dump 逐項目（project-level）數：報告年 / 範疇 / 項目 / 計劃 / 調整前 / 調整 / 調整後 / 完成率 /
設施 / 活動。寫去 results\\project_dump.tsv（gitignored）+ 印出。俾 demo_page.py 讀去 render 真數 pptx。

用法（Windows）：
    python scripts\\report\\inspect_project.py "tableau_combined_25.csv" --entity mgm ^
        --qingdan "data\\投資項目清單\\3. MGM.…投资项目清单.xlsx"
⚠ 輸出係機密（真項目金額）→ 已寫去 gitignored results\\，貼返俾 Claude 時當機密（Claude 會 untrack）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)
import build_project_review_table as B


def main():
    args = sys.argv[1:]
    entity = qingdan = None
    for flag in ("--entity", "--qingdan"):
        if flag in args:
            i = args.index(flag); v = args[i + 1]; del args[i:i + 2]
            if flag == "--entity":
                entity = v.lower()
            else:
                qingdan = v
    if not args:
        print("俾 feed csv 路徑（--entity mgm --qingdan 清單）"); return
    entity = entity or "mgm"

    df = pd.read_csv(args[0], low_memory=False)
    if "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity]
    df["報告年"] = pd.to_numeric(df["報告年"], errors="coerce")
    df["_sub"] = df.apply(lambda r: r["vertical_label"] if str(r["ng_scope"]) == "gaming" else r["ng_label"], axis=1)
    plan = B.load_plan(Path(qingdan)) if qingdan else {}

    def _num(c):
        return pd.to_numeric(df[c], errors="coerce").fillna(0) if c in df.columns else 0

    for c in ("調整前_萬", "調整_萬", "調整後_萬"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    cap = df[df["final_capex_opex"] == "Capex"].groupby(["報告年", "ng_scope", "_sub", "dicj code"])["調整後_萬"].sum()
    ope = df[df["final_capex_opex"] == "Opex"].groupby(["報告年", "ng_scope", "_sub", "dicj code"])["調整後_萬"].sum()

    g = df.groupby(["報告年", "ng_scope", "_sub", "dicj code"]).agg(
        項目名稱=("project", "first"), 調整前=("調整前_萬", "sum"),
        調整=("調整_萬", "sum"), 調整後=("調整後_萬", "sum")).reset_index()

    rows = []
    for _, r in g.iterrows():
        yr = int(r["報告年"]) if pd.notna(r["報告年"]) else 0
        gm = (str(r["ng_scope"]) == "gaming")
        code = B._norm(r["dicj code"])
        pl = (plan.get(yr, {}) or {}).get((gm, code), 0.0)
        rep = round(r["調整前"], 1)
        rate = f"{rep / pl * 100:.1f}%" if pl else ""
        key = (r["報告年"], r["ng_scope"], r["_sub"], r["dicj code"])
        rows.append([yr, "gaming" if gm else "non_gaming", r["_sub"], str(r["dicj code"]),
                     str(r["項目名稱"]), round(pl, 1), rep, round(r["調整"], 1), round(r["調整後"], 1),
                     rate, round(cap.get(key, 0.0), 1), round(ope.get(key, 0.0), 1)])

    hdr = ["報告年", "scope", "範疇", "項目序號", "項目名稱", "計劃", "調整前", "調整", "調整後", "完成率", "設施", "活動"]
    out = Path("results/project_dump.tsv")
    out.parent.mkdir(parents=True, exist_ok=True)
    lines = ["\t".join(hdr)] + ["\t".join(str(x) for x in r) for r in rows]
    out.write_text("\n".join(lines), encoding="utf-8")
    print(f"✓ {out}（{len(rows)} 個 項目×報告年）— gitignored，貼返俾 Claude render")
    print("\n（頭 20 行預覽）")
    for ln in lines[:21]:
        print(ln)


if __name__ == "__main__":
    main()
