#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dump_tableau_feed.py — 睇 prep_tableau 出嘅餵 Tableau 檔（= 報告數字源）有咩，
順手 pivot 出「項目 × 調整一級 × 金額」預覽，做「單個項目審查結果匯總表」native 生成嘅底。

用法（Windows，kpi-main 底下）：
    pip install pandas openpyxl pyarrow     # 如未裝
    # 指定 MGM 餵 Tableau 檔（xlsx / csv / parquet 都得）：
    python scripts\\report\\dump_tableau_feed.py "data\\tableau\\tableau_25_mgm.xlsx"
    # 或指定 combined 檔 + entity 篩選：
    python scripts\\report\\dump_tableau_feed.py "data\\tableau\\tableau_combined_25.csv" --entity mgm
    # 或俾 data\\tableau 資料夾自動揀第一個 mgm 檔：
    python scripts\\report\\dump_tableau_feed.py "data\\tableau" --entity mgm

輸出（貼返嚟俾我）：
  1. 檔、shape、全部欄名 + dtype
  2. 關鍵維度欄 distinct 值（報告年 / year_bucket / entity / ng_label / vertical_label /
     horizontal_label / final_capex_opex / row_type / 調整一級 / 調整二級）
  3. MGM（報告年=25）逐項目：申報(調整前_萬) / Σ調整_萬 / 調整後_萬
  4. pivot 預覽：項目 × 調整一級 → Σ調整_萬（頭 25 項目 + 合計行）＝「審查結果匯總表」雛形
"""
import sys
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("✗ 未裝 pandas → pip install pandas openpyxl pyarrow")
    sys.exit(1)

pd.set_option("display.max_columns", 60)
pd.set_option("display.width", 200)
pd.set_option("display.max_rows", 120)

# 可能存在嘅欄名（prep_tableau schema），實際以檔為準
DIM_COLS = ["報告年", "year_bucket", "entity", "ng_scope", "ng_code", "ng_label",
            "vertical_label", "horizontal_label", "final_capex_opex", "row_type",
            "調整一級", "調整二級", "remark"]
PROJ_COLS = ["dicj code", "dicj_code", "項目編號", "project", "項目名稱",
             "subproject code", "project_code", "subproject"]
AMT_COLS = ["amount_mop", "調整前_萬", "調整_萬", "調整後_萬", "對數金額_萬"]


def _load(path: Path) -> pd.DataFrame:
    s = path.suffix.lower()
    if s == ".parquet":
        return pd.read_parquet(path)
    if s in (".csv", ".txt"):
        return pd.read_csv(path)
    if s in (".xlsx", ".xlsm"):
        return pd.read_excel(path)
    raise ValueError(f"唔識個格式: {s}")


def _pick(df, names):
    return [c for c in names if c in df.columns]


def main():
    args = sys.argv[1:]
    entity = None
    if "--entity" in args:
        i = args.index("--entity"); entity = args[i + 1].lower(); del args[i:i + 2]
    if not args:
        print("俾 tableau feed 檔或 data\\tableau 資料夾")
        return
    p = Path(args[0])
    if p.is_dir():
        cands = sorted([x for x in p.glob("*")
                        if x.suffix.lower() in (".parquet", ".csv", ".xlsx")
                        and (not entity or entity in x.name.lower())])
        if not cands:
            print(f"✗ {p} 揾唔到 feed 檔"); return
        p = cands[0]
        print(f"（自動揀 {p.name}）")

    df = _load(p)
    print(f"\n{'='*74}\n# {p.name}   shape={df.shape}")
    print("\n── 全部欄 + dtype ──")
    for c in df.columns:
        print(f"    {c!r:42} {df[c].dtype}")

    if entity and "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity]
        print(f"\n（篩 entity={entity} → {len(df)} 行）")

    print("\n── 關鍵維度 distinct（cap 40）──")
    for c in _pick(df, DIM_COLS):
        vals = df[c].dropna().unique().tolist()
        show = vals[:40]
        print(f"    {c}（{len(vals)}）: {show}")

    # 只睇報告年=25（2025 版）——同報告「2025年度…單個項目審查結果匯總表」對應
    d = df
    if "報告年" in df.columns:
        yr = df["報告年"].astype(str)
        d = df[yr.isin(["25", "2025"])]
        print(f"\n（報告年=25 → {len(d)} 行）")

    proj = _pick(d, PROJ_COLS)
    amt = _pick(d, AMT_COLS)
    adj1 = "調整一級" if "調整一級" in d.columns else None
    print(f"\n── 用嚟 group 嘅項目欄: {proj}")
    print(f"── 金額欄: {amt}")
    print(f"── 調整類型欄: {adj1}")

    # 逐項目滙總
    key = _pick(d, ["dicj code", "dicj_code", "項目編號", "subproject code",
                    "project_code", "project", "項目名稱", "subproject"])[:2]
    if key and amt:
        g = d.groupby(key, dropna=False)[amt].sum().round(1)
        print(f"\n── 逐項目滙總（key={key}，頭 30）──")
        print(g.head(30).to_string())
        print(f"    …共 {len(g)} 個項目；合計:\n{g.sum().to_frame('Σ').T.to_string()}")

    # pivot：項目 × 調整一級 → Σ調整_萬（＝匯總表雛形）
    if adj1 and key and "調整_萬" in d.columns:
        try:
            pv = pd.pivot_table(d, index=key, columns=adj1, values="調整_萬",
                                aggfunc="sum", fill_value=0).round(1)
            print(f"\n── pivot 預覽 項目×調整一級 → Σ調整_萬（頭 25）──")
            print(pv.head(25).to_string())
            print(f"\n    合計行:\n{pv.sum().to_frame('Σ').T.to_string()}")
        except Exception as e:
            print(f"    pivot 失敗: {e}")


if __name__ == "__main__":
    main()
