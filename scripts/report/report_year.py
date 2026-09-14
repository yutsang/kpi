#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
report_year.py — 報告年度【一處定義，全 pipeline 派生】。

點解要（2026-09-14，user 拍板做參數化）：
年份本來散落成條 pipeline —— 六大章名、bucket label、封面、單項審查表標題、
LLM prompt…… `canned_yearly.py` 掃出嚟 code 層有成百行寫死「2025年度」。
明年冇新 golden 可抽（user：「到時不會有 report 抽取的」），罐頭要沿用今年嗰份，
如果年份仲係逐處寫死，更新就變成逐個檔揾，好易漏。

改法：報告年度得【一個】來源 —— 環境變數 KPI_REPORT_YEAR，預設 2025。
其餘全部由佢派生：

    YEAR      2025        報告年度（四位）
    Y2        25          兩位（feed 嘅 year_bucket 用兩位）
    P1 / P2   2024 / 2023 前兩個計劃年（期後跟進嗰兩個 bucket）
    BUCKETS   {"25": "2025年度投資計劃", "25_24SY": …, "25_23SY": …}
    SECTIONS  六大章名

⚠ BUCKETS 啲 label 同時係【內部 data key】（`_bucket` 欄嘅值由佢派生），
  所以一定要全 pipeline 用同一個來源，唔可以有啲地方自己寫死。

用法：
    import report_year as RY
    RY.YEAR, RY.BUCKETS, RY.SECTIONS
    set KPI_REPORT_YEAR=2026   （Windows；改一個數全份跟住變）
"""
import os

YEAR = int(os.environ.get("KPI_REPORT_YEAR", "2025"))
Y2 = YEAR % 100
P1, P2 = YEAR - 1, YEAR - 2          # 期後跟進嗰兩個計劃年

# feed 嘅 year_bucket 碼：本年 "25"；期後 "25_24SY"（2024年計劃、2025年發生）
YB = f"{Y2:02d}"
YB_P1 = f"{Y2:02d}_{P1 % 100:02d}SY"
YB_P2 = f"{Y2:02d}_{P2 % 100:02d}SY"

LBL = f"{YEAR}年度投資計劃"
LBL_P1 = f"{P1}年度計劃期後投資"
LBL_P2 = f"{P2}年度計劃期後投資"

BUCKETS = {YB: LBL, YB_P1: LBL_P1, YB_P2: LBL_P2}
BUCKET_ORDER = [LBL, LBL_P1, LBL_P2]
BUCKET_PLANYR = {LBL: Y2, LBL_P1: P1 % 100, LBL_P2: P2 % 100}

# 1.4／2.2／2.4 調整匯總表嘅欄名（同 bucket label 唔同寫法）
BK_COL = [(LBL, f"{YEAR}年度投資計劃"),
          (LBL_P1, f"{P1}年度投資計劃期後事項"),
          (LBL_P2, f"{P2}年度投資計劃期後事項")]

SECTIONS = [
    f"{YEAR}年度投資計劃執行情況概述",
    f"過往年度投資計劃在{YEAR}年繼續執行的審查跟進",
    "本年度審查工作的主要發現",
    "其他信息",
    "投資計劃執行報告的六項KPI分析",
    "附件",
]

COVER_TITLE = f"{YEAR}年年度投資計劃執行情況審查"

# 單項審查匯總表標題（render_review_table_pptx）
REVIEW_TITLE = {
    f"報告年{Y2:02d}": f"{{e}} {YEAR}年度投資計劃單個項目審查結果匯總表",
    f"報告年{P1 % 100:02d}": f"{{e}} {P1}年度投資計劃單個項目截至{YEAR}年末的審查結果匯總表",
    f"報告年{P2 % 100:02d}": f"{{e}} {P2}年度投資計劃單個項目截至{YEAR}年末的審查結果匯總表",
}


def banner():
    """build 開頭印一行，一眼睇到跑緊邊個年度（唔係預設就特別標出）。"""
    tag = "" if YEAR == 2025 else "　★ 非預設年度（KPI_REPORT_YEAR）"
    return f"報告年度 {YEAR}（bucket {YB}／{YB_P1}／{YB_P2}）{tag}"
