#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
feed_schema.py — feed 維度規範化：【一處派生，全部 module 共用】。

點解要有呢個檔（user 2026-08-17 睇完 databook 之後定）：

1. `year_bucket` 一條欄塞咗【兩個維度】—— "25_24SY" = 2024年【計劃】、2025年【發生】。
   報告每張表都係喺呢兩個維度上切（1.2 = 計劃年25；2.1 = 計劃年24 且 發生年25；
   4.1 = 發生年25 按計劃年分欄；2.5 = 計劃年 × 已認可/期後），
   所以之前每個 module 都自己 parse 一次字串 → 一改就走樣。
   → `split_year()` 一處拆，出 `plan_year` / `spend_year`。

2. 「範疇」（報告 row label）＝ 博彩睇 vertical_label、非博彩睇 ng_label，
   呢句之前喺 5 個 module 各寫一次 → `sub_of()` / `add_dims()` 統一。

3. 金額欄有三代並存（調整金額 / adjustment_amount / 調整_萬），
   → `MEASURES` 定死 canonical，report layer 只認呢幾條。

全部 function 都 idempotent：feed 已經有嗰條欄就唔會覆蓋（咁樣舊 CSV 同新 CSV 都行得）。
"""
import re

_YB = re.compile(r"^(\d{2})(?:_(\d{2})SY)?$")

# report layer 只應該用呢幾條金額欄（其餘 調整金額/adjustment_amount/… 係上游殘留，唔好掂）
MEASURES = {
    "報告投資金額": "調整前_萬",
    "潛在調整金額": "調整_萬",
    "潛在調整後投資金額": "調整後_萬",
}


def split_year(yb):
    """year_bucket → (plan_year, spend_year)，兩位數 int；認唔到回 (None, None)。

        "25"      → (25, 25)   2025年計劃、2025年發生
        "25_24SY" → (24, 25)   2024年計劃、2025年發生（＝期後）
        "24_23SY" → (23, 24)
        "23"      → (23, 23)
    ⚠ 前面嗰個數字係【發生年】，_NNSY 嗰個先係【計劃年】。
    """
    m = _YB.match(str(yb).strip())
    if not m:
        return None, None
    spend = int(m.group(1))
    return (int(m.group(2)) if m.group(2) else spend), spend


def sub_of(df):
    """報告 row label「範疇」：博彩用 vertical_label、非博彩用 ng_label。
    feed 已經有物化嘅「範疇」欄就直接用（prep_tableau 出）。"""
    if "範疇" in df.columns:
        s = df["範疇"].astype(str).str.strip()
        if not s.isin(["", "nan", "None"]).all():
            return s
    v = df["vertical_label"] if "vertical_label" in df.columns else ""
    n = df["ng_label"] if "ng_label" in df.columns else ""
    gm = df["ng_scope"].astype(str).str.strip().eq("gaming")
    import pandas as pd
    return pd.Series(v, index=df.index).where(gm, pd.Series(n, index=df.index))


def add_dims(df):
    """就地加 plan_year / spend_year / 範疇（已經有就唔郁）→ 回 df。"""
    if "year_bucket" in df.columns and ("plan_year" not in df.columns
                                        or "spend_year" not in df.columns):
        yb = df["year_bucket"].astype(str).str.strip()
        pairs = {v: split_year(v) for v in yb.unique()}
        if "plan_year" not in df.columns:
            df["plan_year"] = yb.map(lambda v: pairs[v][0])
        if "spend_year" not in df.columns:
            df["spend_year"] = yb.map(lambda v: pairs[v][1])
    if "範疇" not in df.columns:
        df["範疇"] = sub_of(df)
    return df
