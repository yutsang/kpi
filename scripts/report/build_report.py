#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from __future__ import annotations
"""
build_report.py — 單一自足檔：由底層數據（feed + 清單）生成表二審查報告 pptx。
毋須任何 prerequisite / 其他模組 / 前置 command：
    python build_report.py [entity]            # 預設 mgm；用現有 {entity}_llm_narrative.json 或清單 fallback
    python build_report.py [entity] --llm      # 即場生成 LLM 敘述（需 KPMG 網 + workbench creds）再出報告
（此檔由各 build/LLM 模組自動合併；LLM 相關 heavy import [openai/msoffcrypto] 全 lazy；報告只作 ref。）
"""
import sys
from pathlib import Path
import re
import json
import os
from typing import Any
import io
from concurrent.futures import ThreadPoolExecutor, as_completed


# ── from render_review_table_pptx ──
try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
except ImportError:
    print("✗ 需要 pandas + python-pptx → pip install pandas python-pptx openpyxl"); sys.exit(1)


# ── from render_review_table_pptx ──
RATE_COLS = {"投資計劃完成率", "潛在調整後投資計劃完成率"}


# ── from render_review_table_pptx ──
TEXT_COLS = {"項目序號", "項目名稱"}


# ── from render_review_table_pptx ──
G1 = ["項目序號", "項目名稱", "計劃投資金額", "報告投資金額", "投資計劃完成率"]


# ── from render_review_table_pptx ──
G3 = ["調整後投資金額", "潛在調整後投資計劃完成率", "設施建設/資本性支出", "活動舉辦/營運性支出"]


# ── from render_review_table_pptx ──
GROUP_LABEL = {"G1": "項目基本信息", "G2": "投資金額的潛在調整事項", "G3": "潛在調整後投資金額"}


# ── from render_review_table_pptx ──
BLUE = RGBColor(0x00, 0x33, 0x8D)


# ── from render_review_table_pptx ──
LBLUE = RGBColor(0xD9, 0xE1, 0xF2)


# ── from render_review_table_pptx ──
GREY = RGBColor(0xE7, 0xE6, 0xE6)


# ── from render_review_table_pptx ──
WHITE = RGBColor(0xFF, 0xFF, 0xFF)


# ── from render_review_table_pptx ──
RED = RGBColor(0xC0, 0x00, 0x00)


# ── from render_review_table_pptx ──
GROUP_FILL = {"G1": RGBColor(0x2E, 0x9B, 0xD6), "G2": RGBColor(0x9D, 0xC3, 0xE6),
              "G3": RGBColor(0x1F, 0x38, 0x64)}


# ── from render_review_table_pptx ──
GROUP_TEXT = {"G1": WHITE, "G2": RGBColor(0x1F, 0x38, 0x64), "G3": WHITE}


# ── from render_review_table_pptx ──
ROWS_PER_SLIDE = 24


# ── from render_review_table_pptx ──
YEAR_TITLE = {
    "報告年25": "MGM 2025年度投資計劃單個項目審查結果匯總表",
    "報告年24": "MGM 2024年度投資計劃單個項目截至2025年末的審查結果匯總表",
    "報告年23": "MGM 2023年度投資計劃單個項目截至2025年末的審查結果匯總表",
}


# ── from render_review_table_pptx ──
def fmt_money(v):
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    n = int(round(f))
    if n == 0:
        return "-"
    return f"({abs(n):,})" if n < 0 else f"{n:,}"


# ── from render_review_table_pptx ──
def fmt_pct(v):
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return f"{float(v)*100:.1f}%"
    except (TypeError, ValueError):
        return str(v)


# ── from render_review_table_pptx ──
def row_kind(seq: str) -> str:
    s = str(seq)
    if s.endswith("小計") or s.endswith("合計"):
        return "subtotal"
    if "—" in s or "－" in s:
        return "section"
    return "data"


# ── from render_review_table_pptx ──
def _set(cell, text, *, size=7, bold=False, align=PP_ALIGN.RIGHT, color=None, fill=None):
    cell.margin_left = cell.margin_right = Emu(18000)
    cell.margin_top = cell.margin_bottom = Emu(9000)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid(); cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run(); r.text = "" if text is None else str(text)
    f = r.font
    f.size = Pt(size); f.bold = bold
    f.name = "Microsoft JhengHei"
    if color is not None:
        f.color.rgb = color


# ── from render_review_table_pptx ──
def col_group(c):
    if c in G1:
        return "G1"
    if c in G3:
        return "G3"
    return "G2"


# ── from render_review_table_pptx ──
def render_sheet(prs, sheet_name, df, cols):
    n = len(df)
    pages = [(i, min(i + ROWS_PER_SLIDE, n)) for i in range(0, n, ROWS_PER_SLIDE)] or [(0, 0)]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    # 欄寬：項目序號窄、名闊、數字中
    def cw(c):
        if c == "項目序號":
            return 0.62
        if c == "項目名稱":
            return 1.7
        return 0.66
    widths = [cw(c) for c in cols]
    total_w = sum(widths)
    left = 0.4
    slide_w = prs.slide_width / 914400.0
    scale = min(1.0, (slide_w - 0.8) / total_w)
    widths = [w * scale for w in widths]

    for pi, (a, b) in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        # 標題
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(slide_w - 0.8), Inches(0.4))
        _set_title(tb, f"{YEAR_TITLE.get(sheet_name, sheet_name)}（{pi+1}/{len(pages)}）")
        sub = df.iloc[a:b]
        nrow = 2 + len(sub)     # 2 header rows
        ncol = len(cols)
        gtab = slide.shapes.add_table(nrow, ncol, Inches(left), Inches(0.75),
                                      Inches(sum(widths)), Inches(0.3 * nrow)).table
        for ci, w in enumerate(widths):
            gtab.columns[ci].width = Inches(w)
        # header row0：group 合併（3 色）
        ci = 0
        while ci < ncol:
            g = col_group(cols[ci])
            cj = ci
            while cj + 1 < ncol and col_group(cols[cj + 1]) == g:
                cj += 1
            if cj > ci:
                gtab.cell(0, ci).merge(gtab.cell(0, cj))
            _set(gtab.cell(0, ci), GROUP_LABEL[g], size=8, bold=True,
                 align=PP_ALIGN.CENTER, color=GROUP_TEXT[g], fill=GROUP_FILL[g])
            ci = cj + 1
        # header row1：欄名（跟欄組色）
        for ci, c in enumerate(cols):
            g = col_group(c)
            _set(gtab.cell(1, ci), c, size=6.5, bold=True, align=PP_ALIGN.CENTER,
                 color=GROUP_TEXT[g], fill=GROUP_FILL[g])
        # data rows
        for ri, (_, row) in enumerate(sub.iterrows(), start=2):
            kind = row_kind(row["項目序號"])
            for ci, c in enumerate(cols):
                v = row.get(c, "")
                if c in TEXT_COLS:
                    txt, al = ("" if v is None else str(v)), (PP_ALIGN.LEFT)
                elif c in RATE_COLS:
                    txt, al = fmt_pct(v), PP_ALIGN.RIGHT
                else:
                    txt, al = fmt_money(v), PP_ALIGN.RIGHT
                fill = LBLUE if kind == "section" else (GREY if kind == "subtotal" else None)
                bold = kind in ("section", "subtotal")
                color = RED if (txt.startswith("(") ) else None
                if kind == "section" and ci == 0:
                    # section 標題橫跨全行
                    gtab.cell(ri, 0).merge(gtab.cell(ri, ncol - 1))
                    _set(gtab.cell(ri, 0), txt, size=7, bold=True, align=PP_ALIGN.LEFT, fill=LBLUE)
                    break
                _set(gtab.cell(ri, ci), txt, size=6.5, bold=bold,
                     align=al, color=color, fill=fill)
        print(f"    {sheet_name} 第 {pi+1}/{len(pages)} 頁：{b-a} 行 × {ncol} 欄")


# ── from render_review_table_pptx ──
def _set_title(tb, text):
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = BLUE; r.font.name = "Microsoft JhengHei"


# ── from build_narrative ──
try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    print("✗ pip install openpyxl"); sys.exit(1)


# ── from build_narrative ──
CONCEPTS = [
    ("項目狀況", ["項目狀況"]),
    ("實施地點", ["實際實施地點", "實際實施地點/空間"]),
    ("計劃投資內容", ["計劃投資內容"]),
    ("實際投資內容", ["實際投資內容"]),
    ("KPMG分析發現", ["分析發現", "投資偏離"]),
    ("管理層解釋", ["管理層解釋", "變更原因"]),
    ("期後調整內容", ["期後調整内容", "期後調整內容"]),
    ("調整事項備註", ["調整事項備註", "調整備註"]),
    # 跨司工作組審計軌跡（報告調整詳述會引）：政府裁決 + KPMG 最終立場
    ("跨司回覆", ["跨司工作組的回", "跨司工作组的回", "跨司工作組回", "跨司工作组回", "第二輪意見", "第二輪意见"]),
    ("KPMG回覆", ["KPMG回覆", "KPMG回复"]),
]


# ── from build_narrative ──
def _norm_code(v):
    s = re.sub(r"\s+", "", str(v if v is not None else ""))
    s = re.sub(r"^項目", "", s)
    m = re.match(r"^0*(\d+)$", s)
    return m.group(1) if m else s


# ── from build_narrative ──
_PLACEHOLDER_RE = re.compile(r"參閱附件|請參閱|見附件|同附件|詳見|組織架構進展|^n/?a$|^-+$")


# ── from build_narrative ──
def _is_placeholder(s):
    s = str(s).strip()
    return len(s) < 8 or bool(_PLACEHOLDER_RE.search(s))


# ── from build_narrative ──
def nlook(narr, ng_scope, code):
    """由 (ng_scope, dicj code) 揾清單 rec —— 處理博彩/非博彩共用項目N 撞號（先試 exact）。"""
    g = (ng_scope == "gaming")
    c = _norm_code(code)
    return narr.get((g, c)) or narr.get((not g, c)) or {}


# ── from build_narrative ──
OUT_COLS = ["項目序號", "項目名稱", "項目類型"] + [n for n, _ in CONCEPTS]


# ── from build_narrative ──
def load_narrative(path, log=lambda *a: None) -> dict:
    """讀清單 → {正規化項目編號: rec}，rec 有 項目序號/名稱/類型 + 每概念（取最新非空）。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = None
    for sn in wb.sheetnames:
        if sn.lower().startswith("database") or sn == "Database":
            ws = wb[sn]; break
    ws = ws or wb[wb.sheetnames[0]]
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        rows.append(r)
        if i > 2500:
            break
    hdr_r = code_c = name_c = type_c = None
    for ri in range(min(14, len(rows))):
        for ci, v in enumerate(rows[ri] or []):
            sv = "" if v is None else str(v)
            if "承批公司項目序號" in sv:
                hdr_r, code_c = ri, ci
            if name_c is None and sv.strip() == "項目名稱":
                name_c = ci
            if type_c is None and "項目類型" in sv:
                type_c = ci
        if hdr_r is not None:
            break
    if hdr_r is None:
        log("✗ 揾唔到『承批公司項目序號』表頭"); return {}
    hdr = [("" if v is None else str(v).replace("\n", "")) for v in rows[hdr_r]]
    concept_cols = {name: [ci for ci, h in enumerate(hdr) if any(k in h for k in keys)]
                    for name, keys in CONCEPTS}
    log(f"清單 header row {hdr_r+1}；概念欄命中：" +
        ", ".join(f"{n}:{len(concept_cols[n])}" for n, _ in CONCEPTS))
    out = {}
    for ri in range(hdr_r + 1, len(rows)):
        row = rows[ri]
        code = _norm_code(row[code_c]) if code_c < len(row) else ""
        if not code:
            continue
        ptype = (row[type_c] if type_c is not None and type_c < len(row) else "") or ""
        gaming = str(ptype).strip().startswith("博彩")     # 博彩/非博彩 共用項目N → key 帶 gaming
        rec = {"項目序號": f"項目{code}",
               "項目名稱": (row[name_c] if name_c is not None and name_c < len(row) else "") or "",
               "項目類型": ptype}
        for name, _ in CONCEPTS:
            vals = []
            for ci in concept_cols[name]:
                if ci < len(row) and row[ci] not in (None, "", 0, "0"):
                    sv = str(row[ci]).strip()
                    if not sv or sv.lower() in ("n/a", "nan"):
                        continue
                    if name in ("實際投資內容", "計劃投資內容") and _is_placeholder(sv):
                        continue                            # 跳 「參閱附件/請參閱/見附件」等 placeholder
                    if sv not in vals:
                        vals.append(sv)
            rec[name] = vals[-1] if vals else ""
        out.setdefault((gaming, code), rec)                 # key = (博彩?, 正規化碼)
    return out


# ── from build_project_review_table ──
try:
    import pandas as pd
except ImportError:
    print("✗ 未裝 pandas → pip install pandas openpyxl"); sys.exit(1)


# ── from build_project_review_table ──
pd.set_option("display.max_columns", 40)


# ── from build_project_review_table ──
pd.set_option("display.width", 260)


# ── from build_project_review_table ──
CANON = {
    "計入報告投資金額的高管及一般支持人員人工成本": "一般支持性部門的人工成本",
    "一般性支持部門的人工成本": "一般支持性部門的人工成本",
    "一般支持性部門的人工成本": "一般支持性部門的人工成本",
    "其他日常營運支出調整": "其他日常營運支出調整",
    "其他日常營運支出調整調整": "其他日常營運支出調整",
    "超過可計入範圍的内部資源支出": "超出可計入範圍的內部資源支出",
    "超出可計入範圍的內部資源支出": "超出可計入範圍的內部資源支出",
    "酒店客房改造支出": "酒店客房改造支出",
    "不符合“吸引外國客源”定義的相關投資支出": "不符合“吸引外國客源”定義的相關投資支出",
    "不符合吸引外國客源": "不符合“吸引外國客源”定義的相關投資支出",
    "未完全實現投資目的的投資支出": "未完全實現投資目的的投資支出",
    "未能實現投資目的的投資支出": "未完全實現投資目的的投資支出",
    "計劃獲批前的投資": "投資計劃獲批前發生且未被認可的投資支出",
    # 以下係跨年機制，報告 7 欄冇（user defer 跨年）→ 落 trailing other，之後 filter 當年就消失
    "將2024年的支出計入2025年報告投資金額": "〔跨年〕將往年支出計入本年報告投資金額",
    "2024年度計劃與2023年度計劃期後調整之間的報告投資金額跨期調整": "〔跨年〕年度計劃期後跨期調整",
}


# ── from build_project_review_table ──
ADJ7 = [
    "一般支持性部門的人工成本", "其他日常營運支出調整", "超出可計入範圍的內部資源支出",
    "酒店客房改造支出", "不符合“吸引外國客源”定義的相關投資支出",
    "未完全實現投資目的的投資支出", "投資計劃獲批前發生且未被認可的投資支出",
]


# ── from build_project_review_table ──
BASE_L = ["計劃投資金額", "報告投資金額", "投資計劃完成率"]


# ── from build_project_review_table ──
TAIL_L = ["潛在調整合計", "調整後投資金額", "潛在調整後投資計劃完成率", "設施建設/資本性支出", "活動舉辦/營運性支出"]


# ── from build_project_review_table ──
_PLAN_RE = {
    25: re.compile(r"2025.*預計投資金額.*合計|2025.*預計投資金額（萬"),
    24: re.compile(r"2024.*預計投資金額.*合計|2024.*預計投資金額（萬"),
    23: re.compile(r"2023.*預計投資金額"),
}


# ── from build_project_review_table ──
def _norm(c) -> str:
    s = re.sub(r"\s+", "", str(c if c is not None else ""))
    s = re.sub(r"^項目", "", s)
    m = re.match(r"^0*(\d+(?:\.\d+)?)$", s)
    return (m.group(1) if m else s).lower()


# ── from build_project_review_table ──
def _plan_year(yb) -> int:
    """year_bucket → 計劃年份（報告匯總表按計劃年分，唔係支出年）：
    25→25(2025計劃), 25_24SY→24(2024計劃嘅2025期後), 25_23SY→23, 24→24, 24_23SY→23, 23→23。"""
    s = str(yb).strip()
    m = re.search(r"_(\d+)SY", s)          # 有 _NNSY = 該計劃年嘅期後
    if m:
        return int(m.group(1))
    m2 = re.match(r"^(\d{2})", s)
    return int(m2.group(1)) if m2 else -1


# ── from build_project_review_table ──
def load_plan(path: Path, log=print) -> dict:
    """{報告年: {(is_gaming, 正規化項目編號): 計劃_萬}}（清單 database；博彩/非博彩共用項目N → key 帶 scope）。"""
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
        hdr_r = code_c = type_c = None
        for ri in range(min(14, len(rows))):
            for ci, v in enumerate(rows[ri] or []):
                sv = "" if v is None else str(v)
                if "承批公司項目序號" in sv:
                    hdr_r, code_c = ri, ci
                if type_c is None and "項目類型" in sv:
                    type_c = ci
            if hdr_r is not None:
                break
        if hdr_r is None:
            continue
        hdr = [("" if v is None else str(v).replace("\n", "")) for v in rows[hdr_r]]
        plan_c = {}
        for yr, rgx in _PLAN_RE.items():
            for ci, h in enumerate(hdr):
                if rgx.search(h):
                    plan_c[yr] = ci; break
        if not plan_c:
            continue
        log("  清單 " + sn + ": 計劃欄 " + ", ".join(f"{yr}→{hdr[ci][:18]!r}" for yr, ci in plan_c.items())
            + (f"；項目類型欄 col{type_c}" if type_c is not None else "；⚠冇項目類型欄(scope 分唔到)"))
        for ri in range(hdr_r + 1, len(rows)):
            row = rows[ri]
            code = _norm(row[code_c]) if code_c < len(row) else ""
            if not code:
                continue
            gaming = False
            if type_c is not None and type_c < len(row) and row[type_c] is not None:
                gaming = str(row[type_c]).strip().startswith("博彩")
            for yr, ci in plan_c.items():
                if ci < len(row):
                    try:
                        out[yr].setdefault((gaming, code), float(row[ci]))
                    except (TypeError, ValueError):
                        pass
        break
    return out


# ── from build_project_review_table ──
def load_category(path: Path, log=lambda *a: None) -> dict:
    """{(is_gaming, 正規化項目編號): 項目性質(D)}（清單；用嚟將零投資項目計劃 attribute 返範疇）。
    含 feed 冇嘅零投資項目（清單有齊全部項目）。"""
    import openpyxl
    out = {}
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
        hdr_r = code_c = type_c = sub_c = None
        for ri in range(min(14, len(rows))):
            for ci, v in enumerate(rows[ri] or []):
                sv = "" if v is None else str(v)
                if "承批公司項目序號" in sv:
                    hdr_r, code_c = ri, ci
                if type_c is None and "項目類型" in sv:
                    type_c = ci
                if sub_c is None and sv.strip() == "項目性質":
                    sub_c = ci
            if hdr_r is not None:
                break
        if hdr_r is None or sub_c is None:
            continue
        for ri in range(hdr_r + 1, len(rows)):
            row = rows[ri]
            code = _norm(row[code_c]) if code_c < len(row) else ""
            if not code:
                continue
            gaming = (type_c is not None and type_c < len(row) and row[type_c] is not None
                      and str(row[type_c]).strip().startswith("博彩"))
            sub = row[sub_c] if sub_c < len(row) else None
            if sub is not None and str(sub).strip():
                out.setdefault((gaming, code), str(sub).strip())
        log(f"  清單 項目性質(D)：{len(out)} 個項目")
        break
    return out


# ── from build_project_review_table ──
def _rate(rep, plan):
    try:
        return round(rep / plan, 4) if plan else None
    except Exception:
        return None


# ── from build_project_review_table ──
def build_year(df: pd.DataFrame, year: int, plan: dict | None = None):
    # 按計劃年份(plan year)分表：跨年期後(25_24SY 等)落返啱嘅計劃年，唔會爆 2025 表小計
    d = df[df["_plan_year"] == year].copy()
    # 只留乾淨「項目N」碼（丟 項目CAPEX-5 等 pseudo/分攤碼）
    d = d[d["dicj code"].astype(str).str.match(r"^項目\s*\d")]
    if d.empty:
        return None, []
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    d["_sub"] = d.apply(lambda r: r["vertical_label"] if r["ng_scope"] == "gaming" else r["ng_label"], axis=1)
    key = ["ng_scope", "dicj code"]

    def _mode(s):
        m = s.dropna()
        return m.mode().iat[0] if len(m.mode()) else (m.iloc[0] if len(m) else "")
    base = d.groupby(key, dropna=False).agg(
        項目名稱=("project", "first"),
        _sub=("_sub", _mode),
        _ngcode=("ng_code", _mode),
        報告投資金額=("調整前_萬", "sum"),
        潛在調整合計=("調整_萬", "sum"),
        調整後投資金額=("調整後_萬", "sum"),
    )
    cap = d[d["final_capex_opex"] == "Capex"].groupby(key)["調整後_萬"].sum()
    ope = d[d["final_capex_opex"] == "Opex"].groupby(key)["調整後_萬"].sum()
    base["設施建設/資本性支出"] = cap.reindex(base.index).fillna(0)
    base["活動舉辦/營運性支出"] = ope.reindex(base.index).fillna(0)
    adj = d[d["調整一級"].notna()]
    pv = adj.pivot_table(index=key, columns="_adj", values="調整_萬", aggfunc="sum", fill_value=0)
    tab = base.join(pv).fillna(0).reset_index()
    tab = tab.rename(columns={"dicj code": "項目序號"})
    # 計劃 + 完成率（(scope,code) join；博彩/非博彩 collision → code-only fallback）
    if plan:
        codeonly = {}
        for (g, c), v in plan.items():
            codeonly.setdefault(c, v)
        tab["計劃投資金額"] = pd.to_numeric(tab.apply(
            lambda r: plan.get((r["ng_scope"] == "gaming", _norm(r["項目序號"])),
                               codeonly.get(_norm(r["項目序號"]))), axis=1), errors="coerce")
    else:
        tab["計劃投資金額"] = pd.NA
    tab["投資計劃完成率"] = tab.apply(lambda r: _rate(r["報告投資金額"], r["計劃投資金額"]), axis=1)
    tab["潛在調整後投資計劃完成率"] = tab.apply(lambda r: _rate(r["調整後投資金額"], r["計劃投資金額"]), axis=1)
    adj_cols = [c for c in ADJ7 if c in tab.columns]
    other = [c for c in pv.columns if c not in ADJ7]     # 跨年/未預期
    num = ["計劃投資金額", "報告投資金額"] + adj_cols + other + ["潛在調整合計",
          "調整後投資金額", "設施建設/資本性支出", "活動舉辦/營運性支出"]
    for c in num:
        tab[c] = pd.to_numeric(tab[c], errors="coerce").round(1)

    # 排序：博彩先，博彩內 vertical_label（娛樂場→設施設備），非博彩 ng_code(數字)；內部項目N
    def _pn(s):
        m = re.search(r"(\d+(?:\.\d+)?)", str(s)); return float(m.group(1)) if m else 9e9
    def _ngn(s):
        m = re.search(r"(\d+)", str(s)); return int(m.group(1)) if m else 99
    gorder = {"博彩娛樂場優化": 0, "博彩娛樂場場地的優化": 0, "博彩設施設備優化": 1, "博彩設施及設備的優化": 1}
    tab["_scope"] = (tab["ng_scope"] != "gaming").astype(int)          # 博彩=0 先
    tab["_go"] = tab["_sub"].map(lambda s: gorder.get(s, 5))
    tab["_ngn"] = tab["_ngcode"].map(_ngn)
    tab["_pn"] = tab["項目序號"].map(_pn)
    tab = tab.sort_values(["_scope", "_go", "_ngn", "_pn"]).reset_index(drop=True)

    # 砌 rows：section header + data + 範疇小計 + 類型合計
    ALL = ["項目序號", "項目名稱"] + BASE_L + adj_cols + other + TAIL_L
    numcols = ["計劃投資金額", "報告投資金額"] + adj_cols + other + ["潛在調整合計",
              "調整後投資金額", "設施建設/資本性支出", "活動舉辦/營運性支出"]
    rows = []

    def _agg_row(sub, label):
        r = {c: "" for c in ALL}
        r["項目序號"] = label
        for c in numcols:
            r[c] = round(pd.to_numeric(sub[c], errors="coerce").sum(), 1)
        r["投資計劃完成率"] = _rate(r["報告投資金額"], r["計劃投資金額"])
        r["潛在調整後投資計劃完成率"] = _rate(r["調整後投資金額"], r["計劃投資金額"])
        return r

    for scope in [s for s in ["gaming", "non_gaming"] if (tab["ng_scope"] == s).any()]:
        sc = tab[tab["ng_scope"] == scope]
        scope_name = "博彩項目" if scope == "gaming" else "非博彩項目"
        for sub_name, sub in sc.groupby(sc["_sub"], sort=False):
            rows.append({**{c: "" for c in ALL}, "項目序號": f"{scope_name}—{sub_name}"})  # section
            for _, row in sub.iterrows():
                rows.append({c: row.get(c, "") for c in ALL})
            rows.append(_agg_row(sub, f"{scope_name}—{sub_name} 小計"))
        rows.append(_agg_row(sc, f"{scope_name}合計"))

    out_df = pd.DataFrame(rows, columns=ALL)
    return out_df, other


# ── from build_summary_tables ──
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)


# ── from build_summary_tables ──
pd.set_option("display.max_columns", 30)


# ── from build_summary_tables ──
pd.set_option("display.width", 200)


# ── from build_summary_tables ──
BUCKET = {"25": "2025年度投資計劃", "25_24SY": "2024年度計劃期後投資", "25_23SY": "2023年度計劃期後投資"}


# ── from build_summary_tables ──
BUCKET_ORDER = ["2025年度投資計劃", "2024年度計劃期後投資", "2023年度計劃期後投資"]


# ── from build_summary_tables ──
GORDER = {"博彩娛樂場優化": 0, "博彩娛樂場場地的優化": 0, "博彩設施設備優化": 1, "博彩設施及設備的優化": 1}


# ── from build_summary_tables ──
def _ngn(s):
    m = re.search(r"(\d+)", str(s)); return int(m.group(1)) if m else 99


# ── from build_summary_tables ──
def _load(feed: Path, entity: str) -> pd.DataFrame:
    df = pd.read_csv(feed, low_memory=False)
    if entity and "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity]
    df = df[df["dicj code"].astype(str).str.match(r"^項目\s*\d")]     # 丟 pseudo 碼
    df["_yb"] = df["year_bucket"].astype(str).str.strip()
    df["_bucket"] = df["_yb"].map(BUCKET)
    df = df[df["_bucket"].notna()].copy()                             # 只留「於2025發生」3 bucket
    df["_sub"] = df.apply(lambda r: r["vertical_label"] if r["ng_scope"] == "gaming" else r["ng_label"], axis=1)
    df["_scope"] = (df["ng_scope"] != "gaming").astype(int)           # 博彩=0 先
    df["_go"] = df["_sub"].map(lambda s: GORDER.get(s, 5))
    df["_ngn"] = df["ng_code"].map(_ngn)
    return df


# ── from build_summary_tables ──
def _order(agg):
    return agg.sort_values(["_scope", "_go", "_ngn", "_sub"])


# ── from build_summary_tables ──
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
        for _, row in sc.iterrows():
            r = {"範疇": row["_sub"]}
            for c in valcols:
                r[c] = round(float(row[c]), 1)
            rows.append(r)
        rows.append(agg_row(sc, f"{nm}小計"))
    rows.append(agg_row(_order(agg), "總計"))
    return pd.DataFrame(rows, columns=ALL)


# ── from build_summary_tables ──
def summary_amount(df) -> pd.DataFrame:
    """4.1 金額匯總：範疇 × bucket → 報告投資金額 / 潛在調整後投資金額 + 合計。"""
    g = df.groupby(["_scope", "_go", "_ngn", "_sub", "_bucket"], dropna=False).agg(
        報告=("調整前_萬", "sum"), 調整後=("調整後_萬", "sum")).reset_index()
    # pivot bucket → 兩個 measure
    base = g.groupby(["_scope", "_go", "_ngn", "_sub"], dropna=False)
    idx = base.size().reset_index()[["_scope", "_go", "_ngn", "_sub"]]
    out = idx.copy()
    valcols = []
    for bk in BUCKET_ORDER:
        sub = g[g["_bucket"] == bk].set_index(["_scope", "_go", "_ngn", "_sub"])
        for meas, lab in [("報告", "報告投資金額"), ("調整後", "潛在調整後投資金額")]:
            col = f"{bk}·{lab}"
            out[col] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
                lambda k: sub[meas].get(k, 0.0)).astype(float).round(1).values
            valcols.append(col)
    out["合計·報告投資金額"] = out[[f"{b}·報告投資金額" for b in BUCKET_ORDER]].sum(axis=1).round(1)
    out["合計·潛在調整後投資金額"] = out[[f"{b}·潛在調整後投資金額" for b in BUCKET_ORDER]].sum(axis=1).round(1)
    valcols += ["合計·報告投資金額", "合計·潛在調整後投資金額"]
    return _emit(out, valcols)


# ── from build_summary_tables ──
def facility_activity(df, bucket_label) -> pd.DataFrame:
    """4.2 設施vs活動（一個 bucket）：範疇 × 設施建設(capex調整後)/活動舉辦(opex調整後)/合計。"""
    d = df[df["_bucket"] == bucket_label]
    if d.empty:
        return pd.DataFrame()
    cap = d[d["final_capex_opex"] == "Capex"].groupby(["_scope", "_go", "_ngn", "_sub"])["調整後_萬"].sum()
    ope = d[d["final_capex_opex"] == "Opex"].groupby(["_scope", "_go", "_ngn", "_sub"])["調整後_萬"].sum()
    idx = d.groupby(["_scope", "_go", "_ngn", "_sub"]).size().reset_index()[["_scope", "_go", "_ngn", "_sub"]]
    out = idx.copy()
    out["設施建設/資本性支出"] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
        lambda k: cap.get(k, 0.0)).astype(float).round(1).values
    out["活動舉辦/營運性支出"] = out.set_index(["_scope", "_go", "_ngn", "_sub"]).index.map(
        lambda k: ope.get(k, 0.0)).astype(float).round(1).values
    out["合計"] = (out["設施建設/資本性支出"] + out["活動舉辦/營運性支出"]).round(1)
    return _emit(out, ["設施建設/資本性支出", "活動舉辦/營運性支出", "合計"])


# ── from build_overview_tables ──
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)


# ── from build_overview_tables ──
BUCKET_PLANYR = {"2025年度投資計劃": 25, "2024年度計劃期後投資": 24, "2023年度計劃期後投資": 23}


# ── from build_overview_tables ──
def _plan_tot(plan, yr, gaming=None):
    d = (plan or {}).get(yr, {})
    return round(sum(v for (g, c), v in d.items() if gaming is None or g == gaming), 1)


# ── from build_overview_tables ──
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

    plan_by_sub = {}
    if is_py and plan:
        cat = category or {}
        sub_of = {}       # (gaming,碼)→範疇；博彩/非博彩共用項目N 都要留（唔可以 drop 淨 code）
        d2sub = {}        # 項目性質(D)→set(範疇)：由 feed 有嘅項目學，用嚟派零投資項目計劃
        for _, r in d.drop_duplicates(["ng_scope", "dicj code"]).iterrows():
            key = (r["ng_scope"] == "gaming", _norm(r["dicj code"]))
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

    rows = []
    for scope in [0, 1]:
        sc = g[g["_scope"] == scope]
        if sc.empty:
            continue
        name = "博彩項目" if scope == 0 else "非博彩項目"
        rows.append({"範疇": name})     # section 標題行（跟報告 IMG_0105：博彩項目 / 非博彩項目）
        for _, row in sc.iterrows():
            rows.append(mk(row["_sub"], row["項目數量"], plan_by_sub.get(row["_sub"], 0.0),
                           row["報告"], row["調整"], row["後"], row["設施"], row["活動"]))
        rows.append(mk(f"{name}小計", sc["項目數量"].sum(), _plan_tot(plan, yr, scope == 0),
                       sc["報告"].sum(), sc["調整"].sum(), sc["後"].sum(), sc["設施"].sum(), sc["活動"].sum()))
    rows.append(mk("總計", g["項目數量"].sum(), _plan_tot(plan, yr, None),
                   g["報告"].sum(), g["調整"].sum(), g["後"].sum(), g["設施"].sum(), g["活動"].sum()))
    if is_py:
        cols = ["範疇", "項目數量", "獲批的計劃投資金額", "報告投資金額", "投資計劃完成率",
                "潛在調整後投資金額", "潛在調整後投資計劃完成率", "設施建設/資本性支出", "活動舉辦/營運性支出"]
    else:
        cols = ["範疇", "項目數量", "報告投資金額", "潛在調整金額", "潛在調整後投資金額",
                "設施建設/資本性支出", "活動舉辦/營運性支出"]
    return pd.DataFrame(rows)[cols]


# ── from build_overview_tables ──
def adjustment_bridge(df):
    """S15-17：7 canonical 調整類型 × {2025計劃/2024期後/2023期後/合計}。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    rows = []
    for adj in ADJ7:
        r = {"潛在調整事項": adj}
        for bk in BUCKET_ORDER:
            r[bk] = round(d[(d["_bucket"] == bk) & (d["_adj"] == adj)]["調整_萬"].sum(), 1)
        r["合計"] = round(sum(r[bk] for bk in BUCKET_ORDER), 1)
        rows.append(r)
    # 跨年及其他調整（唔喺報告 7 類，多數係期後嘅跨期/將往年計入本年）→ 令 bucket 合計對返概況潛在調整
    other = {"潛在調整事項": "跨年及其他調整"}
    for bk in BUCKET_ORDER:
        tot_bk = round(d[d["_bucket"] == bk]["調整_萬"].sum(), 1)
        other[bk] = round(tot_bk - sum(x[bk] for x in rows), 1)
    other["合計"] = round(sum(other[bk] for bk in BUCKET_ORDER), 1)
    if any(abs(other[bk]) > 0.05 for bk in BUCKET_ORDER):
        rows.append(other)
    tot = {"潛在調整事項": "合計"}
    for bk in BUCKET_ORDER + ["合計"]:
        tot[bk] = round(sum(x[bk] for x in rows), 1)
    rows.append(tot)
    return pd.DataFrame(rows, columns=["潛在調整事項"] + BUCKET_ORDER + ["合計"])


# ── from build_overview_tables ──
def zero_investment_summary(df, plan, cat, narr, ent_up="MGM"):
    """報告概述尾段：2025計劃申報投資為零嘅非博彩項目（原計劃有金額、2025 實際=0）分類。
    → (n, 總計劃_萬, groups)；groups={跨年/內部研究/取消: [範疇…]}。分類靠 期後有支出=跨年、
    項目狀況含取消=取消、其餘=內部研究（best-effort，項目組口徑為準）。"""
    if not plan or not plan.get(25):
        return None
    from collections import Counter

    def _key(r):
        return (str(r["ng_scope"]) == "gaming", _norm(r["dicj code"]))

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


# ── from build_overview_tables ──
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


# ── from build_overview_tables ──
def finding_summary(df):
    """S28-40：每個 canonical 調整類型 → 調整額合計 / 涉及項目數 / 主要涉及項目(top3)。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    rows = []
    for adj in ADJ7:
        sub = d[(d["_adj"] == adj) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        projs = sub.groupby("project")["調整_萬"].sum().abs().sort_values(ascending=False)
        rows.append({"潛在調整事項": adj, "調整額合計": round(sub["調整_萬"].sum(), 1),
                     "涉及項目數": int(sub["dicj code"].nunique()),
                     "主要涉及項目": "、".join(str(p) for p in projs.index[:3])})
    return pd.DataFrame(rows, columns=["潛在調整事項", "調整額合計", "涉及項目數", "主要涉及項目"])


# ── from workbench ──
DEFAULT_PROVIDER = "azure"


# ── from workbench ──
DEFAULT_BASE_URL = "https://api.workbench.kpmg/genai/azure/openai"


# ── from workbench ──
DEFAULT_API_VERSION = "2024-12-01-preview"


# ── from workbench ──
DEFAULT_MODEL = "5.5"


# ── from workbench ──
DEFAULT_MODELS = {
    "5.5": "gpt-5-5-2026-04-24-gs-sdc",
    "5.4": "gpt-5-4-2026-03-05-gs-sdc",
}


# ── from workbench ──
_CRED = Path("conf/local/credentials.yml")


# ── from workbench ──
def _cred_section() -> dict:
    if not _CRED.exists():
        return {}
    try:
        import yaml
        d = yaml.safe_load(_CRED.read_text(encoding="utf-8")) or {}
        return dict(d.get("workbench") or {})
    except Exception:
        return {}


# ── from workbench ──
def _resolve(name: str, env: str, default: str = "") -> str:
    """env > config > default（字串）。"""
    v = os.environ.get(env, "").strip()
    if v:
        return v
    v = str(_cred_section().get(name, "") or "").strip()
    return v or default


# ── from workbench ──
def _parse_json_lenient(text: str) -> dict:
    s = (text or "").strip()
    if s.startswith("```"):
        s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
        s = re.sub(r"\s*```$", "", s).strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", s)
        if not m:
            raise
        return json.loads(m.group(0))


# ── from workbench ──
class Workbench:
    def __init__(self, model: str | None = None, *, provider: str | None = None,
                 charge_code: str | None = None, region: str | None = None,
                 base_url: str | None = None, api_version: str | None = None,
                 api_key: str | None = None, models: dict | None = None):
        sec = _cred_section()
        # alias → 實際名 對照表：內置 ← config.models 覆蓋 ← 呼叫者覆蓋
        self.models = {**DEFAULT_MODELS, **(sec.get("models") or {}), **(models or {})}
        self.provider = (provider or _resolve("provider", "WB_PROVIDER", DEFAULT_PROVIDER)).lower()
        self.model_alias = model or _resolve("model", "WB_MODEL", DEFAULT_MODEL)
        self.base_url = base_url or _resolve("base_url", "WB_BASE_URL", DEFAULT_BASE_URL)
        self.api_version = api_version or _resolve("api_version", "WB_API_VERSION", DEFAULT_API_VERSION)
        self.charge_code = charge_code or _resolve("charge_code", "WB_CHARGE_CODE", "0000")
        self.region = region or _resolve("region", "WB_REGION", "westeurope")
        self._api_key = (api_key or "").strip() or None
        self._client = None

    def resolve_model(self, model: str | None = None) -> str:
        """alias → 實際 deployment/model 名（唔喺對照表就當已經係實名）。"""
        m = model or self.model_alias
        return self.models.get(m, m)

    @property
    def model(self) -> str:
        return self.resolve_model()

    # ── key / headers ────────────────────────────────────────────────
    @property
    def api_key(self) -> str:
        if not self._api_key:
            k = _resolve("api_key", "WB_API_KEY")
            if not k:
                raise RuntimeError(
                    "冇 API key：set env WB_API_KEY，"
                    "或 conf/local/credentials.yml 加 workbench.api_key（gitignored）")
            self._api_key = k
        return self._api_key

    def _headers(self) -> dict:
        # KPMG Workbench(azure) 專用 header；openai-compatible（qwen）唔加，SDK 自己用 api_key。
        if self.provider == "azure":
            return {
                "Ocp-Apim-Subscription-Key": self.api_key,
                "x-kpmg-charge-code": self.charge_code,
                "x-kpmg-region-override": self.region,
            }
        return {}

    def _client_obj(self):
        if self._client is None:
            if self.provider == "azure":
                from openai import AzureOpenAI
                self._client = AzureOpenAI(
                    api_key=self.api_key, base_url=self.base_url,
                    api_version=self.api_version, default_headers=self._headers() or None)
            else:  # openai-compatible（qwen 等）
                from openai import OpenAI
                self._client = OpenAI(
                    api_key=self.api_key, base_url=self.base_url,
                    default_headers=self._headers() or None)
        return self._client

    # ── chat ─────────────────────────────────────────────────────────
    def chat(self, user: str, system: str = "You are a helpful assistant.", *,
             model: str | None = None, reasoning_effort: str | None = "high",
             temperature: float | None = 1, max_tokens: int | None = None,
             json_mode: bool = False) -> str:
        """一問一答，回 content。對唔支持嘅參數自動 drop 再試，唔會因 model/provider 差異炒。"""
        cli = self._client_obj()
        kw: dict[str, Any] = {
            "model": self.resolve_model(model),
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        }
        hdr = self._headers()
        if hdr:
            kw["extra_headers"] = {"Content-Type": "application/json", **hdr}
        if temperature is not None:
            kw["temperature"] = temperature
        if reasoning_effort:
            kw["reasoning_effort"] = reasoning_effort
        if max_tokens:
            kw["max_tokens"] = max_tokens
        if json_mode:
            kw["response_format"] = {"type": "json_object"}

        droppable = ["reasoning_effort", "temperature", "response_format", "max_tokens"]
        for _ in range(len(droppable) + 1):
            try:
                resp = cli.chat.completions.create(**kw)
                return resp.choices[0].message.content or ""
            except Exception as e:
                msg = str(e).lower().replace("_", "")
                hit = next((p for p in droppable if p in kw and p.replace("_", "") in msg), None)
                if hit is None:
                    raise
                kw.pop(hit, None)
        raise RuntimeError("chat 失敗：所有可 drop 參數都試過")

    def chat_json(self, user: str, system: str = "You are a helpful assistant. Reply with JSON only.",
                  **kw) -> dict:
        kw.setdefault("json_mode", True)
        return _parse_json_lenient(self.chat(user, system, **kw))

    def ping(self, model: str | None = None) -> str:
        return self.chat("hi", model=model)

    # ── config 檢視（key 遮蔽）────────────────────────────────────────
    def config_masked(self) -> dict:
        try:
            k = self.api_key
            km = f"{k[:4]}…({len(k)} chars)" if k else "(缺)"
            key_ok = True
        except Exception as e:
            km, key_ok = f"⚠ {e}", False
        return {
            "provider": self.provider,
            "model_alias": self.model_alias, "model(resolved)": self.model,
            "models_map": self.models,
            "base_url": self.base_url, "api_version": self.api_version,
            "charge_code": self.charge_code, "region": self.region,
            "api_key": km, "key_ok": key_ok,
        }


# ── from workbench ──
def _cli() -> None:
    import argparse
    ap = argparse.ArgumentParser(description="LLM client（databook 用，config-driven）")
    ap.add_argument("--model", default=None, help="alias（5.5/5.4/qwen…）或完整名；預設由 config")
    ap.add_argument("--provider", default=None, help="azure / openai（覆蓋 config）")
    ap.add_argument("--ping", action="store_true", help="送 'hi' 驗連線（要喺 KPMG 網內跑）")
    ap.add_argument("--config", action="store_true", help="offline：只印解析到嘅 config（key 遮蔽）")
    ap.add_argument("--ask", help="送一句 user prompt 睇回覆")
    a = ap.parse_args()
    wb = Workbench(model=a.model, provider=a.provider)
    if a.config or not (a.ping or a.ask):
        print("LLM config（offline，唔出網）:")
        for k, v in wb.config_masked().items():
            print(f"  {k}: {v}")
        if not (a.ping or a.ask):
            print("\n→ 連線測試喺 KPMG 網內跑： python -m kpi.lib.workbench --ping")
            return
    if a.ping:
        print("\n--ping →", wb.ping())
    if a.ask:
        print("\n--ask →", wb.chat(a.ask))


# ── from inspect_biao2 ──
try:
    import openpyxl
    from openpyxl.utils import get_column_letter
except ImportError:
    print("✗ pip install openpyxl"); sys.exit(1)


# ── from inspect_biao2 ──
PASSWORD = "dicj_kpmg"


# ── from inspect_biao2 ──
def load_wb(path, password=PASSWORD):
    """開 xlsx；『not a zip file』＝加密 → msoffcrypto 用密碼解。回 openpyxl workbook。"""
    try:
        return openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception:
        pass
    import msoffcrypto
    buf = io.BytesIO()
    with open(path, "rb") as f:
        off = msoffcrypto.OfficeFile(f)
        off.load_key(password=password)
        off.decrypt(buf)
    buf.seek(0)
    return openpyxl.load_workbook(buf, data_only=True, read_only=True)


# ── from inspect_biao2 ──
KEY_HINT = ["投資項目序號", "項目序號及名稱", "項目序號"]


# ── from inspect_biao2 ──
NARR_HINT = ["KPMG分析", "關注事項", "管理層解釋", "調整金額", "調整原因", "分析意見",
             "反饋意見", "調整後金額", "項目分類", "備註", "狀態"]


# ── from inspect_biao2 ──
def _find_header(rows):
    """揾 detail header row（含投資項目序號）；上一行＝group header。回 (grp_r, det_r)。"""
    for ri in range(min(12, len(rows))):
        joined = "".join(str(v) for v in (rows[ri] or []) if v)
        if any(h in joined for h in KEY_HINT):
            return (ri - 1 if ri > 0 else ri), ri
    return None, None


# ── from inspect_biao2 ──
def _ffill(seq):
    out, last = [], ""
    for v in seq:
        s = "" if v is None else str(v).replace("\n", "").strip()
        if s:
            last = s
        out.append(last)
    return out


# ── from inspect_biao2 ──
def inspect_one(path: Path):
    print(f"\n{'#'*84}\n# {path.name}")
    try:
        wb = load_wb(path)
    except ImportError:
        print("  ✗ 加密檔要 msoffcrypto → pip install msoffcrypto-tool"); return
    except Exception as e:
        print(f"  ✗ 開唔到: {type(e).__name__}: {e}"); return
    for sn in wb.sheetnames:
        ws = wb[sn]
        rows = []
        for i, r in enumerate(ws.iter_rows(values_only=True)):
            rows.append(r)
            if i > 800:
                break
        grp_r, det_r = _find_header(rows)
        print(f"\n{'='*72}\n## sheet {sn!r}  約 {len(rows)} 行"
              + (f"  header：組 r{grp_r+1} / 欄 r{det_r+1}" if det_r is not None else "  ⚠揾唔到表頭"))
        if det_r is None:
            for ri in range(min(6, len(rows))):
                cells = [("" if v is None else str(v).replace("\n", " ")[:20]) for v in (rows[ri] or [])]
                print(f"  r{ri}:", " | ".join(c for c in cells if c)[:200])
            continue
        grp = _ffill(rows[grp_r]) if grp_r is not None and grp_r >= 0 else [""] * len(rows[det_r])
        det = [("" if v is None else str(v).replace("\n", "").strip()) for v in rows[det_r]]
        ncol = max(len(grp), len(det))
        ndata = sum(1 for ri in range(det_r + 1, len(rows))
                    if rows[ri] and any(v not in (None, "") for v in rows[ri]))
        print(f"  資料行約 {ndata}；共 {ncol} 欄：")
        for ci in range(ncol):
            g = grp[ci] if ci < len(grp) else ""
            dcol = det[ci] if ci < len(det) else ""
            sample = ""
            for ri in range(det_r + 1, min(det_r + 60, len(rows))):
                row = rows[ri]
                if row and ci < len(row) and row[ci] not in (None, ""):
                    sample = str(row[ci]).replace("\n", " ").strip()[:50]; break
            if not (g or dcol or sample):
                continue
            mark = ""
            if any(h in (dcol) for h in KEY_HINT):
                mark = "  ⟸ KEY(項目)"
            elif any(h in (g + dcol) for h in NARR_HINT):
                mark = "  ⟸ finding/調整"
            print(f"    {get_column_letter(ci+1):>3} | {g[:16]:<16} | {dcol[:24]:<24} | {sample}{mark}")


# ── from biao2 ──
_CODE_RE = re.compile(r"^(項目\s*\d+|[A-Za-z]{1,5}\d+(?:\.\d+)?)$")


# ── from biao2 ──
_JUNK_RE = re.compile(r"^(無新增問題|無|是|否|已回覆|未回覆|不適用|n/?a|請參閱附件)")


# ── from biao2 ──
_ENT_ALIASES = {
    "mgm": ["mgm", "美高梅"],
    "galaxy": ["galaxy", "銀河"],
    "sjm": ["sjm", "澳博", "新葡京"],
    "wynn": ["wynn", "永利"],
    "vml": ["vml", "威尼斯", "金沙"],
    "melco": ["melco", "新濠"],
}


# ── from biao2 ──
def _match_entity(fname, entity):
    fl = fname.lower()
    for a in _ENT_ALIASES.get(entity, [entity]):
        if a.lower() in fl:
            return True
    return False


# ── from biao2 ──
def load_biao2(folder, entity, log=lambda *a: None):
    """{(gaming, 正規化碼): [finding 文字…]}。best-effort，逐檔逐 sheet try。"""
    out = {}
    d = Path(folder)
    if not d.exists():
        log(f"（冇 {folder}）"); return out
    allx = [p for p in sorted(d.rglob("*.xls*")) if not p.name.startswith("~$")]
    files = [p for p in allx if _match_entity(p.name, entity.lower()) and "提供附件" not in p.name]
    log(f"表2 folder {folder}：共 {len(allx)} 個 xls*，match「{entity}」{len(files)} 檔")
    for p in files:
        log(f"  ✓ {p.name}")
    if not files and allx:
        log(f"  ⚠ 一個都 match 唔到！全部檔名：" + " | ".join(p.name for p in allx[:20]))
    for p in files:
        gaming = ("博監局" in p.name)     # 博監局檔 = 博彩項目；其餘範疇 = 非博彩
        try:
            wb = load_wb(p)
        except Exception as e:
            log(f"  ⚠ 開唔到 {p.name}: {e}"); continue
        for sn in wb.sheetnames:
            try:
                ws = wb[sn]
                rows = []
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    rows.append(r)
                    if i > 700:
                        break
                ncol = max((len(r) for r in rows if r), default=0)
                if ncol == 0:
                    continue
                # code 欄 = 匹配 _CODE_RE 最多嘅欄
                best, bestn = None, 0
                for ci in range(ncol):
                    n = sum(1 for r in rows if ci < len(r) and r[ci] is not None
                            and _CODE_RE.match(re.sub(r"\s+", "", str(r[ci]))))
                    if n > bestn:
                        bestn, best = n, ci
                if best is None or bestn < 2:
                    continue
                log(f"  · {p.name}｜{sn}：code欄 col{best}（{bestn}個碼，{'博彩' if gaming else '非博彩'}）")
                for r in rows:
                    if best < len(r) and r[best] is not None \
                            and _CODE_RE.match(re.sub(r"\s+", "", str(r[best]))):
                        code = _norm(r[best])
                        snips = []
                        for c in r:
                            if c is None:
                                continue
                            s = str(c).replace("\n", " ").strip()
                            if len(s) >= 30 and not _JUNK_RE.match(s):
                                snips.append(s[:400])
                        if snips:
                            out.setdefault((gaming, code), [])
                            for s in snips[:4]:
                                if s not in out[(gaming, code)]:
                                    out[(gaming, code)].append(s)
            except Exception as e:
                log(f"  ⚠ {p.name}｜{sn}: {e}")
    log(f"表2：{len(files)} 檔 → {len(out)} 個 (gaming,碼) 有 finding")
    return out


# ── from biao2 ──
def b2look(b2, ng_scope, code, joiner="　"):
    """由 (ng_scope, code) 攞表2 finding 文字（合併），撞號用 exact 先。"""
    g = (ng_scope == "gaming")
    c = _norm(code)
    snips = b2.get((g, c)) or b2.get((not g, c)) or []
    return joiner.join(snips)


# ── from build_llm_narrative ──
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)


# ── from build_llm_narrative ──
SYS_ADJ = ("你係畢馬威（KPMG）投資計劃執行情況審查報告嘅專業撰稿員。用【繁體中文】書面語，"
           "審查報告語氣：精簡、客觀、專業、第三人稱（用『我們』）。"
           "只可根據所提供嘅資料撰寫，嚴禁虛構、誇大或加入未提供嘅事實/數字。"
           "直接寫有嘅內容，切勿寫『未獲提供』『資料不足』等 meta/免責語句。"
           "輸出淨係一段連貫文字（唔好標題/項目符號/開場白/結語），忌冗長。")


# ── from build_llm_narrative ──
SYS_CAT = ("你係為承批公司『2025年度投資執行報告』撰寫『按範疇的項目概況』嘅撰稿員。用【繁體中文】書面語，"
           "站喺投資執行角度：描述該範疇實際投資咗啲乜（可含具體項目、子項目、活動／賽事／音樂會場次），"
           "再講完成率點解係咁。語氣自然、貼近企業投資執行報告，唔好似審計底稿、唔好似機器砌 list。"
           "★嚴禁用審計／調整用語：『剔除』『調減』『申報口徑』『可計入範圍』『超出範圍』『再次申報』"
           "『偏離』『不符合定義』『未在計劃中列示』等 —— 呢啲屬另一節（調整事項），概況絕不出現。"
           "完成率原因只用管理層嘅業務解釋（例：施工進度較預期延遲、實際所需資金低於預算、"
           "進度高於預期、活動如期舉辦、發現結構性問題增加成本），唔好用審計理由。"
           "直接寫有嘅內容，切勿寫『未獲提供』『資料不足』等 meta 語。"
           "輸出淨係一段連貫文字（唔好項目符號／開場白／結語），語句要順，忌逐點堆砌。")


# ── from build_llm_narrative ──
def _adj_prompt(adj_type, amt_wan, projects):
    lines = [f"潛在調整類型：{adj_type}", f"涉及潛在調減金額：約{abs(amt_wan):,.0f}萬澳門元",
             "涉及項目及審查發現（審查底稿表2 為最權威來源，優先採用其跨司裁決及具體內容）："]
    for name, find, mgmt, b2, ruling in projects[:6]:
        seg = f"- 項目「{name}」"
        if b2:      # 表2＝審查底稿，最可信，放最前、俾最多
            seg += f"；【審查底稿表2】{b2[:900]}"
        if find:
            seg += f"；KPMG分析發現：{find[:260]}"
        if mgmt:
            seg += f"；管理層解釋：{mgmt[:200]}"
        if ruling:
            seg += f"；跨司工作組／KPMG裁決（清單）：{ruling[:220]}"
        lines.append(seg)
    ctx = "\n".join(lines)
    return (f"以下係一項『潛在調整事項』嘅底層資料（審查底稿表2 內容最詳盡，可用作事實依據）。"
            f"請寫一段報告摘要（約120-200字），說明該調整類型、金額、主要涉及嘅投資項目同調減原因。"
            f"★用字須跟原報告：如有向跨司工作組諮詢得到嘅回覆，用『跨司工作組』集體稱呼帶出其立場"
            f"（例如『根據我們向跨司工作組諮詢得到的回覆，跨司工作組認為／未同意…』），"
            f"【切勿】逐個司局點名（如社會文化司、旅遊局、文化局），亦【切勿】自創『KPMG最終立場』等標籤。"
            f"最後點出審查建議（通常為建議剔除／調減）。\n\n{ctx}")


# ── from build_llm_narrative ──
def _cat_prompt(sub, rate_pct, content, reason, b2=""):
    ctx = (f"投資範疇：{sub}\n投資計劃金額完成率：{rate_pct}\n"
           f"該範疇實際投資內容（項目清單）：{content[:500]}\n"
           f"管理層變更原因／業務解釋：{reason[:340]}\n"
           f"表2 補充（只可攞嚟豐富『投資內容』，例如子項目／活動場次／金額明細；"
           f"切勿抄佢嘅審計措辭或調整理由）：{b2[:700]}")
    return (f"請為承批公司投資執行報告寫一句『按範疇的項目概況』（約70-140字），"
            f"格式：「{sub}：主要包括……（實際投資咗啲乜，如有具體子項目／活動場次請寫）。"
            f"投資計劃金額完成率為{rate_pct}，主要由於……（管理層業務原因）」。"
            f"完成率原因只用管理層業務解釋，唔好用審計／調整措辭。\n\n{ctx}")


# ── from build_llm_narrative ──
def _gen(wb, prompt, effort, sysp):
    return wb.chat(prompt, sysp, reasoning_effort=effort).strip()


# ── from build_llm_narrative ──
def generate_llm_narrative(feed_path, entity, qingdan, biao2_dir="data/表2",
                           model=None, workers=3, out_path=None, log=print):
    """由 feed + 清單 + 表2 用 Workbench 生成 {adj,cat} 敘述；寫 {entity}_llm_narrative.json，回 dict。
    可被 build_report.py --llm 直接調用（唔使另跑 command）。"""
    wb = Workbench(model=model)
    df = _load(Path(feed_path), entity)
    narr = load_narrative(Path(qingdan)) if qingdan else {}
    b2 = load_biao2(biao2_dir, entity or "", log=log)
    plan = load_plan(Path(qingdan)) if qingdan else None
    cat = load_category(Path(qingdan)) if qingdan else None
    ov = overview_by_bucket(df, "2025年度投資計劃", plan, cat)
    adj = adjustment_bridge(df)
    d = df.copy()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])

    # 併裝 tasks（(kind, key, prompt, effort, sys)）
    pb = BUCKET_ORDER[0]      # 2025計劃 bucket：調整詳述只計 2025年度計劃（期後另計，對返報告）
    tasks = []
    for _, r in adj.iterrows():
        t = r["潛在調整事項"]
        if t in ("合計", "跨年及其他調整"):
            continue
        amt = r.get(pb, 0)
        if not isinstance(amt, (int, float)) or abs(amt) < 0.5:
            continue
        sub = d[(d["_adj"] == t) & (d["_bucket"] == pb) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        projs = []
        for _, pp in sub.drop_duplicates("dicj code").iterrows():
            nr = nlook(narr, pp["ng_scope"], pp["dicj code"])
            b2t = b2look(b2, pp["ng_scope"], pp["dicj code"])
            ruling = "；".join(x for x in (nr.get("跨司回覆", ""), nr.get("KPMG回覆", "")) if x)
            projs.append((str(pp["project"]), nr.get("KPMG分析發現", ""),
                          nr.get("管理層解釋", ""), b2t, ruling))
        tasks.append(("adj", t, _adj_prompt(t, amt, projs), "medium", SYS_ADJ))

    proj = d.groupby(["_sub", "dicj code"])["調整前_萬"].sum().reset_index()
    for _, r in ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].iterrows():
        sub = str(r["範疇"]); rate = r.get("投資計劃完成率")
        if not isinstance(rate, (int, float)) or pd.isna(rate):
            continue
        scope = "gaming" if sub.startswith("博彩") else "non_gaming"
        content = reason = b2t = ""
        for _, pp in proj[proj["_sub"] == sub].sort_values("調整前_萬", ascending=False).iterrows():
            nr = nlook(narr, scope, pp["dicj code"])
            content = content or nr.get("實際投資內容", "")
            reason = reason or nr.get("管理層解釋", "")   # 業務原因；唔用 KPMG分析發現（審計腔）
            if not b2t:
                b2t = b2look(b2, scope, pp["dicj code"])
            if content and reason and b2t:
                break
        tasks.append(("cat", sub, _cat_prompt(sub, f"{rate*100:.1f}%", content, reason, b2t), "low", SYS_CAT))

    log(f"（{entity}）批 {len(tasks)} 個 summary，workers={workers}…")
    out = {"adj": {}, "cat": {}}
    with ThreadPoolExecutor(max_workers=workers) as ex:
        fut = {ex.submit(_gen, wb, p, eff, sysp): (kind, key) for kind, key, p, eff, sysp in tasks}
        for f in as_completed(fut):
            kind, key = fut[f]
            try:
                out[kind][key] = f.result()
                log(f"  ✓ {kind}｜{key[:22]}")
            except Exception as e:
                log(f"  ⚠ {kind}｜{key[:22]}: {type(e).__name__}: {e}")

    outp = Path(out_path) if out_path else Path(f"{entity or 'all'}_llm_narrative.json")
    outp.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"✓ {outp.resolve()}（adj {len(out['adj'])}、cat {len(out['cat'])} 段）")
    return out


# ── from make_report ──
try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("✗ pip install pandas python-pptx openpyxl"); sys.exit(1)


# ── from make_report ──
HDR = RGBColor(0x00, 0x33, 0x8D)


# ── from make_report ──
SEC = RGBColor(0xD9, 0xE1, 0xF2)


# ── from make_report ──
SUB = RGBColor(0xE7, 0xE6, 0xE6)


# ── from make_report ──
TOT = RGBColor(0xBD, 0xD7, 0xEE)


# ── from make_report ──
DARK = RGBColor(0x17, 0x17, 0x1C)


# ── from make_report ──
LIGHT = RGBColor(0xFF, 0xFF, 0xFF)


# ── from make_report ──
CYAN = RGBColor(0x00, 0xB0, 0xD8)


# ── from make_report ──
ENTITY_FULL = {
    "mgm": "美高梅金殿超濠股份有限公司",
    "galaxy": "銀河娛樂場股份有限公司",
    "sjm": "澳門博彩股份有限公司",
    "wynn": "永利渡假村（澳門）股份有限公司",
    "vml": "威尼斯人澳門股份有限公司",
    "melco": "新濠博亞博彩（澳門）股份有限公司",
}


# ── from make_report ──
FEED = "tableau_combined_25.csv"


# ── from make_report ──
def _is_num(v):
    try:
        float(v); return True
    except (TypeError, ValueError):
        return False


# ── from make_report ──
def _load_llm(entity):
    """有 {entity}_llm_narrative.json（build_llm_narrative.py 出）就用 LLM summary。"""
    import json
    p = Path(f"{entity}_llm_narrative.json")
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


# ── from make_report ──
def _coverage_probe(df, qingdan):
    """一次性 coverage 探測（❓頁：主體/KPI/藝術品欄喺唔喺 inputs）→ print 供定 coverage。"""
    def pr(cols, *kw):
        return [c for c in cols if any(k in str(c) for k in kw)] or "無"
    print("  [coverage] feed 主體/執行公司欄:", pr(df.columns, "主體", "執行公司", "控股", "子公司"),
          "｜KPI欄:", pr(df.columns, "KPI", "收入", "留宿", "晚數", "毛收入", "國際"))
    if qingdan:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(qingdan, data_only=True, read_only=True)
            ws = next((wb[s] for s in wb.sheetnames if s.lower().startswith("database")), wb[wb.sheetnames[0]])
            rows = [r for i, r in enumerate(ws.iter_rows(values_only=True)) if i < 6]
            hr = next((i for i, r in enumerate(rows) if r and any("項目類型" in str(c or "") for c in r)), 2)
            hdr = [str(c or "").replace("\n", "") for c in rows[hr]]
            print("  [coverage] 清單 股權/主體欄:", pr(hdr, "股權", "主體", "執行公司", "控股", "子公司"))
            print("  [coverage] 清單 KPI欄:", pr(hdr, "KPI", "留宿", "晚數", "毛收入", "增長率", "國際住客"))
            print("  [coverage] 清單 藝術品欄:", pr(hdr, "藝術品", "展出", "館藏", "拍賣"))
        except Exception as e:
            print("  [coverage] 清單 probe skip:", e)


# ── from make_report ──
def _dump_pptx_text(prs, entity):
    """把生成嘅 pptx 逐版文字（含表格 cell）dump 做 txt → user paste 返做 cross-check vs scan（唔使影相）。"""
    lines = []
    for i, s in enumerate(prs.slides, 1):
        parts = []
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                parts.append(sh.text_frame.text.strip())
            if sh.has_table:
                for row in sh.table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if any(cells):
                        parts.append(" | ".join(cells))
        lines.append(f"\n===== slide {i} =====\n" + "\n".join(parts))
    out = Path(f"{entity}_報告_dump.txt")
    out.write_text("\n".join(lines), encoding="utf-8")
    return out


# ── from make_report ──
SECTIONS = ["2025年度投資計劃執行情況概述", "過往年度投資計劃在2025年繼續執行的審查跟進",
            "本年度審查工作的主要發現", "其他信息", "投資計劃執行報告的六項KPI分析", "附件"]


# ── from make_report ──
def _furniture(prs, slide, section_idx=0):
    """報告版面 furniture：頂 nav tabs（當前 section 加粗 navy）+ 底 KPMG copyright + 初稿 + 頁碼。"""
    slide_w = prs.slide_width / 914400.0
    slide_h = prs.slide_height / 914400.0
    grey = RGBColor(0x8C, 0x8C, 0x8C)
    x = 0.5
    for i, s in enumerate(SECTIONS):
        w = len(s) * 0.088 + 0.15
        tb = slide.shapes.add_textbox(Inches(x), Inches(0.04), Inches(w), Inches(0.2))
        r = tb.text_frame.paragraphs[0].add_run(); r.text = s
        r.font.size = Pt(6.5); r.font.name = "Microsoft JhengHei"
        r.font.bold = (i == section_idx)
        r.font.color.rgb = HDR if i == section_idx else grey
        x += w + 0.1
    ft = slide.shapes.add_textbox(Inches(0.5), Inches(slide_h - 0.34), Inches(slide_w - 1.6), Inches(0.28))
    fr = ft.text_frame.paragraphs[0].add_run()
    fr.text = "© 2026畢馬威會計師事務所 — 澳門特別行政區合夥制事務所。版權所有，不得轉載。"
    fr.font.size = Pt(6); fr.font.color.rgb = grey; fr.font.name = "Microsoft JhengHei"
    pg = slide.shapes.add_textbox(Inches(slide_w - 1.05), Inches(slide_h - 0.34), Inches(0.9), Inches(0.28))
    pr = pg.text_frame.paragraphs[0].add_run()
    pr.text = f"初稿　{len(prs.slides._sldIdLst)}"
    pr.font.size = Pt(8); pr.font.bold = True; pr.font.color.rgb = HDR; pr.font.name = "Microsoft JhengHei"


# ── from make_report ──
def _dark_slide(prs):
    """新增一版深黑底（封面/分隔共用），回 (slide, w, h)。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank)
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid(); rect.fill.fore_color.rgb = DARK; rect.line.fill.background()
    kb = slide.shapes.add_textbox(Inches(0.55), Inches(0.35), Inches(2.2), Inches(0.4))
    kr = kb.text_frame.paragraphs[0].add_run(); kr.text = "KPMG"
    kr.font.size = Pt(20); kr.font.bold = True; kr.font.italic = True
    kr.font.color.rgb = LIGHT; kr.font.name = "Arial"
    return slide, prs.slide_width / 914400.0, prs.slide_height / 914400.0


# ── from make_report ──
def render_cover(prs, entity, date="2026年6月30日"):
    """封面（報告 p1）：深底、KPMG 左上、承批公司全名 + 報告標題 + 初稿 + 事務所/日期。"""
    slide, w, h = _dark_slide(prs)
    full = ENTITY_FULL.get(entity, entity.upper())
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(7.2), Inches(2.6))
    tf = tb.text_frame; tf.word_wrap = True
    for i, line in enumerate([full, "2025年年度投資計劃執行情況審查", "專項工作報告"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = line
        r.font.size = Pt(25); r.font.bold = True; r.font.color.rgb = LIGHT
        r.font.name = "Microsoft JhengHei"
    db = slide.shapes.add_textbox(Inches(0.6), Inches(4.7), Inches(4), Inches(0.5))
    dr = db.text_frame.paragraphs[0].add_run(); dr.text = "初稿"
    dr.font.size = Pt(16); dr.font.bold = True; dr.font.color.rgb = LIGHT; dr.font.name = "Microsoft JhengHei"
    fb = slide.shapes.add_textbox(Inches(0.6), Inches(5.5), Inches(5), Inches(0.8))
    ftf = fb.text_frame
    for i, line in enumerate(["畢馬威會計師事務所", date]):
        p = ftf.paragraphs[0] if i == 0 else ftf.add_paragraph()
        r = p.add_run(); r.text = line
        r.font.size = Pt(11); r.font.color.rgb = LIGHT; r.font.name = "Microsoft JhengHei"


# ── from make_report ──
def divider(prs, title, number="", subitems=None):
    """章節分隔頁（深黑底、大數字+章節標題、子項列表），跟報告 p8/p18 等。"""
    slide, w, h = _dark_slide(prs)
    ny = h * 0.30
    tx = 0.7
    if number:
        nb = slide.shapes.add_textbox(Inches(0.65), Inches(ny - 0.15), Inches(1.3), Inches(1.3))
        nr = nb.text_frame.paragraphs[0].add_run(); nr.text = f"{number}."
        nr.font.size = Pt(44); nr.font.bold = True; nr.font.color.rgb = LIGHT; nr.font.name = "Arial"
        tx = 2.05
    tb = slide.shapes.add_textbox(Inches(tx), Inches(ny), Inches(w - tx - 0.7), Inches(1.5))
    ttf = tb.text_frame; ttf.word_wrap = True
    tr = ttf.paragraphs[0].add_run(); tr.text = title
    tr.font.size = Pt(26); tr.font.bold = True; tr.font.color.rgb = LIGHT; tr.font.name = "Microsoft JhengHei"
    if subitems:
        y = ny + 1.7
        for it in subitems:
            label, page = (it if isinstance(it, (tuple, list)) else (it, ""))
            rb = slide.shapes.add_textbox(Inches(tx), Inches(y), Inches(w - tx - 1.4), Inches(0.32))
            rr = rb.text_frame.paragraphs[0].add_run(); rr.text = label
            rr.font.size = Pt(12); rr.font.color.rgb = RGBColor(0xC8, 0xC8, 0xD0); rr.font.name = "Microsoft JhengHei"
            if page != "":
                pb = slide.shapes.add_textbox(Inches(w - 1.3), Inches(y), Inches(0.8), Inches(0.32))
                pr = pb.text_frame.paragraphs[0].add_run(); pr.text = str(page)
                pr.font.size = Pt(12); pr.font.color.rgb = RGBColor(0xC8, 0xC8, 0xD0); pr.font.name = "Microsoft JhengHei"
            y += 0.42


# ── from make_report ──
def _find(dirp, entity, ext, prefer=None):
    d = Path(dirp)
    if not d.exists():
        return None
    cands = [p for p in sorted(d.rglob("*"))
             if p.suffix.lower() == ext and entity.lower() in p.name.lower()
             and not p.name.startswith("~$")]
    if not cands:
        return None
    for pk in (prefer or []):                    # 例：template 優先揀 2025 果份（唔好揀到 2023 舊報告）
        for p in cands:
            if pk in p.name:
                return p
    return cands[0]


# ── from make_report ──
def _draw_table(slide, df, x, y, max_w, font=7):
    """喺 slide (x,y) 畫 navy 表（單 chunk，caller 自行分頁）。max_w=可用闊(吋)。"""
    cols = list(df.columns)
    grouped = any("·" in c for c in cols)
    ncol = len(cols)
    hrows = 2 if grouped else 1
    widths = [1.9 if c in ("範疇", "項目名稱", "潛在調整事項") else
              (2.6 if c == "主要涉及項目" else 0.92) for c in cols]
    scale = min(1.0, max_w / sum(widths))
    widths = [w * scale for w in widths]

    def hdr(cell, text, align=PP_ALIGN.CENTER):
        _set(cell, text, size=font, bold=True, align=align, color=WHITE, fill=HDR)
    t = slide.shapes.add_table(hrows + len(df), ncol, Inches(x), Inches(y),
                               Inches(sum(widths)), Inches(0.3 * (hrows + len(df)))).table
    for ci, w in enumerate(widths):
        t.columns[ci].width = Inches(w)
    if grouped:
        groups = [c.split("·")[0] if "·" in c else c for c in cols]
        t.cell(0, 0).merge(t.cell(1, 0)); hdr(t.cell(0, 0), "萬澳門元", PP_ALIGN.LEFT)
        ci = 1
        while ci < ncol:
            gg = groups[ci]; cj = ci
            while cj + 1 < ncol and groups[cj + 1] == gg:
                cj += 1
            if cj > ci:
                t.cell(0, ci).merge(t.cell(0, cj))
            hdr(t.cell(0, ci), gg)
            ci = cj + 1
        for ci, c in enumerate(cols[1:], start=1):
            hdr(t.cell(1, ci), c.split("·")[1] if "·" in c else c)
    else:
        hdr(t.cell(0, 0), "萬澳門元", PP_ALIGN.LEFT)
        for ci, c in enumerate(cols[1:], start=1):
            hdr(t.cell(0, ci), c)
    for ri, (_, row) in enumerate(df.iterrows(), start=hrows):
        first = str(row[cols[0]]).strip()
        if all(str(row[c]).strip() == "" for c in cols[1:]):
            for ci in range(ncol):
                _set(t.cell(ri, ci), first if ci == 0 else "", size=font, bold=True,
                       align=PP_ALIGN.LEFT, fill=SEC)
            continue
        is_tot = first.endswith("總計")
        is_sub = is_tot or first.endswith(("小計", "合計"))
        fill = TOT if is_tot else (SUB if is_sub else None)
        for ci, c in enumerate(cols):
            v = row[c]
            if ci == 0:
                txt, al = first, PP_ALIGN.LEFT
            elif "率" in c:
                txt, al = fmt_pct(v), PP_ALIGN.RIGHT
            elif _is_num(v):
                txt, al = fmt_money(v), PP_ALIGN.RIGHT
            else:
                txt, al = ("" if v is None else str(v)), PP_ALIGN.LEFT
            _set(t.cell(ri, ci), txt, size=font, bold=is_sub, align=al,
                   color=(RED if txt.startswith("(") else None), fill=fill)


# ── from make_report ──
def _bullets_into(box, bullets, size=8):
    tf = box.text_frame; tf.word_wrap = True
    for i, (head, body) in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        rm = p.add_run(); rm.text = "■ "; rm.font.size = Pt(size); rm.font.color.rgb = HDR
        if head:
            rh = p.add_run(); rh.text = head
            rh.font.bold = True; rh.font.size = Pt(size); rh.font.color.rgb = HDR
            rh.font.name = "Microsoft JhengHei"
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(size); rt.font.name = "Microsoft JhengHei"


# ── from make_report ──
def render_overview_page(prs, subtitle, headline, table_df, bullets):
    """報告概述式 2 欄版：頂 subtitle + navy headline，左 表，右 敘述。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    slide = prs.slides.add_slide(blank)
    _furniture(prs, slide, 0)
    st = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(slide_w - 0.8), Inches(0.24))
    sp = st.text_frame.paragraphs[0]; sr = sp.add_run(); sr.text = subtitle
    sr.font.size = Pt(9); sr.font.color.rgb = RGBColor(0x60, 0x60, 0x60); sr.font.name = "Microsoft JhengHei"
    if headline:
        hb = slide.shapes.add_textbox(Inches(0.4), Inches(0.56), Inches(slide_w - 0.8), Inches(1.0))
        htf = hb.text_frame; htf.word_wrap = True
        hr = htf.paragraphs[0].add_run(); hr.text = headline
        hr.font.bold = True; hr.font.size = Pt(11); hr.font.color.rgb = HDR; hr.font.name = "Microsoft JhengHei"
    top = 1.5 if headline else 0.6
    left_w = 5.5
    if table_df is not None and not table_df.empty:
        _draw_table(slide, table_df, 0.4, top, left_w, font=6.5)
    rx = 0.4 + left_w + 0.2
    box = slide.shapes.add_textbox(Inches(rx), Inches(top), Inches(slide_w - rx - 0.3), Inches(6.6 - top))
    _bullets_into(box, bullets, size=8)


# ── from make_report ──
def render_generic(prs, title, df):
    """單張表（範疇/項目 + 數字欄；·=2-row group header），自行分頁。"""
    n = len(df); ROWS = 28
    pages = [(i, min(i + ROWS, n)) for i in range(0, n, ROWS)] or [(0, 0)]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    for pi, (a, b) in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
        _set_title(tb, f"{title}（{pi+1}/{len(pages)}）" if len(pages) > 1 else title)
        _draw_table(slide, df.iloc[a:b], 0.4, 0.72, slide_w - 0.8, font=7)


# ── from make_report ──
def _finding_body(box, find, mgmt, grey):
    """body 文字框：KPMG分析發現 / 管理層解釋 兩段，label 加粗（跟原報告用字，唔加報告冇嘅 label 欄）。"""
    tf = box.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(45000)
    tf.margin_top = tf.margin_bottom = Emu(27000)
    first = True
    for label, text, col in [("KPMG分析發現", find, None), ("管理層解釋", mgmt, grey)]:
        if not text:
            continue
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        rl = p.add_run(); rl.text = label + "："
        rl.font.bold = True; rl.font.size = Pt(8); rl.font.name = "Microsoft JhengHei"
        rl.font.color.rgb = HDR if col is None else col
        rt = p.add_run(); rt.text = str(text)[:300]
        rt.font.size = Pt(8); rt.font.name = "Microsoft JhengHei"
        if col is not None:
            rt.font.color.rgb = col


# ── from make_report ──
def render_findings(prs, ent_up, df, narr):
    """③ 主要發現（slide 28-40）：每 canonical 調整類型 → 受影響項目 card
    = navy 標題條(項目+金額) + body(KPMG分析發現/管理層解釋 清單抄字)。每頁 2 個項目。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    grey = RGBColor(0x40, 0x40, 0x40)
    for adj in ADJ7:
        sub = d[(d["_adj"] == adj) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        projs = sub.groupby(["ng_scope", "dicj code"]).agg(名稱=("project", "first"),
                             報告=("調整前_萬", "sum"), 調整=("調整_萬", "sum")).reset_index()
        projs = projs.reindex(projs["調整"].abs().sort_values(ascending=False).index)
        recs = []
        for _, p in projs.iterrows():
            nr = nlook(narr, p["ng_scope"], p["dicj code"])
            recs.append((str(p["dicj code"]), str(p["名稱"]), p["報告"], p["調整"],
                         nr.get("KPMG分析發現", ""), nr.get("管理層解釋", "")))
        pages = [recs[i:i + 2] for i in range(0, len(recs), 2)]
        for pi, page in enumerate(pages):
            slide = prs.slides.add_slide(blank)
            tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
            _set_title(tb, f"{ent_up} 本年度主要發現 — {adj}"
                         + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
            y = 0.8
            for code, name, rep, adjv, find, mgmt in page:
                # navy 標題條（項目 + 金額）
                bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(y),
                                             Inches(slide_w - 0.8), Inches(0.3))
                bar.fill.solid(); bar.fill.fore_color.rgb = HDR; bar.line.fill.background()
                btf = bar.text_frame; btf.word_wrap = True
                btf.margin_left = Emu(54000); btf.margin_top = btf.margin_bottom = Emu(9000)
                br = btf.paragraphs[0].add_run()
                br.text = f"{code}　{name[:32]}　│　報告 {fmt_money(rep)}／潛在調整 {fmt_money(adjv)} 萬澳門元"
                br.font.bold = True; br.font.size = Pt(9)
                br.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); br.font.name = "Microsoft JhengHei"
                # body
                body = slide.shapes.add_textbox(Inches(0.4), Inches(y + 0.32),
                                                Inches(slide_w - 0.8), Inches(2.5))
                _finding_body(body, find, mgmt, grey)
                y += 3.0


# ── from make_report ──
def render_site_visits(prs, ent_up, df, narr, threshold=2000):
    """附件二 現場走訪（slide 93-100）：capex ≥ 2,000萬 設施項目（報告走訪準則），
    配清單 實施地點 + 實際投資內容 做走訪概述。每頁 2 個項目 card。"""
    cap = df[df["final_capex_opex"] == "Capex"].copy()
    cap = cap[cap["dicj code"].astype(str).str.match(r"^項目\s*\d")]
    g = cap.groupby(["ng_scope", "dicj code"]).agg(
        名稱=("project", "first"), 報告=("調整前_萬", "sum")).reset_index()
    g = g[g["報告"] >= threshold]
    if g.empty:
        return
    g["_s"] = (g["ng_scope"] != "gaming").astype(int)
    g = g.sort_values(["_s", "報告"], ascending=[True, False])
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    grey = RGBColor(0x40, 0x40, 0x40)
    recs = []
    for _, p in g.iterrows():
        nr = nlook(narr, p["ng_scope"], p["dicj code"])
        recs.append((str(p["dicj code"]), str(p["名稱"]), p["報告"],
                     nr.get("實施地點", ""), nr.get("實際投資內容", "")))
    pages = [recs[i:i + 2] for i in range(0, len(recs), 2)]
    for pi, page in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
        _set_title(tb, f"{ent_up} 附件二 部分項目的現場走訪情況"
                     + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
        y = 0.8
        for code, name, amt, loc, desc in page:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(y),
                                         Inches(slide_w - 0.8), Inches(0.3))
            bar.fill.solid(); bar.fill.fore_color.rgb = HDR; bar.line.fill.background()
            btf = bar.text_frame; btf.word_wrap = True
            btf.margin_left = Emu(54000); btf.margin_top = btf.margin_bottom = Emu(9000)
            br = btf.paragraphs[0].add_run()
            br.text = f"{code}　{name[:30]}　│　設施建設（資本性支出）{fmt_money(amt)} 萬澳門元"
            br.font.bold = True; br.font.size = Pt(9)
            br.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); br.font.name = "Microsoft JhengHei"
            body = slide.shapes.add_textbox(Inches(0.4), Inches(y + 0.32), Inches(slide_w - 0.8), Inches(2.5))
            tf = body.text_frame; tf.word_wrap = True
            tf.margin_left = tf.margin_right = Emu(45000)
            first = True
            for label, text, col in [("地點", loc, HDR), ("現場走訪概述", desc, None),
                                     ("現場走訪圖片", "〔待插入〕", grey)]:
                if label != "現場走訪圖片" and not text:
                    continue
                p0 = tf.paragraphs[0] if first else tf.add_paragraph()
                first = False
                rl = p0.add_run(); rl.text = label + "："
                rl.font.bold = True; rl.font.size = Pt(8); rl.font.name = "Microsoft JhengHei"
                rl.font.color.rgb = HDR if col is None else col
                rt = p0.add_run(); rt.text = str(text)[:300]
                rt.font.size = Pt(8); rt.font.name = "Microsoft JhengHei"
                if col is not None and label != "地點":
                    rt.font.color.rgb = col
            y += 3.0


# ── from make_report ──
def _prose_slide(prs, title, bullets, headline=None):
    """一版敘述（navy 標題 +（可選）headline 粗體導語 + ■ bullet；bullet=(粗體引子, 內文)）。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    slide = prs.slides.add_slide(blank)
    _furniture(prs, slide, 0)
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(slide_w - 0.8), Inches(0.4))
    _set_title(tb, title)
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.95), Inches(slide_w - 1.0), Inches(5.6))
    tf = box.text_frame; tf.word_wrap = True
    started = False
    if headline:
        p = tf.paragraphs[0]; p.space_after = Pt(10)
        r = p.add_run(); r.text = headline
        r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = HDR
        r.font.name = "Microsoft JhengHei"
        started = True
    for head, body in bullets:
        p = tf.add_paragraph() if started else tf.paragraphs[0]
        started = True
        p.space_after = Pt(8)
        rm = p.add_run(); rm.text = "■ "; rm.font.size = Pt(10); rm.font.color.rgb = HDR
        if head:
            rh = p.add_run(); rh.text = head
            rh.font.bold = True; rh.font.size = Pt(10); rh.font.color.rgb = HDR
            rh.font.name = "Microsoft JhengHei"
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(10); rt.font.name = "Microsoft JhengHei"


# ── from make_report ──
def _headline(ent_up, ov, df, plan):
    """slide 10 整體投資支出概況 → 回 (headline 句, bullets)。全 mechanical 由數字生成。"""
    tot = ov[ov["範疇"] == "總計"]
    if not len(tot):
        return "", []
    r = tot.iloc[0]

    def num(x):
        return float(x) if isinstance(x, (int, float)) and not pd.isna(x) else 0.0
    plan_amt = num(r.get("獲批的計劃投資金額"))
    report_amt = num(r.get("報告投資金額"))
    after_amt = num(r.get("潛在調整後投資金額"))
    adj_amt = report_amt - after_amt
    rate = r.get("投資計劃完成率")
    after_rate = r.get("潛在調整後投資計劃完成率")

    def yi(x):
        return f"{x/10000:.1f}"        # 萬 → 億
    d = df.copy()
    d["_adj"] = pd.to_numeric(d["調整_萬"], errors="coerce").fillna(0)
    codes = d["dicj code"].astype(str)
    n_impl = d[pd.to_numeric(d["調整前_萬"], errors="coerce").fillna(0) != 0]["dicj code"].nunique()
    n_plan = len(plan.get(25, {})) if plan else n_impl
    n_zero = max(n_plan - n_impl, 0)
    n_adj = d[d["_adj"] != 0]["dicj code"].nunique()
    headline = (f"{ent_up} 2025年度原獲批計劃開展{n_plan}個投資項目，涉及計劃投資金額約{yi(plan_amt)}億澳門元；"
                f"{ent_up}提交的投資執行報告顯示2025年度投資金額約{yi(report_amt)}億澳門元"
                f"（投資計劃金額完成率{_pct(rate)}）。本次審查工作識別潛在調減金額約{yi(adj_amt)}億澳門元，"
                f"經潛在調整後的2025年度投資金額約{yi(after_amt)}億澳門元（經調整後投資計劃金額完成率{_pct(after_rate)}）。")
    bullets = [
        ("2025年度獲批的計劃投資金額與報告的投資金額：",
         f"根據{ent_up}提交的2025年度投資計劃方案與投資執行報告，{ent_up}獲批計劃開展{n_plan}個投資項目，"
         f"計劃投資金額約{yi(plan_amt)}億澳門元。投資執行報告顯示實際開展其中{n_impl}個投資項目"
         f"（計劃中有{n_zero}個項目未發生投資金額），報告投資金額約{yi(report_amt)}億澳門元，"
         f"報告投資計劃金額完成率為{_pct(rate)}。"),
        ("2025年度投資支出金額的潛在調整事項：",
         f"我們在本次審查工作中發現，{ent_up}報告投資金額中存在部分投資支出可能不應確認為2025年度計劃的投資支出"
         f"（涉及{n_adj}個投資項目，合計約{yi(adj_amt)}億澳門元）。考慮潛在調減事項後，{ent_up} 2025年度投資支出金額"
         f"約{yi(after_amt)}億澳門元，投資計劃金額完成率應為{_pct(after_rate)}。"),
    ]
    return headline, bullets


# ── from make_report ──
def _pct(v):
    return f"{v*100:.1f}%" if isinstance(v, (int, float)) and not pd.isna(v) else "—"


# ── from make_report ──
def _rate_of(df, name, col):
    r = df[df["範疇"] == name]
    if not len(r) or col not in df.columns:
        return None
    v = r.iloc[0][col]
    return v if isinstance(v, (int, float)) and not pd.isna(v) else None


# ── from make_report ──
def _prose_paginated(prs, title, bullets, per):
    if not bullets:
        return
    pages = [bullets[i:i + per] for i in range(0, len(bullets), per)]
    for pi, page in enumerate(pages):
        t = title + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else "")
        _prose_slide(prs, t, page)


# ── from make_report ──
def _prose_2col(prs, title, bullets, per=12, subtitle=None):
    """報告式 2 欄敘述（每頁 per 個 bullet，左右各半）。subtitle=標題下灰色小註。"""
    if not bullets:
        return
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    colw = (slide_w - 1.0) / 2 - 0.1
    top = 0.95 if subtitle else 0.85
    pages = [bullets[i:i + per] for i in range(0, len(bullets), per)]
    for pi, page in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        _furniture(prs, slide, 0)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(slide_w - 0.8), Inches(0.4))
        _set_title(tb, title + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
        if subtitle:
            sb = slide.shapes.add_textbox(Inches(0.4), Inches(0.68), Inches(slide_w - 0.8), Inches(0.24))
            sr = sb.text_frame.paragraphs[0].add_run(); sr.text = subtitle
            sr.font.size = Pt(8); sr.font.italic = True
            sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70); sr.font.name = "Microsoft JhengHei"
        half = (len(page) + 1) // 2
        lb = slide.shapes.add_textbox(Inches(0.4), Inches(top), Inches(colw), Inches(6.0 - top + 0.85))
        _bullets_into(lb, page[:half], size=8)
        if page[half:]:
            rb = slide.shapes.add_textbox(Inches(0.4 + colw + 0.2), Inches(top), Inches(colw), Inches(6.0 - top + 0.85))
            _bullets_into(rb, page[half:], size=8)


# ── from make_report ──
def render_category_overview(prs, ent_up, ov, df, narr, llm=None):
    """slide 13-14 按範疇的項目概況：LLM summary 優先，否則清單「主要包括…完成率X%…」。"""
    if not narr:
        return
    llm_cat = (llm or {}).get("cat", {})
    d = df.copy()
    d["_sub"] = d.apply(lambda r: r["vertical_label"] if r["ng_scope"] == "gaming" else r["ng_label"], axis=1)
    proj = d.groupby(["_sub", "dicj code"])["調整前_萬"].sum().reset_index()
    cats = ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))]
    bullets = []
    for _, r in cats.iterrows():
        sub = str(r["範疇"]); rate = r.get("投資計劃完成率")
        if not isinstance(rate, (int, float)) or pd.isna(rate):
            continue
        if sub in llm_cat and llm_cat[sub]:               # LLM 寫嘅摘要優先
            txt = llm_cat[sub]
            txt = txt[len(sub) + 1:] if txt.startswith(sub + "：") else txt
            bullets.append((f"{sub}：", txt)); continue
        scope = "gaming" if sub.startswith("博彩") else "non_gaming"
        pr = proj[proj["_sub"] == sub].sort_values("調整前_萬", ascending=False)
        content = reason = ""       # content=清單實際投資內容；reason=清單管理層解釋(變更原因)
        for _, pp in pr.iterrows():
            nr = nlook(narr, scope, pp["dicj code"])
            if not content:
                content = nr.get("實際投資內容", "")
            if not reason:
                reason = nr.get("管理層解釋", "")   # 業務原因；唔用 KPMG分析發現（審計腔）
            if content and reason:
                break
        summ = (content[:90] + "…") if len(content) > 90 else content
        rsn = ("，主要由於" + (reason[:80] + "…" if len(reason) > 80 else reason)) if reason else ""
        body = (f"主要包括{summ}。投資計劃金額完成率為{_pct(rate)}{rsn}。" if summ
                else f"投資計劃金額完成率為{_pct(rate)}{rsn}。")
        bullets.append((f"{sub}：", body))
    _prose_2col(prs, f"{ent_up} 按範疇的項目概況", bullets, 12,
                subtitle="若無特別說明，以下為承批公司2025年度投資執行報告的信息")


# ── from make_report ──
def _adj_detail_bullets(ent_up, adj, df, narr, llm=None):
    """slide 16-17 潛在調整事項詳述 → 回 bullets：LLM summary 優先，否則清單分析發現。"""
    if not narr:
        return []
    llm_adj = (llm or {}).get("adj", {})
    pb = BUCKET_ORDER[0]      # 2025計劃 bucket（唔用合計；期後另計）
    d = df.copy()
    d["_adj"] = d["調整一級"].map(CANON).fillna(d["調整一級"])
    bullets = []
    for _, r in adj.iterrows():
        t = str(r["潛在調整事項"])
        if t in ("合計", "跨年及其他調整"):
            continue
        amt = r.get(pb, 0)
        if not isinstance(amt, (int, float)) or abs(amt) < 0.5:
            continue
        if t in llm_adj and llm_adj[t]:                   # LLM 寫嘅摘要優先
            bullets.append((f"{t}（約{abs(amt):,.0f}萬澳門元）：", llm_adj[t])); continue
        sub = d[(d["_adj"] == t) & (d["_bucket"] == pb) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        names = "、".join(str(x) for x in sub.groupby("dicj code")["project"].first().tolist()[:3])
        reason = ruling = ""
        for _, pp in sub.drop_duplicates("dicj code").iterrows():
            nr = nlook(narr, pp["ng_scope"], pp["dicj code"])
            if not reason:
                reason = nr.get("KPMG分析發現", "") or nr.get("調整事項備註", "")
            if not ruling:
                ruling = nr.get("跨司回覆", "") or nr.get("KPMG回覆", "")
            if reason and ruling:
                break
        r2 = (reason[:150] + "…") if len(reason) > 150 else reason
        rl = ("跨司工作組／KPMG意見：" + (ruling[:90] + "…" if len(ruling) > 90 else ruling)) if ruling else ""
        body = f"主要涉及{names}等項目。{r2}{rl}" if (r2 or rl) else f"主要涉及{names}等項目。"
        bullets.append((f"{t}（約{abs(amt):,.0f}萬澳門元）：", body))
    return bullets


# ── from make_report ──
def _exec_bullets(ent_up, ov):
    """整體執行概況敘述（報告 slide 11）：完成率高/低範疇，全部由 overview 數字生成 → 回 bullets。"""
    g = _rate_of(ov, "博彩項目小計", "投資計劃完成率")
    ng = _rate_of(ov, "非博彩項目小計", "投資計劃完成率")
    tot = _rate_of(ov, "總計", "投資計劃完成率")
    ga = _rate_of(ov, "博彩項目小計", "潛在調整後投資計劃完成率")
    nga = _rate_of(ov, "非博彩項目小計", "潛在調整後投資計劃完成率")
    cat = ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].copy()
    cat = cat[cat["投資計劃完成率"].apply(lambda v: isinstance(v, (int, float)) and not pd.isna(v))]
    high = cat[cat["投資計劃完成率"] >= 1.0].sort_values("投資計劃完成率", ascending=False)
    low = cat[cat["投資計劃完成率"] < 0.5].sort_values("投資計劃完成率")
    high_s = "、".join(f"{r['範疇']}（{_pct(r['投資計劃完成率'])}）" for _, r in high.iterrows()) or "—"
    low_s = "、".join(f"{r['範疇']}（{_pct(r['投資計劃完成率'])}）" for _, r in low.iterrows()) or "—"
    bullets = [
        ("", f"{ent_up}於2025年度計劃投資項目涵蓋博彩及非博彩範疇。在報告投資金額中，"
             f"博彩項目的投資計劃完成率為{_pct(g)}，非博彩項目為{_pct(ng)}，整體為{_pct(tot)}。"),
        ("", f"考慮投資金額的潛在調整後，博彩項目的投資計劃完成率為{_pct(ga)}，"
             f"非博彩項目的平均完成率為{_pct(nga)}。"),
        ("報告投資金額完成率較高的範疇包括：", f"{high_s}。"),
        ("報告投資金額完成率相對較低的範疇包括：", f"{low_s}。"),
    ]
    return bullets


# ── from make_report ──
def _adj_summary(ent_up, adj):
    """潛在調整事項匯總（報告 slide 15）→ 回 (headline, bullets)。逐類型金額。
    ⚠ 用 2025計劃 bucket（唔係合計）：報告調整詳述只計 2025年度計劃，期後另有匯總。"""
    pb = BUCKET_ORDER[0]      # "2025年度投資計劃"
    tot_row = adj[adj["潛在調整事項"] == "合計"]
    total = tot_row.iloc[0][pb] if len(tot_row) else 0
    headline = (f"基於各項審查程序，我們認為{ent_up}報告的2025年度投資金額中存在多類潛在調整事項，"
                f"潛在調減投資金額約{abs(total):,.0f}萬澳門元，主要涉及以下類型：")
    bullets = []
    for _, r in adj.iterrows():
        name = r["潛在調整事項"]
        if name in ("合計", "跨年及其他調整"):
            continue
        amt = r[pb] if pb in adj.columns else 0
        if not isinstance(amt, (int, float)) or abs(amt) < 0.5:
            continue
        bullets.append((f"{name}：", f"約{abs(amt):,.0f}萬澳門元。"))
    return headline, bullets


def main():
    entity = (sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "mgm").lower()
    feed = Path(FEED)
    if not feed.exists():
        print(f"✗ 揾唔到 feed {feed}（root 應有 tableau_combined_25.csv）"); return
    qingdan = _find("data/投資項目清單", entity, ".xlsx")
    template = _find("data/reports", entity, ".pptx", prefer=["2025"])
    ent_up = entity.upper()
    print(f"entity={ent_up}  feed={feed.name}  清單={qingdan.name if qingdan else '(冇)'}  "
          f"template={template.name if template else '(冇→用 13.33x7.5)'}")

    df = pd.read_csv(feed, low_memory=False)
    df = df[df["entity"].astype(str).str.lower() == entity]
    df["報告年"] = pd.to_numeric(df["報告年"], errors="coerce")
    df["_plan_year"] = df["year_bucket"].map(_plan_year)
    plan = load_plan(qingdan) if qingdan else None
    cat = load_category(qingdan) if qingdan else None     # 項目性質(D)→派零投資項目計劃返範疇
    narr = load_narrative(qingdan) if qingdan else {}     # 清單 by-project narrative（抄字）
    if narr:
        print(f"    清單 narrative: {sum(1 for r in narr.values() if r.get('KPMG分析發現'))} 個項目有發現")
    _coverage_probe(df, str(qingdan) if qingdan else None)   # 探 ❓頁（主體/KPI/藝術品）coverage
    # 有 workbench creds 就自動即場生成 LLM 敘述（毋須任何 flag）；冇 creds 靜靜跳過用清單 fallback；--no-llm 強制跳過
    av = sys.argv
    if "--no-llm" not in av:
        model = av[av.index("--model") + 1] if "--model" in av else None
        try:
            has_creds = bool(Workbench(model=model).config_masked().get("key_ok"))
        except Exception:
            has_creds = False
        if has_creds:
            workers = int(av[av.index("--workers") + 1]) if "--workers" in av else 3
            biao2_dir = av[av.index("--biao2") + 1] if "--biao2" in av else "data/表2"
            print("  由 feed+清單+表2 即場生成 LLM 敘述…")
            try:
                generate_llm_narrative(str(feed), entity, str(qingdan) if qingdan else None,
                                       biao2_dir=biao2_dir, model=model, workers=workers)
            except Exception as e:
                print(f"  ⚠ LLM 生成失敗（{type(e).__name__}: {e}）→ 用現有 json / 清單 fallback")
    llm = _load_llm(entity)     # {entity}_llm_narrative.json 有就用 LLM 文字，否則清單 fallback
    if llm:
        print(f"    LLM narrative: adj {len(llm.get('adj', {}))}、cat {len(llm.get('cat', {}))} 段")

    prs = Presentation()
    if template:
        ref = Presentation(str(template))
        prs.slide_width, prs.slide_height = ref.slide_width, ref.slide_height
    else:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    sdf = _load(feed, entity)     # 於2025發生 slice（概述 + 金額匯總 共用）

    render_cover(prs, entity)       # 封面（報告 p1）

    # ① 2025年度投資計劃執行情況概述（報告 slide 8-18）
    divider(prs, "2025年度投資計劃執行情況概述", "1", [
        ("1.1  股權架構簡圖及發生投資支出的主體公司", ""),
        ("1.2  2025年度計劃的整體投資支出概況", ""),
        ("1.3  2025年度投資項目的整體執行概況", ""),
        ("1.4  2025年度投資計劃報告投資金額的潛在調整事項匯總", ""),
    ])
    ov = overview_by_bucket(sdf, "2025年度投資計劃", plan, cat)
    adj = adjustment_bridge(sdf)
    if not ov.empty:      # slide 10-11：表左 + headline/執行敘述右（報告 2 欄式）
        hl, hlb = _headline(ent_up, ov, sdf, plan)
        exb = _exec_bullets(ent_up, ov)
        render_overview_page(prs, f"2025年度投資計劃執行情況概述 | {ent_up} 2025年度計劃的整體投資支出及執行概況",
                             hl, ov.fillna(""), hlb + exb)
        render_category_overview(prs, ent_up, ov, sdf, narr, llm)   # slide 13-14 逐範疇概況（LLM 優先）
        zit = zero_investment_text(zero_investment_summary(sdf, plan, cat, narr, ent_up), ent_up)
        if zit:      # 報告概述尾段：2025計劃申報投資為零嘅項目（跨年/內部研究/取消）
            _prose_slide(prs, f"{ent_up} 2025年度計劃申報投資支出為零的項目",
                         [("", zit[0])] + [("• ", x) for x in zit[1:]])
    ahl, ab = _adj_summary(ent_up, adj)      # slide 15：表左 + 匯總敘述右
    render_overview_page(prs, f"2025年度投資計劃執行情況概述 | {ent_up} 報告投資金額的潛在調整事項匯總",
                         ahl, adj.fillna(""), ab)
    _prose_2col(prs, f"{ent_up} 2025年度報告投資金額的潛在調整事項（詳述）",
                _adj_detail_bullets(ent_up, adj, sdf, narr, llm), 6)   # slide 16-17 詳述（LLM 優先）

    # ② 過往年度投資計劃在2025年繼續執行的審查跟進（報告 slide 19-26）
    divider(prs, "過往年度投資計劃在2025年繼續執行的審查跟進", "2", [
        ("2.1  2024年度投資計劃期後投資金額概覽", ""),
        ("2.2  2024年度投資計劃報告投資金額的潛在調整事項匯總", ""),
        ("2.3  2023年度投資計劃期後投資金額概覽", ""),
        ("2.4  2023年度投資計劃報告投資金額的潛在調整事項匯總", ""),
    ])
    for bk in ["2024年度計劃期後投資", "2023年度計劃期後投資"]:
        ov = overview_by_bucket(sdf, bk, plan, cat)
        if not ov.empty:
            render_generic(prs, f"{ent_up} {bk}金額概覽", ov.fillna(""))

    # ③ 本年度審查工作的主要發現（報告 slide 28-40）
    divider(prs, "本年度審查工作的主要發現", "3")
    fs = finding_summary(sdf)
    if not fs.empty:
        render_generic(prs, f"{ent_up} 主要發現摘要", fs.fillna(""))
    if narr:      # 逐調整類型 × 項目：金額(feed) + 發現/管理層解釋(清單抄字)
        render_findings(prs, ent_up, sdf, narr)

    # ④ 其他信息（報告 slide 42-63）
    divider(prs, "其他信息", "4")
    render_generic(prs, f"{ent_up} 2025年度投資計劃及過往年度期後投資於2025年發生的投資金額匯總",
                   summary_amount(sdf).fillna(""))
    for bk in BUCKET_ORDER:
        fa = facility_activity(sdf, bk)
        if not fa.empty:
            render_generic(prs, f"{ent_up} {bk}區分設施建設/活動舉辦的投資金額", fa.fillna(""))
    for yr in (25, 24, 23):     # 單個項目審查匯總（slide 46-63）
        tab, _ = build_year(df, yr, plan.get(yr) if plan else None)
        if tab is not None and not tab.empty:
            render_sheet(prs, f"報告年{yr}", tab.fillna(""), list(tab.columns))

    # ⑥ 附件二 現場走訪（slide 93-100）
    if narr:
        divider(prs, "附件", "6")
        render_site_visits(prs, ent_up, sdf, narr)

    out = Path(f"{entity}_報告數字表.pptx")
    try:
        prs.save(out)
    except PermissionError:      # 舊檔喺 PowerPoint 開住鎖住 → 改名唔 crash
        import time
        out = Path(f"{entity}_報告數字表_{time.strftime('%H%M%S')}.pptx")
        prs.save(out)
        print(f"⚠ 原檔開住(鎖住)，改存 → {out.name}（開之前記得閂舊 pptx）")
    print(f"✓ {out.resolve()}  共 {len(list(prs.slides))} 頁（概述 + 主要發現 + 金額匯總 + 設施 + 單項審查）")
    if "--dump" in sys.argv:      # 要 cross-check 先加 --dump（慳空間；ok 咗嘅唔使 dump）
        dump = _dump_pptx_text(prs, entity)
        print(f"✓ text dump → {dump.name}（逐版文字，cross-check 用）")


if __name__ == "__main__":
    main()
