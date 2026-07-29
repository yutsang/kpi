#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
render_review_table_pptx.py — 把 build_project_review_table 出嘅 xlsx render 落 pptx
native table（報告 slide 46-63「單個項目審查結果匯總表」，取代 Tableau 截圖）。

策略（user 2026-07-21）：**data 行先**（行/欄/範疇/小計/數字擺正 + 分頁），
formatting 顏色後補；**尺寸跟原報告**（--template 讀 slide 尺寸）。

用法（Windows，kpi-main 底下）：
    pip install python-pptx openpyxl
    python scripts\\report\\render_review_table_pptx.py mgm_項目審查匯總.xlsx ^
        --template "data\\reports\\MGM.2025年年度投資計劃執行情況審查專項工作報告.初稿.pptx"
  出 mgm_項目審查匯總.native.pptx（開嚟同報告 slide 46-63 對）+ console 印每 sheet 分幾頁、欄。
"""
import sys
from pathlib import Path

try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
except ImportError:
    print("✗ 需要 pandas + python-pptx → pip install pandas python-pptx openpyxl"); sys.exit(1)

RATE_COLS = {"投資計劃完成率", "潛在調整後投資計劃完成率"}
TEXT_COLS = {"項目序號", "項目名稱"}
# 3 欄組（畫 group header 用）
G1 = ["項目序號", "項目名稱", "計劃投資金額", "報告投資金額", "投資計劃完成率"]
G3 = ["調整後投資金額", "潛在調整後投資計劃完成率", "設施建設/資本性支出", "活動舉辦/營運性支出"]
GROUP_LABEL = {"G1": "項目基本信息", "G2": "投資金額的潛在調整事項", "G3": "潛在調整後投資金額"}

BLUE = RGBColor(0x00, 0x33, 0x8D)        # KPMG 藍（標題）
LBLUE = RGBColor(0xD9, 0xE1, 0xF2)       # section 淺藍
GREY = RGBColor(0xE7, 0xE6, 0xE6)        # 小計灰
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x00, 0x00)
# 3 色欄組（more or less 跟報告 IMG）：基本信息=青、潛在調整事項=淺藍、潛在調整後=深藍
GROUP_FILL = {"G1": RGBColor(0x2E, 0x9B, 0xD6), "G2": RGBColor(0x9D, 0xC3, 0xE6),
              "G3": RGBColor(0x1F, 0x38, 0x64)}
GROUP_TEXT = {"G1": WHITE, "G2": RGBColor(0x1F, 0x38, 0x64), "G3": WHITE}

ROWS_PER_SLIDE = 24                       # 每頁資料行（含 section/小計），近報告
YEAR_TITLE = {
    "報告年25": "MGM 2025年度投資計劃單個項目審查結果匯總表",
    "報告年24": "MGM 2024年度投資計劃單個項目截至2025年末的審查結果匯總表",
    "報告年23": "MGM 2023年度投資計劃單個項目截至2025年末的審查結果匯總表",
}


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


def fmt_pct(v):
    if v is None or v == "" or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return f"{float(v)*100:.1f}%"
    except (TypeError, ValueError):
        return str(v)


def row_kind(seq: str) -> str:
    s = str(seq)
    if s.endswith("小計") or s.endswith("合計"):
        return "subtotal"
    if "—" in s or "－" in s:
        return "section"
    return "data"


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
    f.name = "微软雅黑"
    if color is not None:
        f.color.rgb = color


def col_group(c):
    if c in G1:
        return "G1"
    if c in G3:
        return "G3"
    return "G2"


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


def _set_title(tb, text):
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(13); r.font.bold = True
    r.font.color.rgb = BLUE; r.font.name = "微软雅黑"


def main():
    args = sys.argv[1:]
    template = None
    if "--template" in args:
        i = args.index("--template"); template = args[i + 1]; del args[i:i + 2]
    if not args:
        print("俾 build_project_review_table 出嘅 xlsx（--template <報告.pptx> 跟 slide 尺寸）"); return
    xlsx = Path(args[0])
    sheets = pd.read_excel(xlsx, sheet_name=None)

    # fresh 包（唔可以 clone template 再刪 slide：sldId 刪咗但 slide part 仲喺 → 撞名 corrupt）
    prs = Presentation()
    if template and Path(template).exists():
        ref = Presentation(template)
        prs.slide_width = ref.slide_width
        prs.slide_height = ref.slide_height
        print(f"── 跟 template 尺寸: {prs.slide_width/914400:.2f}x{prs.slide_height/914400:.2f}in（fresh 包避免撞名）")
    else:
        prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
        print("── 冇 template，用 13.33x7.5in（想跟報告尺寸請 --template 報告.pptx）")

    for sn, df in sheets.items():
        df = df.fillna("")
        cols = list(df.columns)
        print(f"\n# sheet {sn}：{len(df)} 行，{len(cols)} 欄")
        print("  欄:", cols)
        render_sheet(prs, sn, df, cols)

    out = xlsx.with_suffix(".native.pptx")
    prs.save(out)
    print(f"\n✓ 寫入 {out.resolve()}（開嚟同報告 slide 46-63 對；顏色/精細格式後補）")


if __name__ == "__main__":
    main()
