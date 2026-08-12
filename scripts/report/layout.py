#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
layout.py — 報告 pptx 版式引擎（KPMG house style）。全部版式常數集中喺呢度，
make_report / render_review_table_pptx 只負責「派數字」，唔再各自砌 formatting。

版式對齊 mgm_2025_report scan（10.83 x 7.5 in、每版：頂 breadcrumb → 灰色「章節 | 子題」
→ navy 粗體導語 → navy caption bar + 表／敘述 → 資料來源 → footer）。

顏色跟 KPMG Visual identity overview（品牌手冊）：
    Primary   KPMG Blue 00338D｜Medium Blue 005EB8｜Light Blue 0091DA
    Secondary Violet 483698｜Purple 470A68｜Light Purple 6D2077｜Green 00A3A1
表格 tint（由 KPMG Blue 派生）：section EEF1F8、小計 D9E1F2、總計 BDD7EE、格線 BFBFBF。
字體：PowerPoint 用 Arial（品牌手冊指定）行數字/英文，中文用微软雅黑。

★ 高度控制：PowerPoint 會自動長高 row 去就內容 → 純靠 row 數分頁一定爆版。
  呢度用 est_lines() 逐 cell 估 wrap 行數 → 逐 row 定實高度 → 按【累積高度】分頁，
  所以永遠唔會超出 slide 底（單項審查「超出 border」嘅正解）。
"""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ── KPMG 品牌色 ──────────────────────────────────────────────────────────
NAVY = RGBColor(0x00, 0x33, 0x8D)          # KPMG Blue（表頭 / 標題 / 導語）
MBLUE = RGBColor(0x00, 0x5E, 0xB8)         # Medium Blue
LBLUE = RGBColor(0x00, 0x91, 0xDA)         # Light Blue
VIOLET = RGBColor(0x48, 0x36, 0x98)
PURPLE = RGBColor(0x47, 0x0A, 0x68)
LPURPLE = RGBColor(0x6D, 0x20, 0x77)
GREEN = RGBColor(0x00, 0xA3, 0xA1)
# 表格 tint（KPMG Blue 派生；報告通篇用呢 3 級）
SECFILL = RGBColor(0xEE, 0xF1, 0xF8)       # 範疇 section 行
SUBTOT = RGBColor(0xD9, 0xE1, 0xF2)        # 小計
TOTAL = RGBColor(0xBD, 0xD7, 0xEE)         # 總計
BORDER = "BFBFBF"                          # 格線（srgbClr hex）
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x22, 0x22, 0x22)           # 內文黑
GREY = RGBColor(0x59, 0x59, 0x59)          # 註 / 資料來源
LGREY = RGBColor(0x8C, 0x8C, 0x8C)         # breadcrumb 非當前
DARK = RGBColor(0x0C, 0x23, 0x3C)          # 封面 / 章節分隔深底

# 負數用括號表示（KPMG palette 冇紅色）→ 唔另外上色。想要紅色改呢個做 RGBColor(0xC0,0,0)。
NEG_COLOR = None

FONT_CN = "微软雅黑"                        # 中文
FONT_NUM = "Arial"                          # 數字 / 英文（品牌手冊：PowerPoint 用 Arial）

SLIDE_W = 10.83                             # 報告 slide 尺寸（scan 量度確認）
SLIDE_H = 7.5

# 版面錨點（吋）
MARGIN = 0.53
CRUMB_Y = 0.13
SUBTITLE_Y = 0.34
HEAD_Y = 0.56
FOOT_Y = 7.16
CONTENT_BOTTOM = 6.98                       # 內容最底（資料來源之上）

SECTIONS = ["2025年度投資計劃執行情況概述", "過往年度投資計劃在2025年繼續執行的審查跟進",
            "本年度審查工作的主要發現", "其他信息", "投資計劃執行報告的六項KPI分析", "附件"]

_CN_RE = None


def _is_cn(ch):
    return "⺀" <= ch <= "鿿" or "＀" <= ch <= "￯" or "　" <= ch <= "〿"


def has_cn(s):
    return any(_is_cn(c) for c in str(s))


def text_w(s, size):
    """估文字闊度（pt）：中文/全形 ≈ 1 em、英數 ≈ 0.52 em。"""
    w = 0.0
    for c in str(s):
        w += size * (1.0 if _is_cn(c) else 0.52)
    return w


def est_lines(s, col_w_in, size, margin_in=0.06):
    """估 wrap 行數（col_w_in = 欄闊吋）。認 \\n 明碼換行。"""
    avail = max((col_w_in - margin_in) * 72.0, 6.0)
    n = 0
    for seg in str(s).split("\n"):
        n += max(1, -(-text_w(seg, size) // avail))     # ceil
    return int(n)


def row_h(cells, widths, size, pad_in=0.045, min_h=0.155):
    """一行嘅需要高度（吋）＝ 最多 wrap 行數 × 行距 + 上下 padding。"""
    lines = 1
    for txt, w in zip(cells, widths):
        lines = max(lines, est_lines(txt, w, size))
    return max(min_h, lines * size * 1.24 / 72.0 + pad_in)


# ── 基本元件 ─────────────────────────────────────────────────────────────
def size_of(prs):
    return prs.slide_width / 914400.0, prs.slide_height / 914400.0


def blank(prs):
    lay = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    return prs.slides.add_slide(lay)


def _tb(slide, x, y, w, h):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    return box


def put(slide, x, y, w, h, text, *, size=8, bold=False, color=INK, align=PP_ALIGN.LEFT,
        italic=False, font=None):
    """一行/一段文字框。"""
    box = _tb(slide, x, y, w, h)
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    r.font.name = font or (FONT_CN if has_cn(text) else FONT_NUM)
    return box


def breadcrumb(slide, W, active=0, entity="MGM"):
    """頂 nav：六大章節，當前 navy 粗體，其餘灰（對 scan 每版頂部）。"""
    box = _tb(slide, MARGIN - 0.23, CRUMB_Y, W - 1.6, 0.18)
    p = box.text_frame.paragraphs[0]
    for i, s in enumerate(SECTIONS):
        if i:
            sep = p.add_run(); sep.text = "  |  "
            sep.font.size = Pt(5.5); sep.font.color.rgb = RGBColor(0xC8, 0xC8, 0xC8)
            sep.font.name = FONT_CN
        r = p.add_run(); r.text = s
        r.font.size = Pt(5.5); r.font.name = FONT_CN
        r.font.bold = (i == active)
        r.font.color.rgb = NAVY if i == active else LGREY
    put(slide, W - 1.05, CRUMB_Y, 0.85, 0.18, f"{entity}  ◀ ⌂ ▶", size=5.5,
        color=LGREY, align=PP_ALIGN.RIGHT)


def footer(slide, W, H, page):
    """底：KPMG 字標 + 版權 + 初稿/頁碼（對 scan）。"""
    kb = _tb(slide, MARGIN - 0.23, H - 0.34, 0.7, 0.22)
    kr = kb.text_frame.paragraphs[0].add_run(); kr.text = "KPMG"
    kr.font.size = Pt(11); kr.font.bold = True; kr.font.italic = True
    kr.font.color.rgb = NAVY; kr.font.name = FONT_NUM
    put(slide, MARGIN + 0.5, H - 0.30, W - 2.2, 0.2,
        "© 2026畢馬威會計師事務所 — 澳門特別行政區合夥制事務所。版權所有，不得轉載。",
        size=5, color=LGREY)
    if page is not None:
        put(slide, W - 1.15, H - 0.32, 0.95, 0.2, f"初稿　{page}", size=7, bold=True,
            color=NAVY, align=PP_ALIGN.RIGHT)


MAX_HEAD_H = 1.05      # 導語最多食呢咁多高（scan 一般 2-4 行）；再長就縮字，唔可以食晒成版


def head_h(headline, W, hsize=8.5):
    """導語需要嘅高度 + 實際字號（長就自動縮到 MAX_HEAD_H 為止）→ (h, size)。"""
    if not headline:
        return 0.06, hsize
    while hsize > 6.0:
        h = est_lines(headline, W - 2 * MARGIN, hsize) * hsize * 1.35 / 72.0
        if h <= MAX_HEAD_H:
            return h, hsize
        hsize -= 0.5
    return MAX_HEAD_H, hsize


def page_head(slide, W, crumb, headline=None, *, hsize=8.5):
    """灰色「章節 | 子題」+ navy 粗體導語 → 回內容起始 y。"""
    put(slide, MARGIN, SUBTITLE_Y, W - 2 * MARGIN, 0.2, crumb, size=8.5, bold=True, color=NAVY)
    if not headline:
        return HEAD_Y + 0.06
    h, hsize = head_h(headline, W, hsize)
    box = _tb(slide, MARGIN, HEAD_Y, W - 2 * MARGIN, h)
    p = box.text_frame.paragraphs[0]
    r = p.add_run(); r.text = str(headline)
    r.font.size = Pt(hsize); r.font.bold = True; r.font.color.rgb = NAVY; r.font.name = FONT_CN
    return HEAD_Y + h + 0.10


def caption_bar(slide, x, y, w, text, *, size=6):
    """表頂 navy caption bar（重覆表名，對 scan 每張表都有）。"""
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.17))
    bar.fill.solid(); bar.fill.fore_color.rgb = NAVY
    bar.line.fill.background(); bar.shadow.inherit = False
    tf = bar.text_frame
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.margin_left = Emu(36000); tf.margin_right = Emu(18000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = True; r.font.color.rgb = WHITE; r.font.name = FONT_CN
    return y + 0.17


def source_note(slide, W, y=None, *, note=None, more=False):
    """表下：資料來源（左）+（下頁待續）（右）。"""
    y = CONTENT_BOTTOM if y is None else y
    put(slide, MARGIN, y, W - 2.0, 0.16,
        note or "資料來源：管理層提供之項目投資計劃及執行報告資料，畢馬威分析", size=6, color=GREY)
    if more:
        put(slide, W - MARGIN - 1.2, y, 1.2, 0.16, "（下頁待續）", size=6, color=GREY,
            align=PP_ALIGN.RIGHT)


# ── 表格 ─────────────────────────────────────────────────────────────────
def _borders(cell):
    """幼灰格線；插喺 tcPr 最前（OOXML schema：ln* 要喺 fill 之前，否則 PowerPoint 會要求修復）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):        # 反序 insert(0) → 出嚟係 L,R,T,B
        ln = tcPr.makeelement(qn(tag), {"w": "3175", "cap": "flat"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:srgbClr"), {"val": BORDER})
        fill.append(clr); ln.append(fill)
        tcPr.insert(0, ln)


def set_cell(cell, text, *, size=6, bold=False, fill=None, align=PP_ALIGN.RIGHT,
             color=None, wrap=True):
    cell.margin_left = cell.margin_right = Emu(18000)
    cell.margin_top = cell.margin_bottom = Emu(9000)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill if fill is not None else WHITE
    tf = cell.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    txt = "" if text is None else str(text)
    if color is None:
        color = NEG_COLOR if (NEG_COLOR is not None and txt.startswith("(")) else INK
    # ⚠ DrawingML 入面 "\n" 唔係換行（會當空白）→ 一定要用 <a:br/>，否則表頭喺 PowerPoint 會擠成一行
    for i, seg in enumerate(txt.split("\n")):
        if i:
            p.add_line_break()
        r = p.add_run(); r.text = seg
        r.font.size = Pt(size); r.font.bold = bold
        r.font.name = FONT_CN if has_cn(seg) else FONT_NUM
        r.font.color.rgb = color


ROW_FILL = {"sec": SECFILL, "subtot": SUBTOT, "tot": TOTAL, "data": None}


def header_h(supers, subs, widths, hfont):
    """表頭需要高度（吋）。"""
    h = row_h(subs, widths, hfont, pad_in=0.05, min_h=0.20)
    return (0.17 + h) if supers else h


def fit_rows(rows, widths, font, avail_h, hh):
    """按【累積高度】切頁 → 保證唔會超出可用高度。rows = [(kind, cells)]。
    keep=True 嘅 row（範疇 block）盡量唔拆：見 fit_blocks。"""
    out, cur, used = [], [], 0.0
    cap = max(avail_h - hh, 0.6)          # guard：導語太長時唔好變 0/負數（會無限開版）
    for kind, cells in rows:
        h = row_h(cells, widths, font)
        if cur and used + h > cap:
            out.append(cur); cur, used = [], 0.0
        cur.append((kind, cells)); used += h
    if cur:
        out.append(cur)
    return out or [[]]


def fit_blocks(blocks, widths, font, avail_h, hh):
    """block = 一個範疇（section + data + 小計）。整個 block 唔拆頁（對 scan：全報告冇「續」）；
    單一 block 大過一版先逼住切。"""
    cap = max(avail_h - hh, 0.6)          # guard：同上
    pages, cur, used = [], [], 0.0
    for blk in blocks:
        bh = sum(row_h(c, widths, font) for _, c in blk)
        if cur and used + bh > cap:
            pages.append(cur); cur, used = [], 0.0
        if bh > cap:                                    # 單一範疇爆版 → 逐行切（安全網）
            for kind, cells in blk:
                h = row_h(cells, widths, font)
                if cur and used + h > cap:
                    pages.append(cur); cur, used = [], 0.0
                cur.append((kind, cells)); used += h
            continue
        cur.extend(blk); used += bh
    if cur:
        pages.append(cur)
    return pages or [[]]


def draw_table(slide, x, y, w, subs, rows, widths, *, supers=None, font=6, hfont=5.5,
               left_cols=1, fill_h=None, max_row_h=0.34):
    """畫 navy 表。subs=欄名（可含 \\n）；rows=[(kind, cells)]；widths=相對闊度（會 scale 到 w）。
    supers=[(label, c0, c1_exclusive)] 兩層表頭。fill_h=想填滿嘅高度（行數少時撐開行高，
    唔好剩一大橛白位；每行最多 max_row_h）。回 (bottom_y, 實際高度)。"""
    ncol = len(subs)
    scale = w / sum(widths)
    wid = [v * scale for v in widths]
    nhdr = 2 if supers else 1
    heights = [row_h(cells, wid, font) for _, cells in rows]
    hsub = row_h(subs, wid, hfont, pad_in=0.05, min_h=0.20)
    if fill_h and heights:
        slack = fill_h - ((0.17 if supers else 0) + hsub + sum(heights))
        if slack > 0.05:
            add = min(slack / len(heights), max(0.0, max_row_h - max(heights)))
            heights = [h + add for h in heights]
    total_h = (0.17 if supers else 0) + hsub + sum(heights)
    tbl = slide.shapes.add_table(nhdr + len(rows), ncol, Inches(x), Inches(y),
                                 Inches(w), Inches(total_h)).table
    tbl.first_row = False; tbl.horz_banding = False
    for i, v in enumerate(wid):
        tbl.columns[i].width = Inches(v)
    if supers:
        for c in range(ncol):
            set_cell(tbl.cell(0, c), "", size=hfont, fill=NAVY, color=WHITE)
        for label, c0, c1 in supers:
            if c1 - c0 > 1:
                tbl.cell(0, c0).merge(tbl.cell(0, c1 - 1))
            if label:
                set_cell(tbl.cell(0, c0), label, size=hfont + 0.5, bold=True, fill=NAVY,
                         color=WHITE, align=PP_ALIGN.CENTER)
        tbl.rows[0].height = Emu(int(0.17 * 914400))
    for c, s in enumerate(subs):
        set_cell(tbl.cell(nhdr - 1, c), s, size=hfont, bold=True, fill=NAVY, color=WHITE,
                 align=PP_ALIGN.LEFT if c < left_cols else PP_ALIGN.CENTER)
    tbl.rows[nhdr - 1].height = Emu(int(hsub * 914400))
    for ri, (kind, cells) in enumerate(rows, start=nhdr):
        fill = ROW_FILL.get(kind)
        bold = kind in ("sec", "subtot", "tot")
        for c, v in enumerate(cells):
            al = PP_ALIGN.LEFT if c < left_cols else PP_ALIGN.RIGHT
            set_cell(tbl.cell(ri, c), v, size=font, bold=bold, fill=fill, align=al)
        tbl.rows[ri].height = Emu(int(heights[ri - nhdr] * 914400))
    for row in tbl.rows:
        for cell in row.cells:
            _borders(cell)
    return y + total_h, total_h


# ── 敘述 ─────────────────────────────────────────────────────────────────
def prose(box, items, *, head_size=7, body_size=6.5, gap=6):
    """scan 敘述格式：navy 粗體小標題一行 + 下面 body 段落（唔用 ■ bullet）。
    items = [(head, body)]；head 可為空。"""
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for head, body in items:
        if head:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.space_before = Pt(0 if first else gap); p.space_after = Pt(1)
            r = p.add_run(); r.text = str(head)
            r.font.size = Pt(head_size); r.font.bold = True
            r.font.color.rgb = NAVY; r.font.name = FONT_CN
            first = False
        if body:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.space_before = Pt(0 if first else (0 if head else gap)); p.space_after = Pt(1)
            r = p.add_run(); r.text = str(body)
            r.font.size = Pt(body_size); r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
            r.font.name = FONT_CN
            first = False


def prose_box(slide, x, y, w, h, items, **kw):
    box = _tb(slide, x, y, w, h)
    prose(box, items, **kw)
    return box


def est_prose_h(items, w, head_size=7, body_size=6.5, gap=6):
    """估敘述高度（吋）→ 用嚟分頁，唔會爆版。"""
    h = 0.0
    for head, body in items:
        if head:
            h += est_lines(head, w, head_size) * head_size * 1.3 / 72.0 + gap / 72.0
        if body:
            h += est_lines(body, w, body_size) * body_size * 1.35 / 72.0 + 2 / 72.0
    return h


def fit_prose(items, w, avail_h, **kw):
    """按估算高度切頁 → [[items]]。"""
    pages, cur, used = [], [], 0.0
    avail_h = max(avail_h, 0.6)           # guard
    for it in items:
        ih = est_prose_h([it], w, **kw)
        if cur and used + ih > avail_h:
            pages.append(cur); cur, used = [], 0.0
        cur.append(it); used += ih
    if cur:
        pages.append(cur)
    return pages or [[]]


# ── 深色版（封面 / 章節分隔）─────────────────────────────────────────────
def dark_slide(prs):
    slide = blank(prs)
    W, H = size_of(prs)
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid(); rect.fill.fore_color.rgb = DARK
    rect.line.fill.background(); rect.shadow.inherit = False
    kb = _tb(slide, 0.55, 0.35, 2.2, 0.4)
    kr = kb.text_frame.paragraphs[0].add_run(); kr.text = "KPMG"
    kr.font.size = Pt(20); kr.font.bold = True; kr.font.italic = True
    kr.font.color.rgb = WHITE; kr.font.name = FONT_NUM
    return slide, W, H
