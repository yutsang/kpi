#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
demo_page.py — 自足 demo：淨用 code 砌「概述 整體執行概況」1 版，模仿 scan
（2 層表頭、navy、小計/總計 shaded、數字右對齊、2 欄：表左+敘述右、breadcrumb、footer）。
數據讀 gitignored results\\project_dump.tsv（inspect_project.py 出，報告年25 aggregate 到範疇）；
冇檔就用明顯假數（本檔【零真實金額】，可 commit）。出 demo_overview.pptx。

用法：python scripts\\report\\inspect_project.py ... 先出 tsv，再 python scripts\\report\\demo_page.py
"""
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

NAVY = RGBColor(0x00, 0x33, 0x8D)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SUBTOT = RGBColor(0xD9, 0xE1, 0xF2)
TOT = RGBColor(0xBD, 0xD7, 0xEE)
SECHDR = RGBColor(0xEE, 0xF1, 0xF8)
GREY = RGBColor(0x7F, 0x7F, 0x7F)
FONT = "微软雅黑"

SUB = ["（萬澳門元）", "項目\n數量", "獲批的計劃\n投資金額", "報告\n投資金額", "投資計劃\n完成率",
       "潛在調整後\n投資金額", "潛在調整後\n完成率", "設施建設/\n資本性支出", "活動舉辦/\n營運性支出"]
# 博彩範疇擺前（其餘當非博彩）；顯示順序
GAMING_SUBS = ["博彩娛樂場優化", "博彩設施設備優化"]
NONG_ORDER = ["吸引外國客源", "會議展覽", "娛樂表演", "體育盛事", "文化藝術", "健康養生",
              "主題遊樂", "美食之都", "社區旅遊", "海上旅遊", "其他"]

# fallback 假數（零真實金額；只為展示版式）
FAKE = [
    ("gaming", "博彩娛樂場優化", 2, 1000, 2000, 2000, 1900, 100),
    ("gaming", "博彩設施設備優化", 2, 1000, 1500, 1500, 1400, 100),
    ("non_gaming", "吸引外國客源", 3, 2000, 1800, 900, 100, 1700),
    ("non_gaming", "會議展覽", 2, 1500, 600, 590, 300, 290),
    ("non_gaming", "娛樂表演", 2, 3000, 4000, 1000, 100, 900),
    ("non_gaming", "其他", 2, 500, 300, 280, 0, 280),
]


def _fmt(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v)
    if n == 0:
        return "-"
    return f"({abs(n):,.0f})" if n < 0 else f"{n:,.0f}"


def _rate(rep, pl):
    return f"{rep / pl * 100:.1f}%" if pl else "-"


def load_agg():
    """讀 results/project_dump.tsv（報告年25）aggregate 到 (scope,範疇)；冇就用假數。回 dict。"""
    p = Path("results/project_dump.tsv")
    recs = []
    if p.exists():
        lines = p.read_text(encoding="utf-8").splitlines()
        hdr = lines[0].split("\t")
        ix = {h: i for i, h in enumerate(hdr)}
        for ln in lines[1:]:
            f = ln.split("\t")
            if not f or len(f) < len(hdr):
                continue
            if f[ix["報告年"]] != "25":
                continue
            recs.append((f[ix["scope"]], f[ix["範疇"]], 1,
                         float(f[ix["計劃"]] or 0), float(f[ix["調整前"]] or 0),
                         float(f[ix["調整後"]] or 0), float(f[ix["設施"]] or 0), float(f[ix["活動"]] or 0)))
        src = "results/project_dump.tsv（真數）"
    else:
        recs = FAKE
        src = "假數（fallback；run inspect_project.py 出真數）"
    agg = {}
    for scope, sub, n, pl, rep, aft, fac, act in recs:
        a = agg.setdefault((scope, sub), [0, 0, 0, 0, 0, 0])
        a[0] += n; a[1] += pl; a[2] += rep; a[3] += aft; a[4] += fac; a[5] += act
    return agg, src


def _rows(agg):
    """agg → 顯示 ROWS：[(label, kind, values[8] or None)]。"""
    def line(scope, subs, label_tot):
        out = []
        tot = [0, 0, 0, 0, 0, 0]
        for sub in subs:
            a = agg.get((scope, sub))
            if not a:
                continue
            n, pl, rep, aft, fac, act = a
            out.append((sub, "data", [str(n), _fmt(pl), _fmt(rep), _rate(rep, pl),
                                      _fmt(aft), _rate(aft, pl), _fmt(fac), _fmt(act)]))
            for i in range(6):
                tot[i] += a[i]
        n, pl, rep, aft, fac, act = tot
        sub_row = (label_tot, "subtot", [str(n), _fmt(pl), _fmt(rep), _rate(rep, pl),
                                         _fmt(aft), _rate(aft, pl), _fmt(fac), _fmt(act)])
        return out, sub_row, tot

    rows = [("博彩項目", "sec", None)]
    g_rows, g_sub, g_tot = line("gaming", GAMING_SUBS, "博彩項目小計")
    rows += g_rows + [g_sub]
    rows.append(("非博彩項目", "sec", None))
    ng_rows, ng_sub, ng_tot = line("non_gaming", NONG_ORDER, "非博彩項目小計")
    rows += ng_rows + [ng_sub]
    grand = [g_tot[i] + ng_tot[i] for i in range(6)]
    n, pl, rep, aft, fac, act = grand
    rows.append(("總計", "tot", [str(n), _fmt(pl), _fmt(rep), _rate(rep, pl),
                                 _fmt(aft), _rate(aft, pl), _fmt(fac), _fmt(act)]))
    # 簡單敘述（由數計，非機密）
    cats = [(s, agg[(sc, s)]) for sc in ("gaming", "non_gaming") for s in (GAMING_SUBS if sc == "gaming" else NONG_ORDER) if (sc, s) in agg]
    hi = [f"{s}（{a[2]/a[1]*100:.1f}%）" for s, a in cats if a[1] and a[2] / a[1] >= 1.0]
    lo = [f"{s}（{a[2]/a[1]*100:.1f}%）" for s, a in cats if a[1] and a[2] / a[1] < 0.5]
    bullets = [
        ("整體概況：", f"博彩項目完成率 {_rate(g_tot[2], g_tot[1])}，非博彩項目 {_rate(ng_tot[2], ng_tot[1])}，"
                     f"整體 {_rate(grand[2], grand[1])}。"),
        ("潛在調整後：", f"博彩項目 {_rate(g_tot[3], g_tot[1])}，非博彩項目平均 {_rate(ng_tot[3], ng_tot[1])}。"),
        ("完成率較高的範疇：", "、".join(hi) + "。" if hi else "—"),
        ("完成率相對較低的範疇：", "、".join(lo) + "。" if lo else "—"),
    ]
    return rows, bullets


def _set_cell(cell, text, *, size=5.5, bold=False, fill=None, align=PP_ALIGN.RIGHT, white=False):
    cell.margin_left = cell.margin_right = Emu(18000)
    cell.margin_top = cell.margin_bottom = Emu(2500)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.fill.solid(); cell.fill.fore_color.rgb = fill if fill is not None else WHITE
    tf = cell.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = str(text)
    r.font.size = Pt(size); r.font.bold = bold; r.font.name = FONT
    r.font.color.rgb = WHITE if white else RGBColor(0x22, 0x22, 0x22)


def _thin_borders(table):
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
                ln = tcPr.makeelement(qn(tag), {"w": "3175", "cap": "flat"})
                fl = ln.makeelement(qn("a:solidFill"), {})
                cl = fl.makeelement(qn("a:srgbClr"), {"val": "BFBFBF"})
                fl.append(cl); ln.append(fl); tcPr.append(ln)


def build():
    agg, src = load_agg()
    ROWS, BULLETS = _rows(agg)
    print(f"  數據來源：{src}")
    prs = Presentation()
    prs.slide_width = Inches(10.83); prs.slide_height = Inches(7.5)
    W = 10.83
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tabs = ["2025年度投資計劃執行情況概述", "過往年度…", "主要發現", "其他信息", "六項KPI分析", "附件"]
    bc = slide.shapes.add_textbox(Inches(0.3), Inches(0.12), Inches(W - 1.5), Inches(0.2))
    bp = bc.text_frame.paragraphs[0]
    for i, t in enumerate(tabs):
        if i:
            sep = bp.add_run(); sep.text = "   |   "
            sep.font.size = Pt(6); sep.font.color.rgb = RGBColor(0xC8, 0xC8, 0xC8); sep.font.name = FONT
        r = bp.add_run(); r.text = t
        r.font.size = Pt(6); r.font.name = FONT
        r.font.bold = (i == 0); r.font.color.rgb = NAVY if i == 0 else GREY
    mg = slide.shapes.add_textbox(Inches(W - 1.0), Inches(0.12), Inches(0.8), Inches(0.2))
    mr = mg.text_frame.paragraphs[0].add_run(); mr.text = "MGM  ◀ ⌂ ▶"
    mr.font.size = Pt(6); mr.font.color.rgb = GREY; mr.font.name = FONT

    st = slide.shapes.add_textbox(Inches(0.3), Inches(0.36), Inches(W - 0.6), Inches(0.24))
    sr = st.text_frame.paragraphs[0].add_run()
    sr.text = "2025年度投資計劃執行情況概述  |  2025年度投資項目的整體執行概況"
    sr.font.size = Pt(10); sr.font.bold = True; sr.font.color.rgb = NAVY; sr.font.name = FONT

    hl = slide.shapes.add_textbox(Inches(0.3), Inches(0.64), Inches(W - 0.6), Inches(0.8))
    hr = hl.text_frame.paragraphs[0].add_run()
    hl.text_frame.word_wrap = True
    hr.text = ("MGM的2025年度計劃投資項目涵蓋博彩及非博彩範疇。下表列示各範疇的獲批計劃投資金額、報告投資金額、"
               "投資計劃完成率，以及考慮潛在調整後的投資金額及完成率。（示範版式；數字視乎所讀數據）")
    hr.font.size = Pt(7.5); hr.font.color.rgb = NAVY; hr.font.name = FONT

    tt = slide.shapes.add_textbox(Inches(0.3), Inches(1.55), Inches(6.4), Inches(0.2))
    ttr = tt.text_frame.paragraphs[0].add_run(); ttr.text = "MGM 2025年度計劃的整體投資執行概況"
    ttr.font.size = Pt(7); ttr.font.bold = True; ttr.font.color.rgb = NAVY; ttr.font.name = FONT

    ncol = 9
    gt = slide.shapes.add_table(2 + len(ROWS), ncol, Inches(0.3), Inches(1.78), Inches(6.45), Inches(4.9)).table
    gt.first_row = False; gt.horz_banding = False
    widths = [1.35, 0.5, 0.72, 0.7, 0.62, 0.72, 0.62, 0.6, 0.6]
    tot_w = sum(widths)
    for i, wd in enumerate(widths):
        gt.columns[i].width = Inches(wd * 6.45 / tot_w)
    super_hdr = ["", "", "", "報告投資金額", "", "潛在調整後投資金額", "", "", ""]
    for c in range(ncol):
        _set_cell(gt.cell(0, c), super_hdr[c], size=5.5, bold=True, fill=NAVY, white=True, align=PP_ALIGN.CENTER)
    gt.cell(0, 3).merge(gt.cell(0, 4))
    gt.cell(0, 5).merge(gt.cell(0, 6))
    for c in range(ncol):
        _set_cell(gt.cell(1, c), SUB[c], size=5.5, bold=True, fill=NAVY, white=True,
                  align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.CENTER)
    for ri, (label, kind, vals) in enumerate(ROWS, start=2):
        fill = {"sec": SECHDR, "subtot": SUBTOT, "tot": TOT}.get(kind)
        bold = kind in ("subtot", "tot", "sec")
        _set_cell(gt.cell(ri, 0), label, size=5.5, bold=bold, fill=fill, align=PP_ALIGN.LEFT)
        for c in range(1, ncol):
            _set_cell(gt.cell(ri, c), (vals[c - 1] if vals else ""), size=5.5, bold=bold, fill=fill)
    _thin_borders(gt)

    ny = 1.78 + 4.9 + 0.03
    nb2 = slide.shapes.add_textbox(Inches(0.3), Inches(ny), Inches(6.45), Inches(0.4))
    b = nb2.text_frame.paragraphs[0].add_run()
    b.text = "註：投資計劃完成率 ＝ 報告投資金額 ／ 獲批的計劃投資金額；潛在調整後完成率 ＝ 潛在調整後投資金額 ／ 獲批的計劃投資金額。"
    b.font.size = Pt(5); b.font.italic = True; b.font.color.rgb = GREY; b.font.name = FONT

    nb = slide.shapes.add_textbox(Inches(7.0), Inches(1.78), Inches(3.5), Inches(4.9))
    ntf = nb.text_frame; ntf.word_wrap = True
    for j, (lead, body) in enumerate(BULLETS):
        p = ntf.paragraphs[0] if j == 0 else ntf.add_paragraph()
        p.space_after = Pt(4)
        rb = p.add_run(); rb.text = "■ "
        rb.font.size = Pt(7.5); rb.font.color.rgb = NAVY; rb.font.name = FONT
        rl = p.add_run(); rl.text = lead
        rl.font.size = Pt(7.5); rl.font.bold = True; rl.font.color.rgb = NAVY; rl.font.name = FONT
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(7.5); rt.font.color.rgb = RGBColor(0x33, 0x33, 0x33); rt.font.name = FONT

    ft = slide.shapes.add_textbox(Inches(0.3), Inches(7.15), Inches(7), Inches(0.25))
    fr = ft.text_frame.paragraphs[0].add_run()
    fr.text = "KPMG　© 2026畢馬威會計師事務所 — 澳門特別行政區合夥制事務所。版權所有，不得轉載。"
    fr.font.size = Pt(6); fr.font.color.rgb = GREY; fr.font.name = FONT
    pg = slide.shapes.add_textbox(Inches(W - 1.1), Inches(7.15), Inches(0.9), Inches(0.25))
    pr = pg.text_frame.paragraphs[0].add_run(); pr.text = "初稿　11"
    pr.font.size = Pt(7); pr.font.bold = True; pr.font.color.rgb = NAVY; pr.font.name = FONT

    prs.save("demo_overview.pptx")
    print("✓ demo_overview.pptx（PowerPoint 開睇）")


if __name__ == "__main__":
    build()
