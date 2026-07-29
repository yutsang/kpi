#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_report.py — 一鍵：由底層數據(feed) 做齊報告數字表 → 一份 pptx（報告只係 ref，全部 data 生成）。

用法（簡單）：
    python scripts\\report\\make_report.py            # 預設 mgm
    python scripts\\report\\make_report.py wynn       # 換 entity

自動揾：feed = tableau_combined_25.csv（root）、清單 = data\\投資項目清單\\*{ENTITY}*、
       template = data\\reports\\*{ENTITY}*.pptx（跟 slide 尺寸）。
出 {entity}_報告數字表.pptx：單個項目審查匯總(25/24/23) + 金額匯總 + 設施vs活動×3，全部 native + 3色。
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("✗ pip install pandas python-pptx openpyxl"); sys.exit(1)

# 報告配色：KPMG 品牌藍 #00338D 表頭白字（報告通篇 heading 用呢個）、section 淺藍、小計灰、總計稍深
HDR = RGBColor(0x00, 0x33, 0x8D)
SEC = RGBColor(0xD9, 0xE1, 0xF2)
SUB = RGBColor(0xE7, 0xE6, 0xE6)
TOT = RGBColor(0xBD, 0xD7, 0xEE)
DARK = RGBColor(0x17, 0x17, 0x1C)      # 封面 / 章節分隔深底（近黑帶少少 navy）
LIGHT = RGBColor(0xFF, 0xFF, 0xFF)
CYAN = RGBColor(0x00, 0xB0, 0xD8)      # KPMG 輔助青（分隔頁大數字/線）

# entity → 承批公司全名（封面用；公開上市公司法定名，非機密）
ENTITY_FULL = {
    "mgm": "美高梅金殿超濠股份有限公司",
    "galaxy": "銀河娛樂場股份有限公司",
    "sjm": "澳門博彩股份有限公司",
    "wynn": "永利渡假村（澳門）股份有限公司",
    "vml": "威尼斯人澳門股份有限公司",
    "melco": "新濠博亞博彩（澳門）股份有限公司",
}

# 公司 KPMG template（空 pptx，得 master + layouts）→ 開佢做 base，formatting 全部嚟自 master。
# 放 repo root / data/ 任一（gitignored，confidential）。搵唔到就 fallback 手砌（fresh Presentation）。
TEMPLATE_NAMES = ["template.pptx", "report_template.pptx", "kpmg_template.pptx"]
TEMPLATE_DIRS = [".", "data", "data/template", "conf/local"]
# 報告版 → template layout 名（跨 master 搵第一個 match；搵唔到就 fallback）
LAYOUTS = {
    "cover": ["TITLE SLIDE - Right vertical dark image", "TITLE SLIDE", "Title Slide"],
    "section": ["Section Divider"],
    "subsection": ["Subsection Divider"],
    "appendix": ["Appendix Divider"],
    "table_text": ["1_ANALYSIS narrow table", "ANALYSIS narrow table"],
    "two_col": ["Two Column Text", "ANALYSIS 2 col text"],
    "one_col": ["One Column Text", "RPOE One Column Text"],
    "title_only": ["Title Only"],
    "findings2": ["KEY FINDINGS 2_text only"],
    "toc": ["Table of content Storage Layout", "Table of content"],
}


def _find_template():
    for d in TEMPLATE_DIRS:
        for n in TEMPLATE_NAMES:
            p = Path(d) / n
            if p.exists():
                return p
    return None


def _layout(prs, key):
    """跨所有 master 搵第一個名 match LAYOUTS[key] 嘅 slide layout；搵唔到回 None。"""
    wanted = LAYOUTS.get(key, [key])
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if lay.name in wanted:
                return lay
    return None


def _strip_slides(prs):
    """正確刪走 template 原有 content slides（drop rel + sldId）→ 避免 duplicate part 名 corruption。"""
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        rId = sldId.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        sldIdLst.remove(sldId)
        try:
            prs.part.drop_rel(rId)
        except Exception:
            pass


def _renumber_slides(prs, base=9000):
    """save 前把生成嘅 slide part 重編做高號（9000+）→ 就算 template 有殘留 orphan part（清唔切）
    都唔會撞名 → 徹底避開 duplicate-name corruption / 要 repair。"""
    try:
        from pptx.opc.packuri import PackURI
    except Exception:
        return
    for i, slide in enumerate(prs.slides, base):
        try:
            slide.part.partname = PackURI(f"/ppt/slides/slide{i}.xml")
        except Exception:
            pass


def _blank_layout(prs):
    """乾淨 layout（手砌 data 版用）：template 揀 'Title Only'/'Blank'（有 master title bar/footer），否則 index。"""
    for master in prs.slide_masters:
        for lay in master.slide_layouts:
            if lay.name in ("Blank", "Title Only"):
                return lay
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]


def _ph(slide, idx):
    """由 placeholder idx 攞 placeholder；冇就 None。"""
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


import build_project_review_table as B
import build_summary_tables as S
import build_overview_tables as O
import build_narrative as N
import render_review_table_pptx as R
from build_llm_narrative import generate_llm_narrative, Workbench   # bundler 會 inline

FEED = "tableau_combined_25.csv"


def _is_num(v):
    try:
        float(v); return True
    except (TypeError, ValueError):
        return False


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


SECTIONS = ["2025年度投資計劃執行情況概述", "過往年度投資計劃在2025年繼續執行的審查跟進",
            "本年度審查工作的主要發現", "其他信息", "投資計劃執行報告的六項KPI分析", "附件"]


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
        r.font.size = Pt(6.5); r.font.name = "微软雅黑"
        r.font.bold = (i == section_idx)
        r.font.color.rgb = HDR if i == section_idx else grey
        x += w + 0.1
    ft = slide.shapes.add_textbox(Inches(0.5), Inches(slide_h - 0.34), Inches(slide_w - 1.6), Inches(0.28))
    fr = ft.text_frame.paragraphs[0].add_run()
    fr.text = "© 2026畢馬威會計師事務所 — 澳門特別行政區合夥制事務所。版權所有，不得轉載。"
    fr.font.size = Pt(6); fr.font.color.rgb = grey; fr.font.name = "微软雅黑"
    pg = slide.shapes.add_textbox(Inches(slide_w - 1.05), Inches(slide_h - 0.34), Inches(0.9), Inches(0.28))
    pr = pg.text_frame.paragraphs[0].add_run()
    pr.text = f"初稿　{len(prs.slides._sldIdLst)}"
    pr.font.size = Pt(8); pr.font.bold = True; pr.font.color.rgb = HDR; pr.font.name = "微软雅黑"


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


def render_cover(prs, entity, date="2026年6月30日"):
    """封面（報告 p1）。template：用 TITLE SLIDE layout 填 placeholder（顏色/圖嚟自 master）；否則手砌深底。"""
    full = ENTITY_FULL.get(entity, entity.upper())
    lay = _layout(prs, "cover")
    if lay is not None:
        slide = prs.slides.add_slide(lay)
        t = _ph(slide, 0)
        if t is not None:
            t.text = f"{full}\n2025年年度投資計劃執行情況審查\n專項工作報告"
        b = _ph(slide, 11)
        if b is not None:
            b.text = f"初稿\n畢馬威會計師事務所\n{date}"
        return
    slide, w, h = _dark_slide(prs)
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.9), Inches(7.2), Inches(2.6))
    tf = tb.text_frame; tf.word_wrap = True
    for i, line in enumerate([full, "2025年年度投資計劃執行情況審查", "專項工作報告"]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = line
        r.font.size = Pt(25); r.font.bold = True; r.font.color.rgb = LIGHT
        r.font.name = "微软雅黑"
    db = slide.shapes.add_textbox(Inches(0.6), Inches(4.7), Inches(4), Inches(0.5))
    dr = db.text_frame.paragraphs[0].add_run(); dr.text = "初稿"
    dr.font.size = Pt(16); dr.font.bold = True; dr.font.color.rgb = LIGHT; dr.font.name = "微软雅黑"
    fb = slide.shapes.add_textbox(Inches(0.6), Inches(5.5), Inches(5), Inches(0.8))
    ftf = fb.text_frame
    for i, line in enumerate(["畢馬威會計師事務所", date]):
        p = ftf.paragraphs[0] if i == 0 else ftf.add_paragraph()
        r = p.add_run(); r.text = line
        r.font.size = Pt(11); r.font.color.rgb = LIGHT; r.font.name = "微软雅黑"


def divider(prs, title, number="", subitems=None):
    """章節分隔。template：用 Section/Appendix Divider layout（深底嚟自 master）+ 加標題 textbox；否則手砌。"""
    lay = _layout(prs, "appendix" if number == "6" else "section")
    if lay is not None:
        slide = prs.slides.add_slide(lay)
        w = prs.slide_width / 914400.0; h = prs.slide_height / 914400.0
        ny = h * 0.34; tx = 0.9
        if number:
            nb = slide.shapes.add_textbox(Inches(0.9), Inches(ny - 0.2), Inches(1.5), Inches(1.2))
            nr = nb.text_frame.paragraphs[0].add_run(); nr.text = f"{number}."
            nr.font.size = Pt(40); nr.font.bold = True; nr.font.color.rgb = LIGHT; nr.font.name = "Arial"
            tx = 2.2
        tb = slide.shapes.add_textbox(Inches(tx), Inches(ny), Inches(w - tx - 0.9), Inches(1.2))
        tr = tb.text_frame.paragraphs[0].add_run(); tr.text = title
        tr.font.size = Pt(24); tr.font.bold = True; tr.font.color.rgb = LIGHT; tr.font.name = "微软雅黑"
        if subitems:
            y = ny + 1.5
            for it in subitems:
                label, _pg = (it if isinstance(it, (tuple, list)) else (it, ""))
                rb = slide.shapes.add_textbox(Inches(tx), Inches(y), Inches(w - tx - 1.2), Inches(0.3))
                rr = rb.text_frame.paragraphs[0].add_run(); rr.text = label
                rr.font.size = Pt(12); rr.font.color.rgb = LIGHT; rr.font.name = "微软雅黑"
                y += 0.4
        return
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
    tr.font.size = Pt(26); tr.font.bold = True; tr.font.color.rgb = LIGHT; tr.font.name = "微软雅黑"
    if subitems:
        y = ny + 1.7
        for it in subitems:
            label, page = (it if isinstance(it, (tuple, list)) else (it, ""))
            rb = slide.shapes.add_textbox(Inches(tx), Inches(y), Inches(w - tx - 1.4), Inches(0.32))
            rr = rb.text_frame.paragraphs[0].add_run(); rr.text = label
            rr.font.size = Pt(12); rr.font.color.rgb = RGBColor(0xC8, 0xC8, 0xD0); rr.font.name = "微软雅黑"
            if page != "":
                pb = slide.shapes.add_textbox(Inches(w - 1.3), Inches(y), Inches(0.8), Inches(0.32))
                pr = pb.text_frame.paragraphs[0].add_run(); pr.text = str(page)
                pr.font.size = Pt(12); pr.font.color.rgb = RGBColor(0xC8, 0xC8, 0xD0); pr.font.name = "微软雅黑"
            y += 0.42


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
        R._set(cell, text, size=font, bold=True, align=align, color=R.WHITE, fill=HDR)
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
                R._set(t.cell(ri, ci), first if ci == 0 else "", size=font, bold=True,
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
                txt, al = R.fmt_pct(v), PP_ALIGN.RIGHT
            elif _is_num(v):
                txt, al = R.fmt_money(v), PP_ALIGN.RIGHT
            else:
                txt, al = ("" if v is None else str(v)), PP_ALIGN.LEFT
            R._set(t.cell(ri, ci), txt, size=font, bold=is_sub, align=al,
                   color=(R.RED if txt.startswith("(") else None), fill=fill)


def _bullets_into(box, bullets, size=8):
    tf = box.text_frame; tf.word_wrap = True
    for i, (head, body) in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        rm = p.add_run(); rm.text = "■ "; rm.font.size = Pt(size); rm.font.color.rgb = HDR
        if head:
            rh = p.add_run(); rh.text = head
            rh.font.bold = True; rh.font.size = Pt(size); rh.font.color.rgb = HDR
            rh.font.name = "微软雅黑"
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(size); rt.font.name = "微软雅黑"


def render_overview_page(prs, subtitle, headline, table_df, bullets):
    """報告概述式 2 欄版：頂 subtitle + navy headline，左 表，右 敘述。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    slide = prs.slides.add_slide(blank)
    _furniture(prs, slide, 0)
    st = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(slide_w - 0.8), Inches(0.24))
    sp = st.text_frame.paragraphs[0]; sr = sp.add_run(); sr.text = subtitle
    sr.font.size = Pt(9); sr.font.color.rgb = RGBColor(0x60, 0x60, 0x60); sr.font.name = "微软雅黑"
    if headline:
        hb = slide.shapes.add_textbox(Inches(0.4), Inches(0.56), Inches(slide_w - 0.8), Inches(1.0))
        htf = hb.text_frame; htf.word_wrap = True
        hr = htf.paragraphs[0].add_run(); hr.text = headline
        hr.font.bold = True; hr.font.size = Pt(11); hr.font.color.rgb = HDR; hr.font.name = "微软雅黑"
    top = 1.5 if headline else 0.6
    left_w = 5.5
    if table_df is not None and not table_df.empty:
        _draw_table(slide, table_df, 0.4, top, left_w, font=6.5)
    rx = 0.4 + left_w + 0.2
    box = slide.shapes.add_textbox(Inches(rx), Inches(top), Inches(slide_w - rx - 0.3), Inches(6.6 - top))
    _bullets_into(box, bullets, size=8)


def render_generic(prs, title, df):
    """單張表（範疇/項目 + 數字欄；·=2-row group header），自行分頁。"""
    n = len(df); ROWS = 28
    pages = [(i, min(i + ROWS, n)) for i in range(0, n, ROWS)] or [(0, 0)]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    for pi, (a, b) in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
        R._set_title(tb, f"{title}（{pi+1}/{len(pages)}）" if len(pages) > 1 else title)
        _draw_table(slide, df.iloc[a:b], 0.4, 0.72, slide_w - 0.8, font=7)


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
        rl.font.bold = True; rl.font.size = Pt(8); rl.font.name = "微软雅黑"
        rl.font.color.rgb = HDR if col is None else col
        rt = p.add_run(); rt.text = str(text)[:300]
        rt.font.size = Pt(8); rt.font.name = "微软雅黑"
        if col is not None:
            rt.font.color.rgb = col


def render_findings(prs, ent_up, df, narr):
    """③ 主要發現（slide 28-40）：每 canonical 調整類型 → 受影響項目 card
    = navy 標題條(項目+金額) + body(KPMG分析發現/管理層解釋 清單抄字)。每頁 2 個項目。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    grey = RGBColor(0x40, 0x40, 0x40)
    for adj in B.ADJ7:
        sub = d[(d["_adj"] == adj) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        projs = sub.groupby(["ng_scope", "dicj code"]).agg(名稱=("project", "first"),
                             報告=("調整前_萬", "sum"), 調整=("調整_萬", "sum")).reset_index()
        projs = projs.reindex(projs["調整"].abs().sort_values(ascending=False).index)
        recs = []
        for _, p in projs.iterrows():
            nr = N.nlook(narr, p["ng_scope"], p["dicj code"])
            recs.append((str(p["dicj code"]), str(p["名稱"]), p["報告"], p["調整"],
                         nr.get("KPMG分析發現", ""), nr.get("管理層解釋", "")))
        pages = [recs[i:i + 2] for i in range(0, len(recs), 2)]
        for pi, page in enumerate(pages):
            slide = prs.slides.add_slide(blank)
            tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
            R._set_title(tb, f"{ent_up} 本年度主要發現 — {adj}"
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
                br.text = f"{code}　{name[:32]}　│　報告 {R.fmt_money(rep)}／潛在調整 {R.fmt_money(adjv)} 萬澳門元"
                br.font.bold = True; br.font.size = Pt(9)
                br.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); br.font.name = "微软雅黑"
                # body
                body = slide.shapes.add_textbox(Inches(0.4), Inches(y + 0.32),
                                                Inches(slide_w - 0.8), Inches(2.5))
                _finding_body(body, find, mgmt, grey)
                y += 3.0


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
        nr = N.nlook(narr, p["ng_scope"], p["dicj code"])
        recs.append((str(p["dicj code"]), str(p["名稱"]), p["報告"],
                     nr.get("實施地點", ""), nr.get("實際投資內容", "")))
    pages = [recs[i:i + 2] for i in range(0, len(recs), 2)]
    for pi, page in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
        R._set_title(tb, f"{ent_up} 附件二 部分項目的現場走訪情況"
                     + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
        y = 0.8
        for code, name, amt, loc, desc in page:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(y),
                                         Inches(slide_w - 0.8), Inches(0.3))
            bar.fill.solid(); bar.fill.fore_color.rgb = HDR; bar.line.fill.background()
            btf = bar.text_frame; btf.word_wrap = True
            btf.margin_left = Emu(54000); btf.margin_top = btf.margin_bottom = Emu(9000)
            br = btf.paragraphs[0].add_run()
            br.text = f"{code}　{name[:30]}　│　設施建設（資本性支出）{R.fmt_money(amt)} 萬澳門元"
            br.font.bold = True; br.font.size = Pt(9)
            br.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); br.font.name = "微软雅黑"
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
                rl.font.bold = True; rl.font.size = Pt(8); rl.font.name = "微软雅黑"
                rl.font.color.rgb = HDR if col is None else col
                rt = p0.add_run(); rt.text = str(text)[:300]
                rt.font.size = Pt(8); rt.font.name = "微软雅黑"
                if col is not None and label != "地點":
                    rt.font.color.rgb = col
            y += 3.0


def _prose_slide(prs, title, bullets, headline=None):
    """一版敘述（navy 標題 +（可選）headline 粗體導語 + ■ bullet；bullet=(粗體引子, 內文)）。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    slide = prs.slides.add_slide(blank)
    _furniture(prs, slide, 0)
    tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.3), Inches(slide_w - 0.8), Inches(0.4))
    R._set_title(tb, title)
    box = slide.shapes.add_textbox(Inches(0.5), Inches(0.95), Inches(slide_w - 1.0), Inches(5.6))
    tf = box.text_frame; tf.word_wrap = True
    started = False
    if headline:
        p = tf.paragraphs[0]; p.space_after = Pt(10)
        r = p.add_run(); r.text = headline
        r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = HDR
        r.font.name = "微软雅黑"
        started = True
    for head, body in bullets:
        p = tf.add_paragraph() if started else tf.paragraphs[0]
        started = True
        p.space_after = Pt(8)
        rm = p.add_run(); rm.text = "■ "; rm.font.size = Pt(10); rm.font.color.rgb = HDR
        if head:
            rh = p.add_run(); rh.text = head
            rh.font.bold = True; rh.font.size = Pt(10); rh.font.color.rgb = HDR
            rh.font.name = "微软雅黑"
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(10); rt.font.name = "微软雅黑"


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


def _pct(v):
    return f"{v*100:.1f}%" if isinstance(v, (int, float)) and not pd.isna(v) else "—"


def _rate_of(df, name, col):
    r = df[df["範疇"] == name]
    if not len(r) or col not in df.columns:
        return None
    v = r.iloc[0][col]
    return v if isinstance(v, (int, float)) and not pd.isna(v) else None


def _prose_paginated(prs, title, bullets, per):
    if not bullets:
        return
    pages = [bullets[i:i + per] for i in range(0, len(bullets), per)]
    for pi, page in enumerate(pages):
        t = title + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else "")
        _prose_slide(prs, t, page)


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
        R._set_title(tb, title + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
        if subtitle:
            sb = slide.shapes.add_textbox(Inches(0.4), Inches(0.68), Inches(slide_w - 0.8), Inches(0.24))
            sr = sb.text_frame.paragraphs[0].add_run(); sr.text = subtitle
            sr.font.size = Pt(8); sr.font.italic = True
            sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70); sr.font.name = "微软雅黑"
        half = (len(page) + 1) // 2
        lb = slide.shapes.add_textbox(Inches(0.4), Inches(top), Inches(colw), Inches(6.0 - top + 0.85))
        _bullets_into(lb, page[:half], size=8)
        if page[half:]:
            rb = slide.shapes.add_textbox(Inches(0.4 + colw + 0.2), Inches(top), Inches(colw), Inches(6.0 - top + 0.85))
            _bullets_into(rb, page[half:], size=8)


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
            nr = N.nlook(narr, scope, pp["dicj code"])
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


def _adj_detail_bullets(ent_up, adj, df, narr, llm=None):
    """slide 16-17 潛在調整事項詳述 → 回 bullets：LLM summary 優先，否則清單分析發現。"""
    if not narr:
        return []
    llm_adj = (llm or {}).get("adj", {})
    pb = S.BUCKET_ORDER[0]      # 2025計劃 bucket（唔用合計；期後另計）
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
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
            nr = N.nlook(narr, pp["ng_scope"], pp["dicj code"])
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


def _adj_summary(ent_up, adj):
    """潛在調整事項匯總（報告 slide 15）→ 回 (headline, bullets)。逐類型金額。
    ⚠ 用 2025計劃 bucket（唔係合計）：報告調整詳述只計 2025年度計劃，期後另有匯總。"""
    pb = S.BUCKET_ORDER[0]      # "2025年度投資計劃"
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
    df["_plan_year"] = df["year_bucket"].map(B._plan_year)
    plan = B.load_plan(qingdan) if qingdan else None
    cat = B.load_category(qingdan) if qingdan else None     # 項目性質(D)→派零投資項目計劃返範疇
    narr = N.load_narrative(qingdan) if qingdan else {}     # 清單 by-project narrative（抄字）
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

    tmpl = _find_template() if "--use-template" in sys.argv else None   # 預設 fresh 手砌（template 樣式已 hardcode）；--use-template 先開 template
    if tmpl:
        prs = Presentation(str(tmpl))
        _strip_slides(prs)      # drop_rel 正確清走原有 content slides（唔會 duplicate 名 corrupt）
        print(f"    template（master-driven）：{tmpl}  layouts={sum(len(m.slide_layouts) for m in prs.slide_masters)}")
    else:
        prs = Presentation()
        if template:
            prs.slide_width, prs.slide_height = Presentation(str(template)).slide_width, Presentation(str(template)).slide_height
        else:
            prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
        print("    無 template → fallback 手砌 formatting")

    sdf = S._load(feed, entity)     # 於2025發生 slice（概述 + 金額匯總 共用）

    render_cover(prs, entity)       # 封面（報告 p1）

    # ① 2025年度投資計劃執行情況概述（報告 slide 8-18）
    divider(prs, "2025年度投資計劃執行情況概述", "1", [
        ("1.1  股權架構簡圖及發生投資支出的主體公司", ""),
        ("1.2  2025年度計劃的整體投資支出概況", ""),
        ("1.3  2025年度投資項目的整體執行概況", ""),
        ("1.4  2025年度投資計劃報告投資金額的潛在調整事項匯總", ""),
    ])
    ov = O.overview_by_bucket(sdf, "2025年度投資計劃", plan, cat)
    adj = O.adjustment_bridge(sdf)
    if not ov.empty:      # slide 10-11：表左 + headline/執行敘述右（報告 2 欄式）
        hl, hlb = _headline(ent_up, ov, sdf, plan)
        exb = _exec_bullets(ent_up, ov)
        render_overview_page(prs, f"2025年度投資計劃執行情況概述 | {ent_up} 2025年度計劃的整體投資支出及執行概況",
                             hl, ov.fillna(""), hlb + exb)
        render_category_overview(prs, ent_up, ov, sdf, narr, llm)   # slide 13-14 逐範疇概況（LLM 優先）
        zit = O.zero_investment_text(O.zero_investment_summary(sdf, plan, cat, narr, ent_up), ent_up)
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
        ov = O.overview_by_bucket(sdf, bk, plan, cat)
        if not ov.empty:
            render_generic(prs, f"{ent_up} {bk}金額概覽", ov.fillna(""))

    # ③ 本年度審查工作的主要發現（報告 slide 28-40）
    divider(prs, "本年度審查工作的主要發現", "3")
    fs = O.finding_summary(sdf)
    if not fs.empty:
        render_generic(prs, f"{ent_up} 主要發現摘要", fs.fillna(""))
    if narr:      # 逐調整類型 × 項目：金額(feed) + 發現/管理層解釋(清單抄字)
        render_findings(prs, ent_up, sdf, narr)

    # ④ 其他信息（報告 slide 42-63）
    divider(prs, "其他信息", "4")
    render_generic(prs, f"{ent_up} 2025年度投資計劃及過往年度期後投資於2025年發生的投資金額匯總",
                   S.summary_amount(sdf).fillna(""))
    for bk in S.BUCKET_ORDER:
        fa = S.facility_activity(sdf, bk)
        if not fa.empty:
            render_generic(prs, f"{ent_up} {bk}區分設施建設/活動舉辦的投資金額", fa.fillna(""))
    for yr in (25, 24, 23):     # 單個項目審查匯總（slide 46-63）
        tab, _ = B.build_year(df, yr, plan.get(yr) if plan else None)
        if tab is not None and not tab.empty:
            R.render_sheet(prs, f"報告年{yr}", tab.fillna(""), list(tab.columns))

    # ⑥ 附件二 現場走訪（slide 93-100）
    if narr:
        divider(prs, "附件", "6")
        render_site_visits(prs, ent_up, sdf, narr)

    if tmpl:      # template mode：重編 slide 高號，徹底避開 template 殘留 orphan part 撞名 corruption
        _renumber_slides(prs)

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
