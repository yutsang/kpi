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

import build_project_review_table as B
import build_summary_tables as S
import build_overview_tables as O
import build_narrative as N
import render_review_table_pptx as R

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


def divider(prs, text):
    """章節分隔頁（navy 底白字），俾 deck 有報告 section 結構。"""
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide = prs.slides.add_slide(blank)
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid(); rect.fill.fore_color.rgb = HDR
    rect.line.fill.background()
    tf = rect.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.size = Pt(26); r.font.bold = True
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); r.font.name = "Microsoft JhengHei"


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
            rh.font.name = "Microsoft JhengHei"
        rt = p.add_run(); rt.text = body
        rt.font.size = Pt(size); rt.font.name = "Microsoft JhengHei"


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
    """body 文字框：KPMG分析發現 / 管理層解釋 兩段，label 加粗。"""
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
                br.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); br.font.name = "Microsoft JhengHei"
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
            sr.font.color.rgb = RGBColor(0x70, 0x70, 0x70); sr.font.name = "Microsoft JhengHei"
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
    llm = _load_llm(entity)     # {entity}_llm_narrative.json（build_llm_narrative.py 出）有就用 LLM 文字
    if llm:
        print(f"    LLM narrative: adj {len(llm.get('adj', {}))}、cat {len(llm.get('cat', {}))} 段")

    prs = Presentation()
    if template:
        ref = Presentation(str(template))
        prs.slide_width, prs.slide_height = ref.slide_width, ref.slide_height
    else:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    sdf = S._load(feed, entity)     # 於2025發生 slice（概述 + 金額匯總 共用）

    # ① 2025年度投資計劃執行情況概述（報告 slide 8-18）
    divider(prs, "一、2025年度投資計劃執行情況概述")
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
    divider(prs, "二、過往年度投資計劃在2025年繼續執行的審查跟進")
    for bk in ["2024年度計劃期後投資", "2023年度計劃期後投資"]:
        ov = O.overview_by_bucket(sdf, bk, plan, cat)
        if not ov.empty:
            render_generic(prs, f"{ent_up} {bk}金額概覽", ov.fillna(""))

    # ③ 本年度審查工作的主要發現（報告 slide 28-40）
    divider(prs, "三、本年度審查工作的主要發現")
    fs = O.finding_summary(sdf)
    if not fs.empty:
        render_generic(prs, f"{ent_up} 主要發現摘要", fs.fillna(""))
    if narr:      # 逐調整類型 × 項目：金額(feed) + 發現/管理層解釋(清單抄字)
        render_findings(prs, ent_up, sdf, narr)

    # ④ 其他信息（報告 slide 42-63）
    divider(prs, "四、其他信息")
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
        divider(prs, "六、附件")
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


if __name__ == "__main__":
    main()
