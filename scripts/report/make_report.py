#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_report.py — 一鍵：由底層數據(feed) 做齊報告數字表 → 一份 pptx（報告只係 ref，全部 data 生成）。

用法（簡單）：
    python scripts\\report\\make_report.py            # 預設 mgm
    python scripts\\report\\make_report.py wynn       # 換 entity

自動揾：feed = tableau_combined_25.csv（root）、清單 = data\\投資項目清單\\*{ENTITY}*、
       template = data\\reports\\*{ENTITY}*.pptx（跟 slide 尺寸；冇就用 10.83x7.5 報告標準）。
出 {entity}_report_llm.pptx：概述 + 期後 + 主要發現 + 金額匯總 + 設施vs活動 + 單項審查(25/24/23)
+ 附件走訪，全部 native table／文字，版式跟 layout.py（KPMG house style，對齊報告 scan）。
體檢：python scripts\\report\\inspect_pptx.py {entity}_report_llm.pptx [--preview]
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

import layout as L                     # 版式引擎（KPMG house style；顏色/字體/表格/分頁全部喺嗰度）

# 報告配色 alias（真身喺 layout.py，跟 KPMG Visual identity overview）
HDR = L.NAVY                           # KPMG Blue 00338D：表頭白字 / 標題 / 導語
SEC = L.SECFILL                        # 範疇 section 行
SUB = L.SUBTOT                         # 小計
TOT = L.TOTAL                          # 總計
DARK = L.DARK                          # 封面 / 章節分隔深底
LIGHT = L.WHITE
GREY = L.GREY

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


USE_TEMPLATE = False    # 只有 --use-template 真係開咗 KPMG template 先可以用 template layout


def _layout(prs, key):
    """跨所有 master 搵第一個名 match LAYOUTS[key] 嘅 slide layout；搵唔到回 None。
    ⚠ 冇開 template 時一定回 None —— 否則會撞正 python-pptx 預設包嗰個 'Title Slide'
    layout，封面變一版白底光板（冇深底/冇 KPMG 字標/冇日期）。"""
    if not USE_TEMPLATE:
        return None
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
import biao2 as B2
from build_llm_narrative import (generate_llm_narrative, Workbench, tbl_key,   # bundler 會 inline
                                 proj_key, bkt_key)

BUILD_STAMP = "dev"          # bundler 會換做 git sha + 日期（睇 output 就知跑緊邊版）
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


ENT_UP = "MGM"      # 由 main() 設；breadcrumb 右上角用（章節清單見 layout.SECTIONS）


def _page(prs, section_idx=0, crumb=None, headline=None):
    """開一版內容頁 = breadcrumb + footer(+頁碼) + 灰標題 + navy 導語。
    回 (slide, W, H, top_y)：top_y = 內容可以由邊開始。"""
    slide = L.blank(prs)
    W, H = L.size_of(prs)
    L.breadcrumb(slide, W, section_idx, ENT_UP)
    L.footer(slide, W, H, len(prs.slides._sldIdLst))
    top = L.page_head(slide, W, crumb, headline) if crumb else 0.5
    return slide, W, H, top


def _furniture(prs, slide, section_idx=0):
    """（保留舊 API）頂 nav tabs + 底 KPMG copyright + 初稿 + 頁碼。"""
    W, H = L.size_of(prs)
    L.breadcrumb(slide, W, section_idx, ENT_UP)
    L.footer(slide, W, H, len(prs.slides._sldIdLst))


def _dark_slide(prs):
    """新增一版深底（封面/分隔共用），回 (slide, w, h)。"""
    return L.dark_slide(prs)


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


def _fmt_ratio(v):
    """比例欄：scan 寫『(89.4%)』—— 調減（負數）用括號。"""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return "" if v is None or str(v).strip() == "" else str(v)
    return f"({abs(f)*100:.1f}%)" if f < 0 else f"{f*100:.1f}%"


def _cell_txt(c, v):
    """跟欄名格式化：率／比例→%，數字→千分位（負數括號），其餘原文。"""
    if "比例" in str(c):
        return _fmt_ratio(v)
    if "率" in str(c):
        return R.fmt_pct(v)
    if _is_num(v):
        return R.fmt_money(v)
    return "" if v is None else str(v)


def _df_table(df, first_label=None):
    """DataFrame → (subs, rows, widths, supers)：認 `大組·細名` 做兩層表頭，
    第一欄空 = 範疇 section 行，尾『小計/合計/總計』= shaded。
    第一欄表頭：範疇表用「（萬澳門元）」（跟 scan 角位放單位），其餘用返欄名。"""
    cols = list(df.columns)
    if first_label is None:
        first_label = "（萬澳門元）" if cols[0] == "範疇" else cols[0]
    grouped = any("·" in c for c in cols)
    widths = [2.0 if c in ("範疇", "項目名稱", "潛在調整事項") else
              (2.8 if c == "主要涉及項目" else 0.95) for c in cols]
    subs = [first_label] + [(c.split("·")[1] if "·" in c else c) for c in cols[1:]]
    supers = None
    if grouped:
        groups = [""] + [(c.split("·")[0] if "·" in c else "") for c in cols[1:]]
        supers, ci = [], 1
        while ci < len(cols):
            cj = ci
            while cj + 1 < len(cols) and groups[cj + 1] == groups[ci]:
                cj += 1
            supers.append((groups[ci], ci, cj + 1)); ci = cj + 1
        supers.insert(0, ("", 0, 1))
    rows = []
    for _, row in df.iterrows():
        first = str(row[cols[0]]).strip()
        if all(str(row[c]).strip() == "" for c in cols[1:]):
            rows.append(("sec", [first] + [""] * (len(cols) - 1))); continue
        kind = ("tot" if first.endswith("總計") else
                "subtot" if first.endswith(("小計", "合計")) else "data")
        rows.append((kind, [first] + [_cell_txt(c, row[c]) for c in cols[1:]]))
    return subs, rows, widths, supers


def _draw_table(slide, df, x, y, max_w, font=6.5):
    """喺 slide (x,y) 畫 navy 表（單 chunk，caller 自行分頁）。max_w=可用闊(吋)。"""
    subs, rows, widths, supers = _df_table(df)
    return L.draw_table(slide, x, y, max_w, subs, rows, widths, supers=supers,
                        font=font, hfont=font - 0.5)


def _bullets_into(box, bullets, size=8):
    """（保留舊 API）scan 敘述格式：navy 粗體小標題 + body 段落。"""
    L.prose(box, bullets, head_size=size - 1, body_size=size - 1.5)


def render_overview_page(prs, crumb, headline, table_df, bullets, *, sec=0, table_name=None,
                         note=None):
    """報告概述式 2 欄版（對 scan slide 10/15）：crumb + navy 導語，左 表，右 敘述。"""
    slide, W, H, top = _page(prs, sec, crumb, headline)
    left_w = W * 0.60
    tbl_bot = top
    if table_df is not None and not table_df.empty:
        if table_name:
            top = L.caption_bar(slide, L.MARGIN, top, left_w, table_name)
        subs, rows, widths, supers = _df_table(table_df)
        wid = [w * left_w / sum(widths) for w in widths]
        avail = L.CONTENT_BOTTOM - top - 0.30          # 留位俾表下面個「註」
        hh = L.header_h(supers, subs, wid, 5.5)
        font = 6.0
        while font > 4.2 and sum(L.row_h(c, wid, font) for _, c in rows) > avail - hh:
            font -= 0.25
        tbl_bot, _ = L.draw_table(slide, L.MARGIN, top, left_w, subs, rows, widths,
                                  supers=supers, font=font, hfont=max(4.5, font - 0.5),
                                  fill_h=avail)
    if note:      # 「註」貼喺表底下，唔可以同底部嘅資料來源疊字
        L.put(slide, L.MARGIN, min(tbl_bot + 0.06, L.CONTENT_BOTTOM - 0.30), left_w, 0.3,
              note, size=5, italic=True, color=L.GREY)
    rx = L.MARGIN + left_w + 0.22
    L.prose_box(slide, rx, top - 0.02, W - rx - L.MARGIN, L.CONTENT_BOTTOM - top, bullets)
    L.source_note(slide, W)


def _total_line(df):
    """由表自己嘅『總計』行砌一句機械導語（避免「淨得個表冇文字」）。"""
    cols = list(df.columns)
    tot = df[df[cols[0]].astype(str).str.strip().str.endswith("總計")]
    if tot.empty:
        return ""
    r = tot.iloc[0]
    parts = []
    for c in cols[1:]:
        v = r[c]
        if not _is_num(v) or float(v) == 0:
            continue
        lab = c.replace("·", "－")
        parts.append(f"{lab} {_cell_txt(c, v)}")
    return "；".join(parts[:6]) + "（單位：萬澳門元，除完成率外）。" if parts else ""


def _table_bullets(df):
    """冇 LLM 時嘅機械表旁敘述（由表自己嘅小計/總計行計，自洽）→ [(小標題, 內容)]。"""
    cols = list(df.columns)
    first = df[cols[0]].astype(str).str.strip()
    num = [c for c in cols[1:] if pd.to_numeric(df[c], errors="coerce").notna().any()]
    if not num:
        return []
    out = []
    tot = df[first.str.endswith("總計")]
    if not tot.empty:
        out.append(("整體情況", "全部範疇合計 " +
                    "；".join(f"{c.replace('·', '－')} {_cell_txt(c, tot.iloc[0][c])}"
                              for c in num[:5]) + "（單位：萬澳門元，除完成率外）。"))
    subs = df[first.str.endswith("小計")]
    if not subs.empty:
        out.append(("博彩／非博彩分佈", "；".join(
            f"{r[cols[0]]} " + "、".join(f"{c.replace('·', '－')} {_cell_txt(c, r[c])}"
                                         for c in num[:3]) for _, r in subs.iterrows()) + "。"))
    val = next((c for c in num if "報告" in c or "合計" in c), num[0])
    data = df[~first.str.endswith(("小計", "總計")) &
              df[cols[1:]].astype(str).apply(lambda r: any(x.strip() for x in r), axis=1)].copy()
    data["_v"] = pd.to_numeric(data[val], errors="coerce")
    top = data.sort_values("_v", ascending=False).head(4)
    if not top.empty:
        out.append((f"金額最大的範疇（按{val.replace('·', '－')}）", "、".join(
            f"{r[cols[0]]}（{_cell_txt(val, r[val])}）" for _, r in top.iterrows()) + "。"))
    return out


def _bucket_adj_table(ov):
    """期後調整事項匯總嘅左表（對 scan p-11）：範疇 × 報告(a)｜潛在調整(b)｜調整後(c=a+b)｜b/a。"""
    keep = ["範疇", "項目數量", "報告投資金額", "潛在調整金額", "潛在調整後投資金額"]
    d = ov[[c for c in keep if c in ov.columns]].copy()
    rep = pd.to_numeric(d.get("報告投資金額"), errors="coerce")
    adj = pd.to_numeric(d.get("潛在調整金額"), errors="coerce")
    ratio = (adj / rep).where(rep.abs() > 0).astype(object)   # object：section 行要填 ""（避 pandas FutureWarning）
    ratio[d["範疇"].astype(str).str.strip().eq("") | rep.isna()] = ""
    d["潛在調整金額佔報告投資金額比例"] = ratio
    return d


def render_bucket_adjustment(prs, ent_up, bk, sdf, ov, narr, llm=None):
    """② 期後【調整事項匯總】（scan p-11 / p-13）：左 = 範疇 × 調整表，
    右 = navy 小標題 + 編號清單『N. 類型（金額）：說明』，編號跟七大類 canonical 序（會跳號）。"""
    yr = bk[:4]
    d = sdf[sdf["_bucket"] == bk].copy()
    if d.empty:
        return
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
    llm_bkt = (llm or {}).get("bkt", {})
    items = []
    for i, t in enumerate(B.ADJ7, start=1):      # 編號＝七大類 canonical 序
        sub = d[d["_adj"] == t]
        amt = pd.to_numeric(sub["調整_萬"], errors="coerce").sum()
        if abs(amt) < 0.5:
            continue
        body = llm_bkt.get(bkt_key(bk, t)) or (
            f"在{yr}年度投資計劃期後投資金額中，{ent_up}同樣申報了{t}。")
        body += (f"與前述的「2025年度投資計劃報告投資金額的潛在調整事項匯總」一致，"
                 f"我們就{yr}年度投資計劃期後投資金額同樣建議了該項調整。詳見後續主要發現「{t}」。")
        items.append((i, f"{t}（{_amt(amt)}）：", body))
    if not items:
        return
    tot = pd.to_numeric(d["調整_萬"], errors="coerce").sum()
    npj = int(d[pd.to_numeric(d["調整_萬"], errors="coerce") != 0]["dicj code"].nunique())
    aft = pd.to_numeric(d["調整後_萬"], errors="coerce").sum()
    head = (f"基於我們的各項審查程序，我們認為{ent_up}報告的{yr}年度投資計劃在2025年繼續執行的支出中，"
            f"存在{_cn(len(items))}大類的調整事項，潛在調減投資金額約{_amt(tot)}"
            f"（涉及{npj}個投資項目），經潛在調減後的{yr}年度計劃投資項目在2025年的投資金額約{_amt(aft)}。")
    S2 = "過往年度投資計劃在2025年繼續執行的審查跟進"
    tname = f"{ent_up} {yr}年度投資計劃於2025年申報的期後投資金額的潛在調整"
    W, H = L.size_of(prs)
    lw = W * 0.55
    tbl = _bucket_adj_table(ov)
    subs, rows, widths, supers = _df_table(tbl)
    wid = [w * lw / sum(widths) for w in widths]
    cw = W - (L.MARGIN + lw + 0.22) - L.MARGIN
    pages = L.fit_prose([(h, b) for _n, h, b in items], cw, L.CONTENT_BOTTOM - 1.9)
    idx = 0
    for pi, page in enumerate(pages):
        suffix = _pg(pi + 1, len(pages))
        slide, W, H, top = _page(prs, 1, f"{S2}  |  {yr}年度投資計劃報告投資金額的潛在調整事項匯總",
                                 head + suffix)
        if pi == 0:
            t2 = L.caption_bar(slide, L.MARGIN, top, lw, tname)
            L.draw_table(slide, L.MARGIN, t2, lw, subs, rows, widths, supers=supers,
                         font=6, hfont=5.5, fill_h=L.CONTENT_BOTTOM - t2 - 0.28)
            L.put(slide, L.MARGIN, L.CONTENT_BOTTOM - 0.26, lw, 0.3,
                  "註：金額單位為萬澳門元；括號表示調減。", size=5, italic=True, color=L.GREY)
        box = L._tb(slide, L.MARGIN + lw + 0.22, top - 0.02, cw, L.CONTENT_BOTTOM - top)
        L.prose_numbered(box, items[idx:idx + len(page)], size=6.5,
                         title=(tname if pi == 0 else tname + "（續）"))
        idx += len(page)
        L.source_note(slide, W, more=(pi < len(pages) - 1))


def render_generic(prs, title, df, *, sec=3, crumb=None, headline=None, note=None,
                   llm=None, tbl_id=None, side=None):
    """單張表（範疇/項目 + 數字欄；·=2-row group header）。逐頁：crumb + 導語 + caption bar
    + 表 + 資料來源，按【累積高度】分頁（唔會超出版面）。

    side=True → 報告式 2 欄（表左 + 敘述右，對 scan p-10~p-13）。敘述優先用 LLM 寫嘅
    `llm['tbl'][tbl_id]`，冇就用 _table_bullets 機械生成。欄數少（≤8）而且一版放得落先會用。"""
    if df is None or df.empty:
        return
    raw = (llm or {}).get("tbl", {}).get(tbl_id) or []
    if isinstance(raw, dict):                      # 新 shape：{"導語": …, "段落": [[h,b],…]}
        headline = raw.get("導語") or headline
        raw = raw.get("段落") or []
    bullets = [tuple(x) for x in raw] or _table_bullets(df)
    if side is None:
        side = len(df.columns) <= 8 and bool(bullets)
    if side and bullets:
        W, _H = L.size_of(prs)
        lw = W * 0.60
        s2, r2, w2, sp2 = _df_table(df)
        wid2 = [w * lw / sum(w2) for w in w2]
        need = L.header_h(sp2, s2, wid2, 5.0) + sum(L.row_h(c, wid2, 5.0) for _, c in r2)
        if need <= L.CONTENT_BOTTOM - 1.9:              # 一版放得落先用 2 欄，否則落返全闊分頁
            render_overview_page(prs, (crumb or title), headline or _total_line(df), df,
                                 bullets, sec=sec, table_name=title, note=note)
            return
    subs, rows, widths, supers = _df_table(df)
    W, H = L.size_of(prs)
    tw = W - 2 * L.MARGIN
    wid = [w * tw / sum(widths) for w in widths]
    head = headline or _total_line(df)
    crumb = crumb or title
    # 先用一版試高度（導語行數會食掉可用高）
    probe_top = L.HEAD_Y + L.head_h(head, W)[0] + 0.10
    avail = L.CONTENT_BOTTOM - probe_top - 0.24
    hh = L.header_h(supers, subs, wid, 5.5)
    pages = L.fit_rows(rows, wid, 6.5, avail, hh)
    for pi, chunk in enumerate(pages):
        suffix = f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""
        slide, W, H, top = _page(prs, sec, crumb, (head or "") + suffix)
        top = L.caption_bar(slide, L.MARGIN, top, tw, title + suffix)
        L.draw_table(slide, L.MARGIN, top, tw, subs, chunk, widths, supers=supers,
                     font=6.5, hfont=6, fill_h=L.CONTENT_BOTTOM - top - 0.28)
        L.source_note(slide, W, note=note, more=(pi < len(pages) - 1))


def _cards(prs, sec, crumb, headline, recs, *, note=None):
    """逐個項目一張 card（navy 標題條 + 敘述段），按【累積高度】排版分頁 → 填滿版面唔留大白位。
    recs = [(bar_text, [(label, body)])]。"""
    W, H = L.size_of(prs)
    cw = W - 2 * L.MARGIN
    probe = L.HEAD_Y + L.head_h(headline, W)[0] + 0.10
    avail = L.CONTENT_BOTTOM - probe

    def card_h(items):
        return 0.24 + L.est_prose_h(items, cw - 0.12, head_size=7.5, body_size=7.5, gap=3) + 0.14
    pages, cur, used = [], [], 0.0
    for rec in recs:
        h = card_h(rec[1])
        if cur and used + h > avail:
            pages.append(cur); cur, used = [], 0.0
        cur.append(rec); used += min(h, avail)
    if cur:
        pages.append(cur)
    for pi, page in enumerate(pages):
        suffix = f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""
        slide, W, H, y = _page(prs, sec, crumb, (headline or "") + suffix)
        for bar_text, items in page:
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(L.MARGIN), Inches(y),
                                         Inches(cw), Inches(0.22))
            bar.fill.solid(); bar.fill.fore_color.rgb = HDR
            bar.line.fill.background(); bar.shadow.inherit = False
            btf = bar.text_frame; btf.word_wrap = True
            btf.margin_left = Emu(54000); btf.margin_top = btf.margin_bottom = Emu(0)
            btf.vertical_anchor = MSO_ANCHOR.MIDDLE
            br = btf.paragraphs[0].add_run(); br.text = bar_text
            br.font.bold = True; br.font.size = Pt(8)
            br.font.color.rgb = LIGHT; br.font.name = "微软雅黑"
            bh = min(L.est_prose_h(items, cw - 0.12, head_size=7.5, body_size=7.5, gap=3),
                     L.CONTENT_BOTTOM - y - 0.26)
            L.prose_box(slide, L.MARGIN + 0.06, y + 0.26, cw - 0.12, bh, items,
                        head_size=7.5, body_size=7.5, gap=3)
            y += 0.24 + bh + 0.14
        L.source_note(slide, W, note=note, more=(pi < len(pages) - 1))


def _finding_body(box, find, mgmt, grey=None):
    """（保留舊 API）KPMG分析發現 / 管理層解釋 兩段。"""
    L.prose(box, [(l + "：", t) for l, t in
                  [("KPMG分析發現", find), ("管理層解釋", mgmt)] if t],
            head_size=7.5, body_size=7.5, gap=3)


def render_findings(prs, ent_up, df, narr, llm=None, b2=None):
    """③ 主要發現：每 canonical 調整類型 → 受影響項目 card = navy 標題條(項目+金額) + body。
    body 優先用 LLM 寫嘅『事項描述』（ground 住表2＋清單），管理層原話照樣保留；
    冇 LLM 就 fallback 清單抄字（KPMG分析發現／管理層解釋）。"""
    llm_proj = (llm or {}).get("proj", {})
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
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
            desc = llm_proj.get(proj_key(adj, p["ng_scope"], p["dicj code"]), "")
            if desc:      # LLM 寫嘅事項描述（ground 表2＋清單）＋ 管理層原話
                items = [("事項描述：", desc)]
                if nr.get("管理層解釋"):
                    items.append(("管理層解釋：", nr["管理層解釋"]))
            else:
                items = [(l + "：", t) for l, t in
                         [("KPMG分析發現", nr.get("KPMG分析發現", "")),
                          ("管理層解釋", nr.get("管理層解釋", "")),
                          ("跨司工作組／KPMG意見", nr.get("跨司回覆", "") or nr.get("KPMG回覆", ""))] if t]
                if not items and b2:      # 清單冇 → 用表2 抽到嘅原文頂住
                    t2 = B2.b2text(b2, p["ng_scope"], p["dicj code"])
                    if t2:
                        items = [("事項描述：", t2[:600])]
            if not items:
                items = [("", "清單未提供本項目之分析發現，待項目組補充。")]
            recs.append((f"{p['dicj code']}　{str(p['名稱'])[:34]}　│　報告 "
                         f"{R.fmt_money(p['報告'])}／潛在調整 {R.fmt_money(p['調整'])} 萬澳門元", items))
        tot = sub["調整_萬"].sum()
        head = (f"{ent_up} 報告的投資金額中，屬「{adj}」之潛在調減金額合計約{abs(tot):,.0f}萬澳門元，"
                f"涉及{len(recs)}個投資項目，逐項說明如下：")
        _cards(prs, 2, f"本年度審查工作的主要發現  |  {adj}", head, recs)


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
    recs = []
    for _, p in g.iterrows():
        nr = N.nlook(narr, p["ng_scope"], p["dicj code"])
        items = [(l + "：", t) for l, t in
                 [("實施地點", nr.get("實施地點", "")),
                  ("現場走訪概述", nr.get("實際投資內容", "")),
                  ("現場走訪圖片", "〔待插入〕")] if t]
        recs.append((f"{p['dicj code']}　{str(p['名稱'])[:32]}　│　設施建設（資本性支出）"
                     f"{R.fmt_money(p['報告'])} 萬澳門元", items))
    head = (f"我們就{ent_up}報告投資金額中設施建設（資本性支出）達{threshold:,.0f}萬澳門元或以上之"
            f"{len(recs)}個投資項目進行了現場走訪，走訪情況如下：")
    _cards(prs, 5, "附件  |  部分項目的現場走訪情況", head, recs,
           note="資料來源：現場走訪記錄、管理層提供之項目資料，畢馬威分析")


def _prose_slide(prs, title, bullets, headline=None, *, sec=0):
    """一版敘述（crumb + navy 導語 + 段落），按估算高度自動分頁。"""
    W, H = L.size_of(prs)
    cw = W - 2 * L.MARGIN
    probe = L.HEAD_Y + L.head_h(headline, W)[0] + 0.10
    pages = L.fit_prose(bullets, cw, L.CONTENT_BOTTOM - probe, head_size=8, body_size=8)
    for pi, page in enumerate(pages):
        suffix = f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""
        slide, W, H, top = _page(prs, sec, title, (headline or "") + suffix)
        L.prose_box(slide, L.MARGIN, top, cw, L.CONTENT_BOTTOM - top, page,
                    head_size=8, body_size=8, gap=7)
        L.source_note(slide, W, more=(pi < len(pages) - 1))


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

    d = df.copy()
    d["_adj"] = pd.to_numeric(d["調整_萬"], errors="coerce").fillna(0)
    codes = d["dicj code"].astype(str)
    n_impl = d[pd.to_numeric(d["調整前_萬"], errors="coerce").fillna(0) != 0]["dicj code"].nunique()
    n_plan = len(plan.get(25, {})) if plan else n_impl
    n_zero = max(n_plan - n_impl, 0)
    n_adj = d[d["_adj"] != 0]["dicj code"].nunique()
    headline = (f"{ent_up} 2025年度原獲批計劃開展{n_plan}個投資項目，涉及計劃投資金額約{_amt(plan_amt)}；"
                f"{ent_up}提交的投資執行報告顯示2025年度投資金額約{_amt(report_amt)}"
                f"（投資計劃金額完成率{_pct(rate)}）。本次審查工作識別潛在調減金額約{_amt(adj_amt)}，"
                f"經潛在調整後的2025年度投資金額約{_amt(after_amt)}（經調整後投資計劃金額完成率{_pct(after_rate)}）。")
    bullets = [
        ("2025年度獲批的計劃投資金額與報告的投資金額：",
         f"根據{ent_up}提交的2025年度投資計劃方案與投資執行報告，{ent_up}獲批計劃開展{n_plan}個投資項目，"
         f"計劃投資金額約{_amt(plan_amt)}。投資執行報告顯示實際開展其中{n_impl}個投資項目"
         f"（計劃中有{n_zero}個項目未發生投資金額），報告投資金額約{_amt(report_amt)}，"
         f"報告投資計劃金額完成率為{_pct(rate)}。"),
        ("2025年度投資支出金額的潛在調整事項：",
         f"我們在本次審查工作中發現，{ent_up}報告投資金額中存在部分投資支出可能不應確認為2025年度計劃的投資支出"
         f"（涉及{n_adj}個投資項目，合計約{_amt(adj_amt)}）。考慮潛在調減事項後，{ent_up} 2025年度投資支出金額"
         f"約{_amt(after_amt)}，投資計劃金額完成率應為{_pct(after_rate)}。"),
    ]
    return headline, bullets


def _pct(v):
    return f"{v*100:.1f}%" if isinstance(v, (int, float)) and not pd.isna(v) else "—"


# ── 報告用字 helper（逐句對返 scan slide 10/11/15/19/23）────────────────────
_CN_NUM = "零一二三四五六七八九十"


def _cn(n):
    """1→一、7→七、12→十二（報告寫『存在七大類的調整事項』）。"""
    n = int(n)
    if n <= 10:
        return _CN_NUM[n]
    if n < 20:
        return "十" + (_CN_NUM[n - 10] if n > 10 else "")
    return f"{n}"


def _amt(wan):
    """萬 → 報告用字。scan：≥1億寫『6.4億澳門元』，<1億寫『5,527萬澳門元』。"""
    try:
        w = abs(float(wan or 0))
    except (TypeError, ValueError):
        return "—"
    return f"{w/10000:.1f}億澳門元" if w >= 10000 else f"{w:,.0f}萬澳門元"


def _cats_of(ov, col="報告投資金額", n=3, scope=None):
    """由 overview 表攞金額最大嘅幾個範疇名（scan：『主要涉及會議展覽、文化藝術、社區旅遊等…』）。"""
    if ov is None or ov.empty or col not in ov.columns:
        return ""
    d = ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].copy()
    d["_v"] = pd.to_numeric(d[col], errors="coerce").fillna(0)
    d = d[d["_v"] > 0].sort_values("_v", ascending=False)
    named = d[~d["範疇"].astype(str).isin(["其他", "其它"])]      # 「其他」唔好排喺點名範疇最前
    d = (named if len(named) >= min(n, 2) else d).head(n)
    return "、".join(str(r["範疇"]) for _, r in d.iterrows())


def _pg(i, n):
    """scan 導語尾有頁碼標記『（1/2）』。"""
    return f"（{i}/{n}）" if n > 1 else ""


def _bucket_headline(ent_up, bucket, ov):
    """②期後概覽導語（由表自己嘅總計行計，自洽）。"""
    tot = ov[ov["範疇"].astype(str).str.strip() == "總計"]
    if tot.empty:
        return ""
    r = tot.iloc[0]

    def n(c):
        v = r.get(c)
        return float(v) if isinstance(v, (int, float)) and not pd.isna(v) else 0.0
    rep, a, aft = n("報告投資金額"), n("潛在調整金額"), n("潛在調整後投資金額")
    yr, npj = bucket[:4], int(n("項目數量"))
    cats = _cats_of(ov, "報告投資金額", 3)
    tail = (f"，主要涉及{cats}等非博彩投資範疇的{npj}個項目" if cats else f"，涉及{npj}個投資項目")
    return (f"{ent_up}在2025年度執行報告中申報的「因發生期後事項需作後續調整之{yr}年度博彩／非博彩項目」"
            f"投資金額為{_amt(rep)}。本次審查工作識別潛在調減金額約{_amt(a)}，"
            f"經潛在調減後的{yr}年度計劃投資項目在2025年的投資金額約{_amt(aft)}{tail}。")


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


def _prose_2col(prs, title, bullets, per=12, subtitle=None, *, sec=0, headline=None):
    """報告式 2 欄敘述（對 scan slide 16-17：左右兩欄，每欄 navy 小標題 + body）。
    每頁裝幾多由【估算高度】決定，唔會爆版。"""
    if not bullets:
        return
    numbered = bool(bullets) and len(bullets[0]) == 3      # (no, head, body) = scan 編號清單
    W, H = L.size_of(prs)
    colw = (W - 2 * L.MARGIN - 0.24) / 2
    probe = L.HEAD_Y + L.head_h(headline, W)[0] + 0.10
    avail = L.CONTENT_BOTTOM - probe - (0.2 if subtitle else 0)
    if numbered:
        hs_all = [L.est_numbered_h([b], colw, size=7) for b in bullets]
        half_pages, cur, used = [], [], 0.0
        for b, hh in zip(bullets, hs_all):
            if cur and used + hh > avail * 2:
                half_pages.append(cur); cur, used = [], 0.0
            cur.append(b); used += hh
        if cur:
            half_pages.append(cur)
    else:
        half_pages = L.fit_prose(bullets, colw, avail * 2, head_size=7.5, body_size=7)
    for pi, page in enumerate(half_pages):
        suffix = f"（{pi+1}/{len(half_pages)}）" if len(half_pages) > 1 else ""
        slide, W, H, top = _page(prs, sec, title, (headline or "") + suffix)
        if subtitle:
            L.put(slide, L.MARGIN, top, W - 2 * L.MARGIN, 0.18, subtitle, size=6.5,
                  italic=True, color=L.GREY)
            top += 0.20
        # 斷欄：以【總高一半】為目標令左右大致平均（對 scan），但唔可以超過一欄可用高
        lim = L.CONTENT_BOTTOM - top
        hs = [(L.est_numbered_h([it], colw, size=7) if numbered
               else L.est_prose_h([it], colw, head_size=7.5, body_size=7)) for it in page]
        target = sum(hs) / 2.0
        cut, used = len(page), 0.0
        for i, ih in enumerate(hs):
            if i and (used >= target or used + ih > lim):
                cut = i; break
            used += ih
        cut = max(1, cut)
        if numbered:
            L.prose_numbered(L._tb(slide, L.MARGIN, top, colw, lim), page[:cut], size=7)
            if page[cut:]:
                L.prose_numbered(L._tb(slide, L.MARGIN + colw + 0.24, top, colw, lim),
                                 page[cut:], size=7)
        else:
            L.prose_box(slide, L.MARGIN, top, colw, lim, page[:cut], head_size=7.5, body_size=7)
            if page[cut:]:
                L.prose_box(slide, L.MARGIN + colw + 0.24, top, colw, lim, page[cut:],
                            head_size=7.5, body_size=7)
        L.source_note(slide, W, more=(pi < len(half_pages) - 1))


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
    # scan p-06 句式：著重於投入{博彩範疇}等博彩項目，以及{非博彩範疇}等非博彩投資項目 + 前後完成率
    gm = ov[ov["範疇"].astype(str).str.startswith("博彩") &
            ~ov["範疇"].astype(str).str.endswith(("小計", "項目"))]
    ng = ov[~ov["範疇"].astype(str).str.startswith("博彩") &
            ~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))]
    g_cats = _cats_of(gm, "報告投資金額", 2) or "博彩娛樂場場地的優化"
    n_cats = _cats_of(ng, "報告投資金額", 5)
    g_r, ng_r = _rate_of(ov, "博彩項目小計", "投資計劃完成率"), _rate_of(ov, "非博彩項目小計", "投資計劃完成率")
    g_a, ng_a = (_rate_of(ov, "博彩項目小計", "潛在調整後投資計劃完成率"),
                 _rate_of(ov, "非博彩項目小計", "潛在調整後投資計劃完成率"))
    head = (f"{ent_up}的2025年度計劃投資項目著重於投入{g_cats}等博彩項目，以及{n_cats}等非博彩投資項目。"
            f"在報告投資金額中，博彩項目的投資計劃金額完成率為{_pct(g_r)}，"
            f"非博彩項目的投資計劃金額完成率為{_pct(ng_r)}。考慮投資金額的潛在調整後，"
            f"博彩項目的投資計劃金額完成率為{_pct(g_a)}，非博彩項目的平均完成率為{_pct(ng_a)}。")
    _prose_2col(prs, f"2025年度投資計劃執行情況概述  |  {ent_up} 2025年度投資項目的整體執行概況",
                bullets, 12, sec=0, headline=head,
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
        no = (B.ADJ7.index(t) + 1) if t in B.ADJ7 else len(B.ADJ7) + 1
        if t in llm_adj and llm_adj[t]:                   # LLM 寫嘅摘要優先
            bullets.append((no, f"{t}（{_amt(amt)}）：", llm_adj[t])); continue
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
        bullets.append((no, f"{t}（{_amt(amt)}）：", body))
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


def _adj_summary(ent_up, adj, ov=None, sdf=None):
    """潛在調整事項匯總（報告 slide 15）→ 回 (headline, bullets)。逐類型金額。
    ⚠ 用 2025計劃 bucket（唔係合計）：報告調整詳述只計 2025年度計劃，期後另有匯總。"""
    pb = S.BUCKET_ORDER[0]      # "2025年度投資計劃"
    tot_row = adj[adj["潛在調整事項"] == "合計"]
    total = tot_row.iloc[0][pb] if len(tot_row) else 0
    n_type = sum(1 for _, r in adj.iterrows()
                 if r["潛在調整事項"] not in ("合計", "跨年及其他調整")
                 and isinstance(r.get(pb), (int, float)) and abs(r.get(pb, 0)) >= 0.5)
    n_proj = after = None
    if sdf is not None:
        d = sdf[(sdf["_bucket"] == pb) & (pd.to_numeric(sdf["調整_萬"], errors="coerce") != 0)]
        n_proj = int(d["dicj code"].nunique())
    if ov is not None and not ov.empty:
        t = ov[ov["範疇"].astype(str).str.strip() == "總計"]
        if len(t):
            after = t.iloc[0].get("潛在調整後投資金額")
    # scan p-08 句式：存在七大類的調整事項，潛在調減約6.9億澳門元（涉及22個投資項目），經潛在調減後約11.8億澳門元
    headline = (f"基於我們的各項審查程序，我們認為{ent_up}報告的2025年度投資金額中，"
                f"存在{_cn(n_type)}大類的調整事項，潛在調減投資金額約{_amt(total)}"
                + (f"（涉及{n_proj}個投資項目）" if n_proj else "")
                + (f"，經潛在調減後的2025年度計劃的投資金額約{_amt(after)}。" if after is not None else "。"))
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
    global ENT_UP
    ent_up = ENT_UP = entity.upper()
    print(f"build {BUILD_STAMP}")
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
            workers = int(av[av.index("--workers") + 1]) if "--workers" in av else 8   # default 8（`mgm` 一個 command 就並行）
            biao2_dir = av[av.index("--biao2") + 1] if "--biao2" in av else "data/表2"
            print("  由 feed+清單+表2 即場生成 LLM 敘述…")
            try:
                generate_llm_narrative(str(feed), entity, str(qingdan) if qingdan else None,
                                       biao2_dir=biao2_dir, model=model, workers=workers)
            except Exception as e:
                print(f"  ⚠ LLM 生成失敗（{type(e).__name__}: {e}）→ 用現有 json / 清單 fallback")
    llm = _load_llm(entity)     # {entity}_llm_narrative.json 有就用 LLM 文字，否則清單 fallback
    if llm:
        print("    LLM narrative: " + "、".join(f"{k} {len(llm.get(k, {}))}"
                                              for k in ("adj", "cat", "tbl", "proj", "bkt")) + " 段")

    global USE_TEMPLATE
    tmpl = _find_template() if "--use-template" in sys.argv else None   # 預設 fresh 手砌（template 樣式已 hardcode）；--use-template 先開 template
    USE_TEMPLATE = bool(tmpl)
    if tmpl:
        prs = Presentation(str(tmpl))
        _strip_slides(prs)      # drop_rel 正確清走原有 content slides（唔會 duplicate 名 corrupt）
        print(f"    template（master-driven）：{tmpl}  layouts={sum(len(m.slide_layouts) for m in prs.slide_masters)}")
    else:
        prs = Presentation()
        if template:
            prs.slide_width, prs.slide_height = Presentation(str(template)).slide_width, Presentation(str(template)).slide_height
        else:
            prs.slide_width, prs.slide_height = Inches(L.SLIDE_W), Inches(L.SLIDE_H)
        print(f"    無 template → fallback 手砌 formatting（{prs.slide_width/914400:.2f}x"
              f"{prs.slide_height/914400:.2f}in）")

    sdf = S._load(feed, entity)     # 於2025發生 slice（概述 + 金額匯總 共用）

    render_cover(prs, entity)       # 封面（報告 p1）

    # ① 2025年度投資計劃執行情況概述（報告 slide 8-18）
    S1 = "2025年度投資計劃執行情況概述"
    divider(prs, S1, "1", [
        ("1.1  股權架構簡圖及發生投資支出的主體公司", ""),
        ("1.2  2025年度計劃的整體投資支出概況", ""),
        ("1.3  2025年度投資項目的整體執行概況", ""),
        ("1.4  2025年度投資計劃報告投資金額的潛在調整事項匯總", ""),
    ])
    ov = O.overview_by_bucket(sdf, "2025年度投資計劃", plan, cat)
    adj = O.adjustment_bridge(sdf)
    NOTE_RATE = ("註：投資計劃完成率 ＝ 報告投資金額 ／ 獲批的計劃投資金額；潛在調整後完成率 ＝ "
                 "潛在調整後投資金額 ／ 獲批的計劃投資金額。金額單位為萬澳門元。")
    if not ov.empty:      # slide 10-11：表左 + headline/執行敘述右（報告 2 欄式）
        hl, hlb = _headline(ent_up, ov, sdf, plan)
        exb = _exec_bullets(ent_up, ov)
        render_overview_page(prs, f"{S1}  |  {ent_up} 2025年度計劃的整體投資支出及執行概況",
                             hl, ov.fillna(""), hlb + exb, sec=0,
                             table_name=f"{ent_up} 2025年度的整體投資支出概況", note=NOTE_RATE)
        render_category_overview(prs, ent_up, ov, sdf, narr, llm)   # slide 13-14 逐範疇概況（LLM 優先）
        zit = O.zero_investment_text(O.zero_investment_summary(sdf, plan, cat, narr, ent_up), ent_up)
        if zit:      # 報告概述尾段：2025計劃申報投資為零嘅項目（跨年/內部研究/取消）
            _prose_slide(prs, f"{S1}  |  {ent_up} 2025年度計劃申報投資支出為零的項目",
                         [("", x) for x in zit[1:]], headline=zit[0], sec=0)
    ahl, ab = _adj_summary(ent_up, adj, ov, sdf)   # slide 15：表左 + 匯總敘述右
    render_overview_page(prs, f"{S1}  |  {ent_up} 2025年度投資計劃報告投資金額的潛在調整事項匯總",
                         ahl, adj.fillna(""), ab, sec=0,
                         table_name=f"{ent_up} 2025年度投資計劃報告投資金額的潛在調整事項匯總",
                         note="註：金額單位為萬澳門元；括號表示調減。")
    _prose_2col(prs, f"{S1}  |  {ent_up} 2025年度報告投資金額的潛在調整事項（詳述）",
                _adj_detail_bullets(ent_up, adj, sdf, narr, llm), 6, sec=0,
                headline=ahl)   # slide 16-17 詳述（LLM 優先）

    # ② 過往年度投資計劃在2025年繼續執行的審查跟進（報告 slide 19-26）
    S2 = "過往年度投資計劃在2025年繼續執行的審查跟進"
    divider(prs, S2, "2", [
        ("2.1  2024年度投資計劃期後投資金額概覽", ""),
        ("2.2  2024年度投資計劃報告投資金額的潛在調整事項匯總", ""),
        ("2.3  2023年度投資計劃期後投資金額概覽", ""),
        ("2.4  2023年度投資計劃報告投資金額的潛在調整事項匯總", ""),
    ])
    for bk in ["2024年度計劃期後投資", "2023年度計劃期後投資"]:
        ov = O.overview_by_bucket(sdf, bk, plan, cat)
        if not ov.empty:
            render_generic(prs, f"{ent_up} {bk}金額概覽", ov.fillna(""), sec=1,
                           crumb=f"{S2}  |  {bk}金額概覽",
                           headline=_bucket_headline(ent_up, bk, ov),
                           note="註：金額單位為萬澳門元；括號表示調減。",
                           llm=llm, tbl_id=tbl_key("期後概覽", bk))
            render_bucket_adjustment(prs, ent_up, bk, sdf, ov, narr, llm)   # 2.2 / 2.4

    # ③ 本年度審查工作的主要發現（報告 slide 28-40）
    S3 = "本年度審查工作的主要發現"
    divider(prs, S3, "3")
    fs = O.finding_summary(sdf)
    if not fs.empty:
        render_generic(prs, f"{ent_up} 本年度審查工作的主要發現摘要", fs.fillna(""), sec=2,
                       crumb=f"{S3}  |  主要發現摘要",
                       headline=(f"本次審查工作就{ent_up}報告的投資金額識別出{len(fs)}類潛在調整事項，"
                                 f"合計潛在調減約{abs(pd.to_numeric(fs['調整額合計'], errors='coerce').sum()):,.0f}"
                                 f"萬澳門元，摘要如下；逐項說明見後頁。"),
                       note="註：金額單位為萬澳門元；括號表示調減。",
                       llm=llm, tbl_id=tbl_key("發現摘要"))
    if narr:      # 逐調整類型 × 項目：金額(feed) + 事項描述(LLM ground 表2＋清單) / 清單抄字
        b2 = {}
        try:            # 表2＝審查底稿，清單冇料時頂住（加密檔，開唔到就靜靜跳過）
            b2 = B2.load_biao2_struct(av[av.index("--biao2") + 1] if "--biao2" in av else "data/表2",
                               entity, log=lambda *a: None)
        except Exception:
            pass
        render_findings(prs, ent_up, sdf, narr, llm=llm, b2=b2)

    # ④ 其他信息（報告 slide 42-63）
    S4 = "其他信息"
    divider(prs, S4, "4")
    render_generic(prs, f"{ent_up} 2025年發生的投資金額匯總",
                   S.summary_amount(sdf).fillna(""), sec=3,
                   crumb=f"{S4}  |  2025年發生的投資金額匯總",
                   headline=(f"下表匯總{ent_up} 2025年度投資計劃及過往年度計劃期後投資"
                             f"於2025年發生的投資金額（報告投資金額及潛在調整後投資金額）。"),
                   note="註：金額單位為萬澳門元。", llm=llm, tbl_id=tbl_key("金額匯總"))
    for bk in S.BUCKET_ORDER:
        fa = S.facility_activity(sdf, bk)
        if not fa.empty:
            render_generic(prs, f"{ent_up} {bk}區分設施建設/活動舉辦的投資金額", fa.fillna(""), sec=3,
                           crumb=f"{S4}  |  2025年發生的投資金額區分設施建設/活動舉辦",
                           headline=(f"下表按範疇列示{ent_up} {bk}於2025年發生的投資金額，"
                                     f"區分設施建設（資本性支出）及活動舉辦（營運性支出）。"),
                           note="註：金額為潛在調整後金額，單位為萬澳門元。",
                           llm=llm, tbl_id=tbl_key("設施活動", bk))
    for yr in (25, 24, 23):     # 單個項目審查匯總（slide 46-63）
        tab, _ = B.build_year(df, yr, plan.get(yr) if plan else None)
        if tab is not None and not tab.empty:
            R.render_sheet(prs, f"報告年{yr}", tab.fillna(""), list(tab.columns),
                           ent_up=ent_up, sec=3, crumb=f"{S4}  |  單個項目審查結果匯總")

    # ⑥ 附件二 現場走訪（slide 93-100）
    if narr:
        divider(prs, "附件", "6")
        render_site_visits(prs, ent_up, sdf, narr)

    if tmpl:      # template mode：重編 slide 高號，徹底避開 template 殘留 orphan part 撞名 corruption
        _renumber_slides(prs)

    out = Path(f"{entity}_report_llm.pptx")
    try:
        prs.save(out)
    except PermissionError:      # 舊檔喺 PowerPoint 開住鎖住 → 改名唔 crash
        import time
        out = Path(f"{entity}_report_llm_{time.strftime('%H%M%S')}.pptx")
        prs.save(out)
        print(f"⚠ 原檔開住(鎖住)，改存 → {out.name}（開之前記得閂舊 pptx）")
    print(f"✓ {out.resolve()}  共 {len(list(prs.slides))} 頁（概述 + 主要發現 + 金額匯總 + 設施 + 單項審查）")
    if "--dump" in sys.argv:      # 要 cross-check 先加 --dump（慳空間；ok 咗嘅唔使 dump）
        dump = _dump_pptx_text(prs, entity)
        print(f"✓ text dump → {dump.name}（逐版文字，cross-check 用）")


if __name__ == "__main__":
    main()
