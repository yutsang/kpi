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
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    print("✗ pip install pandas python-pptx openpyxl"); sys.exit(1)

# 報告配色（IMG_0105）：navy 表頭白字、section 淺藍、小計灰、總計稍深
HDR = RGBColor(0x1F, 0x38, 0x64)
SEC = RGBColor(0xD9, 0xE1, 0xF2)
SUB = RGBColor(0xE7, 0xE6, 0xE6)
TOT = RGBColor(0xC5, 0xD0, 0xE6)

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


def render_generic(prs, title, df):
    """render 一張表（範疇/項目 + 數字欄；欄名有『·』= 2-row group header）。報告配色 IMG_0105。"""
    cols = list(df.columns)
    grouped = any("·" in c for c in cols)
    n = len(df)
    ROWS = 28
    pages = [(i, min(i + ROWS, n)) for i in range(0, n, ROWS)] or [(0, 0)]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    widths = [1.9 if c in ("範疇", "項目名稱", "潛在調整事項") else
              (2.6 if c == "主要涉及項目" else 0.92) for c in cols]
    scale = min(1.0, (slide_w - 0.8) / sum(widths))
    widths = [w * scale for w in widths]
    hrows = 2 if grouped else 1
    ncol = len(cols)

    def hdr(cell, text, align=PP_ALIGN.CENTER):
        R._set(cell, text, size=7, bold=True, align=align, color=R.WHITE, fill=HDR)

    for pi, (a, b) in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
        R._set_title(tb, f"{title}（{pi+1}/{len(pages)}）" if len(pages) > 1 else title)
        sub = df.iloc[a:b]
        t = slide.shapes.add_table(hrows + len(sub), ncol, Inches(0.4), Inches(0.72),
                                   Inches(sum(widths)), Inches(0.3 * (hrows + len(sub)))).table
        for ci, w in enumerate(widths):
            t.columns[ci].width = Inches(w)
        # navy 表頭 + 左上角「萬澳門元」
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
        # data rows：section 標題 / 小計 / 總計 / 一般
        for ri, (_, row) in enumerate(sub.iterrows(), start=hrows):
            first = str(row[cols[0]]).strip()
            if all(str(row[c]).strip() == "" for c in cols[1:]):        # section 標題行
                for ci in range(ncol):
                    R._set(t.cell(ri, ci), first if ci == 0 else "", size=7, bold=True,
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
                R._set(t.cell(ri, ci), txt, size=7, bold=is_sub, align=al,
                       color=(R.RED if txt.startswith("(") else None), fill=fill)


def render_findings(prs, ent_up, df, narr):
    """③ 主要發現（slide 28-40）：每 canonical 調整類型 → 受影響項目，
    金額(feed 報告/調整) + 清單抄字(KPMG分析發現 / 管理層解釋)。text slides，每頁 3 個項目。"""
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    grey = RGBColor(0x40, 0x40, 0x40)
    for adj in B.ADJ7:
        sub = d[(d["_adj"] == adj) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        if sub.empty:
            continue
        projs = sub.groupby("dicj code").agg(名稱=("project", "first"), 報告=("調整前_萬", "sum"),
                                             調整=("調整_萬", "sum")).reset_index()
        projs = projs.reindex(projs["調整"].abs().sort_values(ascending=False).index)
        recs = []
        for _, p in projs.iterrows():
            nr = narr.get(N._norm_code(p["dicj code"]), {})
            recs.append((str(p["dicj code"]), str(p["名稱"]), p["報告"], p["調整"],
                         nr.get("KPMG分析發現", ""), nr.get("管理層解釋", "")))
        pages = [recs[i:i + 3] for i in range(0, len(recs), 3)]
        for pi, page in enumerate(pages):
            slide = prs.slides.add_slide(blank)
            tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.22), Inches(slide_w - 0.8), Inches(0.4))
            R._set_title(tb, f"{ent_up} 主要發現：{adj}"
                         + (f"（{pi+1}/{len(pages)}）" if len(pages) > 1 else ""))
            y = 0.78
            for code, name, rep, adjv, find, mgmt in page:
                box = slide.shapes.add_textbox(Inches(0.5), Inches(y), Inches(slide_w - 1.0), Inches(1.95))
                tf = box.text_frame; tf.word_wrap = True
                r0 = tf.paragraphs[0].add_run()
                r0.text = f"{code}　{name[:34]}　│　報告 {R.fmt_money(rep)}／調整 {R.fmt_money(adjv)} 萬"
                r0.font.bold = True; r0.font.size = Pt(9); r0.font.color.rgb = HDR
                r0.font.name = "Microsoft JhengHei"
                if find:
                    r1 = tf.add_paragraph().add_run(); r1.text = "發現：" + find[:230]
                    r1.font.size = Pt(8); r1.font.name = "Microsoft JhengHei"
                if mgmt:
                    r2 = tf.add_paragraph().add_run(); r2.text = "管理層解釋：" + mgmt[:190]
                    r2.font.size = Pt(8); r2.font.color.rgb = grey; r2.font.name = "Microsoft JhengHei"
                y += 2.05


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
    narr = N.load_narrative(qingdan) if qingdan else {}     # 清單 by-project narrative（抄字）
    if narr:
        print(f"    清單 narrative: {sum(1 for r in narr.values() if r.get('KPMG分析發現'))} 個項目有發現")

    prs = Presentation()
    if template:
        ref = Presentation(str(template))
        prs.slide_width, prs.slide_height = ref.slide_width, ref.slide_height
    else:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    sdf = S._load(feed, entity)     # 於2025發生 slice（概述 + 金額匯總 共用）

    # ① 2025年度投資計劃執行情況概述（報告 slide 8-18）
    divider(prs, "一、2025年度投資計劃執行情況概述")
    ov = O.overview_by_bucket(sdf, "2025年度投資計劃", plan)
    if not ov.empty:
        render_generic(prs, f"{ent_up} 2025年度投資項目的整體執行概況", ov.fillna(""))
    render_generic(prs, f"{ent_up} 2025年度投資計劃報告投資金額的潛在調整事項匯總",
                   O.adjustment_bridge(sdf).fillna(""))

    # ② 過往年度投資計劃在2025年繼續執行的審查跟進（報告 slide 19-26）
    divider(prs, "二、過往年度投資計劃在2025年繼續執行的審查跟進")
    for bk in ["2024年度計劃期後投資", "2023年度計劃期後投資"]:
        ov = O.overview_by_bucket(sdf, bk, plan)
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

    out = Path(f"{entity}_報告數字表.pptx")
    prs.save(out)
    print(f"✓ {out.resolve()}  共 {len(list(prs.slides))} 頁（單個項目審查匯總 + 金額匯總 + 設施vs活動）")


if __name__ == "__main__":
    main()
