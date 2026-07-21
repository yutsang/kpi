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
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
except ImportError:
    print("✗ pip install pandas python-pptx openpyxl"); sys.exit(1)

import build_project_review_table as B
import build_summary_tables as S
import render_review_table_pptx as R

FEED = "tableau_combined_25.csv"


def _find(dirp, entity, ext):
    d = Path(dirp)
    if not d.exists():
        return None
    cands = [p for p in sorted(d.rglob("*"))
             if p.suffix.lower() == ext and entity.lower() in p.name.lower()
             and not p.name.startswith("~$")]
    return cands[0] if cands else None


def render_generic(prs, title, df):
    """render 一張 summary 表（範疇 + 數字欄；欄名有『·』= 2-row group header）。"""
    cols = list(df.columns)
    grouped = any("·" in c for c in cols)
    n = len(df)
    ROWS = 26
    pages = [(i, min(i + ROWS, n)) for i in range(0, n, ROWS)] or [(0, 0)]
    blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    slide_w = prs.slide_width / 914400.0
    widths = [1.7 if c in ("範疇", "項目名稱") else 0.92 for c in cols]
    scale = min(1.0, (slide_w - 0.8) / sum(widths))
    widths = [w * scale for w in widths]
    hrows = 2 if grouped else 1
    for pi, (a, b) in enumerate(pages):
        slide = prs.slides.add_slide(blank)
        tb = slide.shapes.add_textbox(Inches(0.4), Inches(0.25), Inches(slide_w - 0.8), Inches(0.4))
        R._set_title(tb, f"{title}（{pi+1}/{len(pages)}）" if len(pages) > 1 else title)
        sub = df.iloc[a:b]
        ncol = len(cols)
        t = slide.shapes.add_table(hrows + len(sub), ncol, Inches(0.4), Inches(0.75),
                                   Inches(sum(widths)), Inches(0.3 * (hrows + len(sub)))).table
        for ci, w in enumerate(widths):
            t.columns[ci].width = Inches(w)
        if grouped:
            groups = [c.split("·")[0] if "·" in c else c for c in cols]
            ci = 0
            while ci < ncol:
                g = groups[ci]; cj = ci
                while cj + 1 < ncol and groups[cj + 1] == g:
                    cj += 1
                if cj > ci:
                    t.cell(0, ci).merge(t.cell(0, cj))
                R._set(t.cell(0, ci), g, size=7, bold=True, align=PP_ALIGN.CENTER, color=R.WHITE, fill=R.BLUE)
                ci = cj + 1
            for ci, c in enumerate(cols):
                R._set(t.cell(1, ci), c.split("·")[1] if "·" in c else c, size=6.5, bold=True,
                       align=PP_ALIGN.CENTER, color=R.WHITE, fill=R.BLUE)
        else:
            for ci, c in enumerate(cols):
                R._set(t.cell(0, ci), c, size=6.5, bold=True, align=PP_ALIGN.CENTER, color=R.WHITE, fill=R.BLUE)
        for ri, (_, row) in enumerate(sub.iterrows(), start=hrows):
            first = str(row[cols[0]])
            bold = first.endswith(("小計", "合計", "總計"))
            fill = R.GREY if bold else None
            for ci, c in enumerate(cols):
                if ci == 0:
                    txt, al = first, PP_ALIGN.LEFT
                else:
                    txt, al = R.fmt_money(row[c]), PP_ALIGN.RIGHT
                R._set(t.cell(ri, ci), txt, size=6.5, bold=bold, align=al,
                       color=(R.RED if txt.startswith("(") else None), fill=fill)


def main():
    entity = (sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("-") else "mgm").lower()
    feed = Path(FEED)
    if not feed.exists():
        print(f"✗ 揾唔到 feed {feed}（root 應有 tableau_combined_25.csv）"); return
    qingdan = _find("data/投資項目清單", entity, ".xlsx")
    template = _find("data/reports", entity, ".pptx")
    ent_up = entity.upper()
    print(f"entity={ent_up}  feed={feed.name}  清單={qingdan.name if qingdan else '(冇)'}  "
          f"template={template.name if template else '(冇→用 13.33x7.5)'}")

    df = pd.read_csv(feed, low_memory=False)
    df = df[df["entity"].astype(str).str.lower() == entity]
    df["報告年"] = pd.to_numeric(df["報告年"], errors="coerce")
    df["_plan_year"] = df["year_bucket"].map(B._plan_year)
    plan = B.load_plan(qingdan) if qingdan else None

    prs = Presentation()
    if template:
        ref = Presentation(str(template))
        prs.slide_width, prs.slide_height = ref.slide_width, ref.slide_height
    else:
        prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)

    # 1) 單個項目審查匯總（計劃年 25/24/23）
    for yr in (25, 24, 23):
        tab, _ = B.build_year(df, yr, plan.get(yr) if plan else None)
        if tab is not None and not tab.empty:
            R.render_sheet(prs, f"報告年{yr}", tab.fillna(""), list(tab.columns))

    # 2) 金額匯總 + 設施vs活動（於2025發生）
    sdf = S._load(feed, entity)
    amt = S.summary_amount(sdf)
    render_generic(prs, f"{ent_up} 2025年度投資計劃及過往年度期後投資於2025年發生的投資金額匯總", amt.fillna(""))
    for bk in S.BUCKET_ORDER:
        fa = S.facility_activity(sdf, bk)
        if not fa.empty:
            render_generic(prs, f"{ent_up} {bk} 區分設施建設/活動舉辦的投資金額", fa.fillna(""))

    out = Path(f"{entity}_報告數字表.pptx")
    prs.save(out)
    print(f"✓ {out.resolve()}  共 {len(list(prs.slides))} 頁（單個項目審查匯總 + 金額匯總 + 設施vs活動）")


if __name__ == "__main__":
    main()
