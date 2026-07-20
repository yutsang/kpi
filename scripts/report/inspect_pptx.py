#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_pptx.py — 拆解一份 KPMG 表二報告 PPTX 嘅完整結構，俾我哋設計「底層數據 →
自動生成報告」pipeline：睇清楚邊啲係數字（chart/table，來自 Tableau）、邊啲文字係
template（固定模板，機械填）vs descriptive（要 LLM 寫）。

用法（Windows，喺 kpi-main 底下）：
    pip install python-pptx        # 如未裝
    python scripts\\report\\inspect_pptx.py "data\\reports\\MGM.2025年年度投資計劃執行情況審查專項工作報告.初稿.pptx"
    # 或成個 data\\reports\\ 一次過
    python scripts\\report\\inspect_pptx.py "data\\reports"

輸出：同名 .inspect.txt（貼返嚟俾我）＋ stdout。逐 slide 印：
  · slide index / layout 名 / 尺寸
  · 每個 shape：type、name、placeholder 類型、位置(吋)
  · text frame：逐段逐 run（文字全文＋bold/size/顏色）→ 睇 wording 判 template/descriptive
  · table：逐行逐格文字（＋合併/表頭）
  · chart：圖種、categories、每個 series 名＋**逐點數值**（= Tableau 數）
  · picture：標記圖片（可能係 Tableau screenshot）＋ alt text
  · group：遞迴入去
  · notes：備註頁文字
另出一個「疑似 placeholder / 數字 token」滙總（XX / __ / 0.0 / {{}} / 【】等），
幫手快速分辨 template 空格。
"""
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("✗ 未裝 python-pptx → 執行：pip install python-pptx")
    sys.exit(1)


def _emu_in(v):
    try:
        return round(Emu(v).inches, 2)
    except Exception:
        return "?"


def _font_desc(run):
    f = run.font
    bits = []
    if f.bold:
        bits.append("B")
    if f.italic:
        bits.append("I")
    if f.size is not None:
        try:
            bits.append(f"{int(f.size.pt)}pt")
        except Exception:
            pass
    try:
        if f.color and f.color.type is not None and f.color.rgb is not None:
            bits.append(f"#{f.color.rgb}")
    except Exception:
        pass
    return ",".join(bits) if bits else "-"


# template 空格 / 數字 token 嘅疑似 pattern（幫手分 template vs descriptive）
_PLACEHOLDER = re.compile(
    r"\{\{.*?\}\}|【.*?】|＿{2,}|_{2,}|X{2,}|[xX]{2,}|〇{2,}|○{2,}|"
    r"\bTBD\b|待填|待補|\[.*?\]|（.*?填.*?）|\.{3,}")
_NUM = re.compile(r"-?\d[\d,]*(?:\.\d+)?%?")


def _classify_text(t: str) -> str:
    """粗略提示：純數字/短標籤=likely template 值；有句子=likely descriptive。"""
    s = t.strip()
    if not s:
        return ""
    if _PLACEHOLDER.search(s):
        return "  ⟸ [疑似 PLACEHOLDER]"
    letters = re.sub(r"[\d\s,.\-%()（）]", "", s)
    if len(letters) <= 3:                       # 幾乎全數字/符號
        return "  ⟸ [疑似 數字/短值→template]"
    if len(s) >= 40 and ("。" in s or "，" in s or "、" in s):
        return "  ⟸ [疑似 敘述句→descriptive/LLM]"
    return ""


def dump_textframe(tf, indent, out):
    for pi, para in enumerate(tf.paragraphs):
        runs = para.runs
        full = "".join(r.text for r in runs) or para.text
        if not full.strip() and not runs:
            continue
        lvl = getattr(para, "level", 0) or 0
        hint = _classify_text(full)
        out(f"{indent}  ¶[{pi}] lvl{lvl}: {full!r}{hint}")
        # run 級 font（睇同段落內 bold/非 bold 混合，同 adhoc bold 問題一樣要睇）
        if len(runs) > 1 or (runs and _font_desc(runs[0]) != "-"):
            for ri, r in enumerate(runs):
                if r.text.strip():
                    out(f"{indent}      run[{ri}] {_font_desc(r)}: {r.text!r}")


def dump_table(tbl, indent, out):
    nr, nc = len(tbl.rows), len(tbl.columns)
    out(f"{indent}  TABLE {nr}×{nc}")
    for ri, row in enumerate(tbl.rows):
        cells = []
        for ci, cell in enumerate(row.cells):
            txt = cell.text.replace("\n", "⏎").strip()
            cells.append(f"[{ri},{ci}]{txt[:40]}")
        out(f"{indent}    " + " | ".join(cells))


def dump_chart(chart, indent, out):
    try:
        ctype = str(chart.chart_type)
    except Exception:
        ctype = "?"
    out(f"{indent}  CHART type={ctype}")
    try:
        cats = list(chart.plots[0].categories) if chart.plots else []
        out(f"{indent}    categories({len(cats)}): {[str(c) for c in cats]}")
    except Exception as e:
        out(f"{indent}    categories: <讀唔到 {e}>")
    try:
        for s in chart.series:
            vals = list(s.values)
            out(f"{indent}    series {getattr(s,'name',None)!r}: {vals}")
    except Exception as e:
        out(f"{indent}    series: <讀唔到 {e}>")


def dump_shape(sh, indent, out, idx=""):
    try:
        st = sh.shape_type
    except Exception:
        st = None
    name = getattr(sh, "name", "")
    ph = ""
    try:
        if sh.is_placeholder:
            ph = f" PH={sh.placeholder_format.type}"
    except Exception:
        pass
    pos = ""
    try:
        pos = f" @({_emu_in(sh.left)},{_emu_in(sh.top)}) {_emu_in(sh.width)}x{_emu_in(sh.height)}in"
    except Exception:
        pass
    out(f"{indent}▸ shape{idx} [{st}] name={name!r}{ph}{pos}")

    if st == MSO_SHAPE_TYPE.GROUP:
        for j, child in enumerate(sh.shapes):
            dump_shape(child, indent + "   ", out, idx=f"{idx}.{j}")
        return
    if getattr(sh, "has_chart", False):
        dump_chart(sh.chart, indent, out)
    if getattr(sh, "has_table", False):
        dump_table(sh.table, indent, out)
    if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
        if sh.text_frame.text.strip():
            dump_textframe(sh.text_frame, indent, out)
    if st == MSO_SHAPE_TYPE.PICTURE:
        alt = ""
        try:
            alt = sh._element.getparent  # noqa
        except Exception:
            pass
        out(f"{indent}  PICTURE（可能係 Tableau 截圖／logo）")


def inspect_file(path: Path):
    lines = []

    def out(s=""):
        lines.append(s)
    out(f"\n{'#'*78}\n# {path.name}")
    try:
        prs = Presentation(str(path))
    except Exception as e:
        out(f"  ✗ 開唔到: {type(e).__name__}: {e}")
        print("\n".join(lines))
        return lines
    out(f"  投影片尺寸: {_emu_in(prs.slide_width)}x{_emu_in(prs.slide_height)}in，"
        f"共 {len(prs.slides)} 頁")
    for si, slide in enumerate(prs.slides):
        lay = ""
        try:
            lay = slide.slide_layout.name
        except Exception:
            pass
        out(f"\n{'='*70}\n## SLIDE {si+1}/{len(prs.slides)}  layout={lay!r}  shapes={len(slide.shapes)}")
        for k, sh in enumerate(slide.shapes):
            dump_shape(sh, "  ", out, idx=str(k))
        # notes
        try:
            if slide.has_notes_slide:
                nt = slide.notes_slide.notes_text_frame.text.strip()
                if nt:
                    out(f"  📝 NOTES: {nt!r}")
        except Exception:
            pass
    txt = "\n".join(lines)
    print(txt)
    outp = path.with_suffix(".inspect.txt")
    try:
        outp.write_text(txt, encoding="utf-8")
        print(f"\n✓ 寫入 {outp}（貼返嚟俾我）")
    except Exception as e:
        print(f"\n⚠ 寫唔到 txt: {e}")
    return lines


def main():
    args = sys.argv[1:]
    if not args:
        print('俾 pptx 路徑或 data\\reports 資料夾')
        return
    targets = []
    for a in args:
        p = Path(a)
        if p.is_dir():
            targets += sorted(p.glob("*.pptx"))
        elif p.exists():
            targets.append(p)
        else:
            print(f"✗ 唔存在: {p}")
    for t in targets:
        if t.name.startswith("~$"):
            continue
        inspect_file(t)


if __name__ == "__main__":
    main()
