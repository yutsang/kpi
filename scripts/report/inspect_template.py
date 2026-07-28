#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_template.py — dump pptx 嘅 slide master + 每個 slide layout + placeholders（idx/type/名/位置/字體）。
用嚟評估：可唔可以直接開間公司空 template 做 base、用佢 layouts 填內容（formatting 來自 master），
唔使 Claude 手砌顏色/furniture。

用法：python scripts\\report\\inspect_template.py "你的空template.pptx"
   （用【空 template】——得 master+layouts、冇內容；唔好用填好嘅初稿報告）
"""
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    print("✗ pip install python-pptx"); sys.exit(1)


def _in(emu):
    try:
        return round(Emu(emu).inches, 2)
    except Exception:
        return emu


def _ph_type(ph):
    try:
        return str(ph.placeholder_format.type).split()[0]
    except Exception:
        return "?"


def _font_of(ph):
    try:
        r = ph.text_frame.paragraphs[0].runs
        f = (r[0].font if r else ph.text_frame.paragraphs[0].font)
        col = ""
        try:
            col = f"#{f.color.rgb}" if f.color and f.color.type is not None else ""
        except Exception:
            col = ""
        return f"{f.name or ''} {f.size.pt if f.size else ''}pt {'B' if f.bold else ''} {col}".strip()
    except Exception:
        return ""


def main():
    if len(sys.argv) < 2:
        print("俾 template pptx 路徑"); return
    prs = Presentation(sys.argv[1])
    print(f"slide size: {_in(prs.slide_width)} x {_in(prs.slide_height)} in")
    print(f"slide_masters: {len(prs.slide_masters)}　| 直接可用 layouts: {len(prs.slide_layouts)}")
    print("=" * 90)
    for mi, master in enumerate(prs.slide_masters):
        print(f"\n■ MASTER {mi}: name='{master.name}'")
        # master 背景色
        try:
            fill = master.background.fill
            print(f"   background fill type: {fill.type}")
        except Exception:
            pass
        for li, layout in enumerate(master.slide_layouts):
            phs = list(layout.placeholders)
            print(f"\n   ▸ Layout {li}: '{layout.name}'  （{len(phs)} placeholders）")
            for ph in phs:
                print(f"       [idx {ph.placeholder_format.idx}] {_ph_type(ph):<10} "
                      f"name='{ph.name}'  pos=({_in(ph.left)},{_in(ph.top)},{_in(ph.width)},{_in(ph.height)})"
                      f"  font：{_font_of(ph)}")
    print("\n" + "=" * 90)
    print("→ paste 返：我睇 layout 名 + placeholder 佈局，決定可唔可以直接用 master 填內容")


if __name__ == "__main__":
    main()
