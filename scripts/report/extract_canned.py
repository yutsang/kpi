#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_canned.py — 由項目組原報告 pptx 抽【罐頭版】（法律聲明／呈送函／注意事項／
縮寫定義／附件1 工作範圍／封底）成一份 JSON，餵返 build_report 直接出返嗰幾版。

點解要咁做：呢批版純文字＋固定表，冇任何數據依賴，逐段人手貼返落 code 又長又易錯，
而且報告原文係客戶機密，唔可以入 public repo。所以：
    code 入 git（呢個檔）  ｜  抽出嚟嘅 JSON 唔入 git（gitignore 嘅 conf/local/）

用法（Windows，原報告喺手嗰部機）：
    python scripts\\report\\extract_canned.py "MGM…報告.pptx"
    python scripts\\report\\extract_canned.py "MGM…報告.pptx" --slides 2-6,88-96,109
    python scripts\\report\\extract_canned.py "MGM…報告.pptx" --out conf\\local\\canned_mgm.json

預設抽邊幾版：由 diff_report ③ 揾到嗰批（2-6 前置、88-96 附件1、109 封底）。
抽咗之後 build_report 會自動搵 conf\\local\\canned_{entity}.json，有就出返嗰幾版。
"""
import json
import re
import sys
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("✗ pip install python-pptx"); sys.exit(1)

EMU_IN = 914400.0
DEFAULT_SLIDES = "2-6,88-96,109"

# 模板 placeholder／工作痕跡：唔算內容
SKIP = ("已更新表格", "定稿後還需手動更新", "目錄手動修改為繁體字", "DO NOT DELETE",
        "Workspace (", "Click to edit", "单击以编辑", "此处插入", "此处添加",
        "点击图标", "Click icon", "‹#›", "文檔分類", "文档分类")


def _in(v):
    return round((v or 0) / EMU_IN, 3)


def _walk(shapes):
    for sh in shapes:
        st = str(sh.shape_type) if sh.shape_type is not None else ""
        if st.startswith("GROUP"):
            try:
                yield from _walk(sh.shapes); continue
            except Exception:
                pass
        yield sh


def _hex(c):
    try:
        return f"{c[0]:02X}{c[1]:02X}{c[2]:02X}"
    except Exception:
        return None


def _run_style(tf):
    """(pt, bold, 色) —— 攞第一個有字嘅 run。"""
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            col = None
            try:
                col = _hex(r.font.color.rgb)
            except Exception:
                pass
            return (round(r.font.size.pt, 1) if r.font.size else None,
                    bool(r.font.bold), col)
    return (None, False, None)


def _cell(c):
    fill = None
    try:
        fill = _hex(c.fill.fore_color.rgb)
    except Exception:
        pass
    sz, bold, col = _run_style(c.text_frame)
    return {"t": c.text, "fill": fill, "fg": col, "size": sz, "bold": bold}


def _parse_slides(spec):
    out = set()
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out |= set(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return out


def extract(path, want, out_path):
    prs = Presentation(str(path))
    W = _in(prs.slide_width)
    slides = []
    for i, sl in enumerate(prs.slides, 1):
        if i not in want:
            continue
        shapes, title, marker = [], "", ""
        for sh in _walk(sl.shapes):
            x, w = _in(sh.left), _in(sh.width)
            y, h = _in(sh.top), _in(sh.height)
            if x + w < 0.05 and y >= 0:            # 畫布外泊住嘅舊嘢
                continue
            if getattr(sh, "has_table", False):
                t = sh.table
                shapes.append({
                    "kind": "table", "x": x, "y": y, "w": w, "h": h,
                    "colw": [_in(c.width) for c in t.columns],
                    "rowh": [_in(r.height) for r in t.rows],
                    "cells": [[_cell(c) for c in r.cells] for r in t.rows],
                })
                continue
            if not sh.has_text_frame:
                continue
            txt = (sh.text_frame.text or "").strip()
            if not txt or any(s in txt for s in SKIP):
                continue
            if y < 0:                               # 畫布外嘅 UpSlide 章節標記
                marker = marker or txt
                continue
            sz, bold, col = _run_style(sh.text_frame)
            if (sz or 0) >= 16 and not title:       # 版標題
                title = txt
            shapes.append({"kind": "text", "x": x, "y": y, "w": w, "h": h,
                           "size": sz, "bold": bold, "color": col, "text": txt})
        if shapes:
            slides.append({"n": i, "title": title, "marker": marker, "shapes": shapes})

    data = {"source": Path(path).name, "slide_w": W, "slides": slides}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    n_t = sum(1 for s in slides for x in s["shapes"] if x["kind"] == "text")
    n_b = sum(1 for s in slides for x in s["shapes"] if x["kind"] == "table")
    print(f"✓ {out_path}")
    print(f"  {len(slides)} 版：{', '.join(str(s['n']) for s in slides)}")
    print(f"  文字框 {n_t} 個、表 {n_b} 張")
    for s in slides:
        print(f"   · s{s['n']:>3}  {s['title'][:44] or '（冇標題）'}"
              f"　{len(s['shapes'])} shape")
    print("\n⚠ 呢個 JSON 含客戶報告原文 —— 唔好 commit（conf/local/ 已 gitignore）")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = sys.argv[1:]
    pos = [a for a in args if not a.startswith("--")]
    if not pos:
        print(__doc__); return
    src = pos[0]
    spec = args[args.index("--slides") + 1] if "--slides" in args else DEFAULT_SLIDES
    if "--out" in args:
        out = Path(args[args.index("--out") + 1])
    else:
        ent = re.split(r"[.\s]", Path(src).name)[0].lower() or "entity"
        out = Path("conf/local") / f"canned_{ent}.json"
    extract(src, _parse_slides(spec), out)


if __name__ == "__main__":
    main()
