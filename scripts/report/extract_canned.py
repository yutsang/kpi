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
    python scripts\\report\\extract_canned.py "MGM…報告.pptx" --skeleton 9,97-108  # 指定邊幾版只抽版式
    python scripts\\report\\extract_canned.py "MGM…報告.pptx" --skeleton none       # 連今年內容一齊抽

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
# 前置(2-6)｜1.1 股權架構圖(9)｜4.4 執行管理流程(68)｜4.5 編制基礎(69-70)｜
# 4.6 本次審查工作執行的程序匯總(71-75)｜第5章 六項KPI(77-86)｜附件全部(88-109：工作範圍 88-96、現場走訪 97-104、
# 藝術品清單 105-106、補充圖片 107-108、封底 109)
# ★ 呢批全部係圖／流程／現場相，冇底層數據砌得出（架構圖、swimlane、走訪相片、
#   藝術品清單本身就係人手做嘅嘢）→ 連圖一齊抽返原位擺，係唯一做到「同原報告一樣」嘅方法。
DEFAULT_SLIDES = "2-6,9,68-75,77-86,88-109"

# ★ 骨架模式：呢批版【只抽版式，唔抽今年內容】。
#   用途係做【明年 report automation 嘅格式底】—— 抄今年嘅走訪相、藝術品清單、
#   官網截圖落去係反效果（明年要換晒）。所以呢啲版：
#     圖片   → 唔抽 blob，改為虛線佔位框（標住原尺寸）
#     表格   → 保留行列／欄闊／底色／字號／框線，格仔文字換成〔…〕
#     文字框 → ≤30 字（標題、欄標、頁腳、資料來源）照留；長文換成〔…〕
#   想連內容一齊抽（例如要重現今年份報告）就 --skeleton none。
DEFAULT_SKELETON = "9,97-108"   # 9 = 股權架構圖（逐家逐年唔同，抄今年冇意義）
SKEL_MARK = "〔…〕"

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


# 原報告好多字同底色用【主題色】唔係直接 RGB。python-pptx 撞到主題色，.rgb 會拋錯 →
# 之前一律當「冇色」存 None，render_canned 就 fallback 去 INK(#222222)，
# 於是抄返出嚟嘅罐頭版顏色全部走樣（diff ⑤ 捉到 breadcrumb 變 #222222）。
# 呢個色盤由 theme1.xml clrScheme + slideMaster clrMap 解出（三個 master 一樣）。
_THEME_RGB = {"BACKGROUND_1": "FFFFFF", "LIGHT_1": "FFFFFF",
              "TEXT_1": "000000", "DARK_1": "000000",
              "BACKGROUND_2": "E5E5E5", "LIGHT_2": "E5E5E5",
              "TEXT_2": "00338D", "DARK_2": "00338D"}


def _color_hex(colorformat):
    """ColorFormat → hex。直接 RGB 就照攞；主題色查色盤；都唔得回 None。"""
    try:
        if colorformat.type is None:
            return None
    except Exception:
        return None
    try:
        return _hex(colorformat.rgb)
    except Exception:
        pass
    try:
        key = str(colorformat.theme_color).split(".")[-1].split(" ")[0]
        return _THEME_RGB.get(key)
    except Exception:
        return None


def _run_style(tf):
    """(pt, bold, 色) —— 攞第一個有字嘅 run。"""
    for p in tf.paragraphs:
        for r in p.runs:
            if not r.text.strip():
                continue
            return (round(r.font.size.pt, 1) if r.font.size else None,
                    bool(r.font.bold), _color_hex(r.font.color))
    return (None, False, None)


_NSA = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _cell_borders(c):
    """{"T": ["00338D", 1.0, dashed], …}。原報告附件表四邊都有 navy 1pt 線，
    唔抄返出嚟嘅表會冇框，同原報告一睇就唔同。"""
    tcPr = c._tc.find(f"{_NSA}tcPr")
    if tcPr is None:
        return None
    out = {}
    for side in "TBLR":
        ln = tcPr.find(f"{_NSA}ln{side}")
        if ln is None:
            continue
        clr = ln.find(f".//{_NSA}srgbClr")
        w = ln.get("w")
        out[side] = [(clr.get("val") if clr is not None else "000000"),
                     round(int(w) / 12700.0, 2) if w else 1.0,
                     ln.find(f"{_NSA}prstDash") is not None]
    return out or None


def _cell(c):
    fill = None
    try:
        fill = _color_hex(c.fill.fore_color)
    except Exception:
        pass
    sz, bold, col = _run_style(c.text_frame)
    d = {"t": c.text, "fill": fill, "fg": col, "size": sz, "bold": bold}
    bd = _cell_borders(c)
    if bd:
        d["bd"] = bd
    return d


def _fill_hex(sh):
    """實色填充 → hex（主題色一樣解到）；漸變／圖片／無填充回 None。"""
    try:
        if sh.fill.type is not None and int(sh.fill.type) == 1:      # MSO_FILL.SOLID
            return _color_hex(sh.fill.fore_color)
    except Exception:
        pass
    return None


def _line_hex(sh):
    try:
        return _color_hex(sh.line.color)
    except Exception:
        return None


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


def _skel_text(t, keep=30):
    """骨架模式嘅文字：短嘅（標題／欄標／頁腳）照留，長文換佔位符 + 原字數。"""
    t = (t or "").strip()
    if len(t) <= keep:
        return t
    return f"{SKEL_MARK}{len(t)}字"


def extract(path, want, out_path, skel=frozenset()):
    prs = Presentation(str(path))
    W, H = _in(prs.slide_width), _in(prs.slide_height)
    img_dir = out_path.parent / (out_path.stem + "_img")
    slides, n_img = [], 0
    for i, sl in enumerate(prs.slides, 1):
        if i not in want:
            continue
        shapes, title, marker = [], "", ""
        for sh in _walk(sl.shapes):
            x, w = _in(sh.left), _in(sh.width)
            y, h = _in(sh.top), _in(sh.height)
            if x + w < 0.05 and y >= 0:            # 畫布外泊住嘅舊嘢
                continue
            try:                                    # 圖（架構圖／流程圖截圖）：blob 抽出嚟做檔
                blob, ext = sh.image.blob, (sh.image.ext or "png").lower()
            except Exception:
                blob = None
            if blob:
                if i in skel:            # 骨架：唔抽相，淨係留返個位
                    shapes.append({"kind": "shape", "x": x, "y": y, "w": w, "h": h,
                                   "text": f"〔圖片 {w:.1f}x{h:.1f}in〕", "fill": None,
                                   "line": "BFBFBF", "size": 8.0, "bold": False,
                                   "color": "8C8C8C", "ph": True})
                    continue
                img_dir.mkdir(parents=True, exist_ok=True)
                fn = f"s{i}_{len(shapes)}.{ext}"
                (img_dir / fn).write_bytes(blob); n_img += 1
                shapes.append({"kind": "pic", "x": x, "y": y, "w": w, "h": h, "file": fn})
                continue
            if getattr(sh, "has_table", False):
                t = sh.table
                cells = [[_cell(c) for c in r.cells] for r in t.rows]
                if i in skel:            # 骨架：保留表結構同樣式，格仔文字換佔位符
                    for row in cells:
                        for c in row:
                            c["t"] = _skel_text(c.get("t", ""))
                shapes.append({
                    "kind": "table", "x": x, "y": y, "w": w, "h": h,
                    "colw": [_in(c.width) for c in t.columns],
                    "rowh": [_in(r.height) for r in t.rows],
                    "cells": cells,
                })
                continue
            fill = _fill_hex(sh)
            txt = (sh.text_frame.text or "").strip() if sh.has_text_frame else ""
            if any(s in txt for s in SKIP):
                continue
            if not txt and not fill:                # 冇字又冇底色 → 冇嘢可以重畫
                continue
            if fill and not txt and w * h > W * H * 0.95:
                continue                            # 成版咁大嘅底色 → 背景，唔好蓋住其他 shape
            if y < 0:                               # 畫布外嘅 UpSlide 章節標記
                marker = marker or txt
                continue
            sz, bold, col = (_run_style(sh.text_frame) if sh.has_text_frame
                             else (None, False, None))
            if (sz or 0) >= 16 and not title:       # 版標題
                title = txt
            # 有底色 = 流程圖／架構圖嘅方框（要連框一齊重畫）；冇底色 = 淨文字框
            shapes.append({"kind": "shape" if fill else "text", "x": x, "y": y, "w": w, "h": h,
                           "size": sz, "bold": bold, "color": col,
                           "text": _skel_text(txt) if i in skel else txt,
                           "fill": fill, "line": _line_hex(sh)})
        if shapes:
            slides.append({"n": i, "title": title, "marker": marker, "shapes": shapes})

    data = {"source": Path(path).name, "slide_w": W, "slides": slides}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    cnt = {k: sum(1 for s in slides for x in s["shapes"] if x["kind"] == k)
           for k in ("text", "shape", "table", "pic")}
    print(f"✓ {out_path}")
    print(f"  {len(slides)} 版：{', '.join(str(s['n']) for s in slides)}")
    print(f"  文字框 {cnt['text']}｜方框(有底色) {cnt['shape']}｜表 {cnt['table']}｜圖 {cnt['pic']}")
    if skel:
        sk_hit = sorted(x["n"] for x in slides if x["n"] in skel)
        print(f"  骨架模式（只抽版式、唔抽今年內容）：{len(sk_hit)} 版 "
              f"→ {', '.join('s%d' % n for n in sk_hit[:14])}{' …' if len(sk_hit) > 14 else ''}")
    if n_img:
        print(f"  圖檔寫咗 {n_img} 個入 {img_dir}（同 JSON 一齊抄去 Windows）")
    for s in slides:
        k = {}
        for x in s["shapes"]:
            k[x["kind"]] = k.get(x["kind"], 0) + 1
        print(f"   · s{s['n']:>3}  {s['title'][:40] or '（冇標題）'}"
              f"　{'、'.join(f'{v}{n}' for n, v in k.items())}")
    print("\n⚠ 呢個 JSON 同圖檔含客戶報告原文 —— 唔好 commit（conf/local/ 已 gitignore）")


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
    sk = args[args.index("--skeleton") + 1] if "--skeleton" in args else DEFAULT_SKELETON
    skel = frozenset() if str(sk).lower() in ("none", "0", "") else _parse_slides(sk)
    if "--out" in args:
        out = Path(args[args.index("--out") + 1])
    else:
        ent = re.split(r"[.\s]", Path(src).name)[0].lower() or "entity"
        out = Path("conf/local") / f"canned_{ent}.json"
    extract(src, _parse_slides(spec), out, skel)


if __name__ == "__main__":
    main()
