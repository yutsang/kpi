#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_pptx.py — 報告 pptx 版面體檢（唔使開 PowerPoint 逐版睇）。逐版報：

  OVERFLOW-X/Y   shape 超出 slide 邊界（右／底）——「超出 border」直接捉到
  TABLE-GROW     表格估算實際高度 > requested（PowerPoint 會自動長高 row）→ 會爆版
  TEXT-OVERFLOW  文字框估算 wrap 高度 > 框高
  NO-TEXT        淨得表 / 圖，冇任何敘述文字（報告唔應該有淨表冇字嘅版）
  EMPTY          乜都冇
  OFF-PALETTE    用咗 KPMG 品牌色以外嘅顏色（IMG_0420 Visual identity overview）
  NO-FURNITURE   冇 breadcrumb / footer / 頁碼

用法：
    python scripts\\report\\inspect_pptx.py mgm_report_llm.pptx            # 體檢
    python scripts\\report\\inspect_pptx.py mgm_report_llm.pptx --slide 12 # 淨睇某版 shape 清單
    python scripts\\report\\inspect_pptx.py mgm_report_llm.pptx --render   # 用 PowerPoint 出 PDF/PNG（Mac）
"""
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
try:
    from pptx import Presentation
    from pptx.util import Emu
except ImportError:
    print("✗ pip install python-pptx"); sys.exit(1)

import layout as L

EMU_IN = 914400.0
PALETTE = {  # KPMG 品牌色 + 由品牌色派生嘅表格 tint（其餘算 off-palette）
    "00338D": "KPMG Blue", "005EB8": "Medium Blue", "0091DA": "Light Blue",
    "483698": "Violet", "470A68": "Purple", "6D2077": "Light Purple", "00A3A1": "Green",
    "FFFFFF": "White", "EEF1F8": "tint-section", "D9E1F2": "tint-subtotal",
    "BDD7EE": "tint-total", "BFBFBF": "格線", "222222": "內文", "333333": "內文",
    "595959": "註", "8C8C8C": "breadcrumb", "C8C8C8": "分隔", "0C233C": "封面深底",
    "C8C8D0": "分隔頁副題",
}


def _in(v):
    return (v or 0) / EMU_IN


def _shape_text(sh):
    try:
        return sh.text_frame.text if sh.has_text_frame else ""
    except Exception:
        return ""


def _colors(sh, out):
    """收集 shape 用到嘅 srgbClr（fill + font）。"""
    for el in sh._element.iter():
        if el.tag.endswith("}srgbClr"):
            v = el.get("val")
            if v:
                out.add(v.upper())


def _table_h(tbl):
    """(requested_h, 估算實際 h)：PowerPoint 會按內容長高 row。用返 layout.row_h 同一把尺，
    唔會同 build 時嘅分頁計算行開（否則成堆假警報）。"""
    req = sum(_in(r.height) for r in tbl.rows)
    widths = [_in(c.width) for c in tbl.columns]
    est = 0.0
    for r in tbl.rows:
        sizes = [run.font.size.pt for cell in r.cells for p in cell.text_frame.paragraphs
                 for run in p.runs if run.font.size]
        size = max(sizes) if sizes else 10.0
        txts, ws = [], []
        for ci, cell in enumerate(r.cells):
            if getattr(cell, "is_spanned", False):       # 被合併掉嘅格：唔重複計
                continue
            n = cell.span_width if getattr(cell, "is_merge_origin", False) else 1
            txts.append(cell.text)
            ws.append(sum(widths[ci:ci + n]))            # 合併格：加埋跨住嘅欄闊
        need = L.row_h(txts, ws, size)
        est += max(_in(r.height), need)
    return req, est


def _text_h(sh):
    """(框高, 估算 wrap 高)。"""
    box_h = _in(sh.height)
    w = _in(sh.width)
    need = 0.0
    for p in sh.text_frame.paragraphs:
        runs = [r for r in p.runs if r.text]
        if not runs:
            continue
        size = max((r.font.size.pt for r in runs if r.font.size), default=10.0)
        txt = "".join(r.text for r in runs)
        need += L.est_lines(txt, w, size) * size * 1.3 / 72.0
        need += (p.space_after.pt if p.space_after else 0) / 72.0
        need += (p.space_before.pt if p.space_before else 0) / 72.0
    return box_h, need


def audit(path, tol=0.02):
    prs = Presentation(str(path))
    W, H = _in(prs.slide_width), _in(prs.slide_height)
    print(f"── {Path(path).name}：{len(prs.slides._sldIdLst)} 版，{W:.2f} x {H:.2f} in")
    if abs(W - L.SLIDE_W) > 0.05 or abs(H - L.SLIDE_H) > 0.05:
        print(f"   ⚠ slide 尺寸 ≠ 報告標準 {L.SLIDE_W} x {L.SLIDE_H} in")
    bad, palette_hits = 0, {}
    for i, slide in enumerate(prs.slides, 1):
        issues, has_table, has_prose, has_foot = [], False, False, False
        for sh in slide.shapes:
            x, y, w, h = _in(sh.left), _in(sh.top), _in(sh.width), _in(sh.height)
            name = sh.shape_type
            if x + w > W + tol:
                issues.append(f"OVERFLOW-X  {name} 右邊到 {x+w:.2f}in > {W:.2f}")
            if y + h > H + tol:
                issues.append(f"OVERFLOW-Y  {name} 底到 {y+h:.2f}in > {H:.2f}")
            if sh.has_table:
                has_table = True
                req, est = _table_h(sh.table)
                if est > req + 0.05:
                    issues.append(f"TABLE-GROW  表 requested {req:.2f}in → 估實際 {est:.2f}in")
                if y + est > H + tol:
                    issues.append(f"TABLE-OVERFLOW 表底到 {y+est:.2f}in > {H:.2f}（會被切）")
            txt = _shape_text(sh)
            if txt.strip():
                if sh.has_text_frame and not sh.has_table:
                    bh, need = _text_h(sh)
                    if need > bh + 0.08 and y + need > H:
                        issues.append(f"TEXT-OVERFLOW 文字框 {bh:.2f}in → 需 {need:.2f}in（超版底）")
                if "初稿" in txt:
                    has_foot = True
                if len(txt.strip()) > 40 and "初稿" not in txt and "©" not in txt:
                    has_prose = True
            _colors(sh, palette_hits.setdefault(i, set()))
        n_sh = len(slide.shapes)
        # 封面 / 章節分隔＝滿版深色底 → 冇 footer 係正常（跟 scan）
        cover = any(_in(s.width) > W - 0.1 and _in(s.height) > H - 0.1 for s in slide.shapes)
        if n_sh == 0:
            issues.append("EMPTY  冇任何 shape")
        elif has_table and not has_prose:
            issues.append("NO-TEXT  淨得表冇敘述文字")
        if not has_foot and n_sh > 3 and not cover:
            issues.append("NO-FURNITURE  冇 footer/頁碼")
        off = {c for c in palette_hits.get(i, set()) if c not in PALETTE}
        if off:
            issues.append("OFF-PALETTE  " + ", ".join("#" + c for c in sorted(off)))
        if issues:
            bad += 1
            print(f"\n  [slide {i}] {n_sh} shapes")
            for m in issues:
                print(f"      ✗ {m}")
    print(f"\n── 有問題 {bad} / {len(prs.slides._sldIdLst)} 版")
    return bad


def dump_slide(path, n):
    prs = Presentation(str(path))
    slide = list(prs.slides)[n - 1]
    W, H = _in(prs.slide_width), _in(prs.slide_height)
    print(f"── slide {n}（{W:.2f}x{H:.2f}in）")
    for sh in slide.shapes:
        x, y, w, h = _in(sh.left), _in(sh.top), _in(sh.width), _in(sh.height)
        kind = "TABLE" if sh.has_table else ("TEXT" if sh.has_text_frame else str(sh.shape_type))
        extra = ""
        if sh.has_table:
            req, est = _table_h(sh.table)
            extra = f"  {len(sh.table.rows)}x{len(sh.table.columns)} req {req:.2f} est {est:.2f}in"
        print(f"  {kind:7s} x={x:5.2f} y={y:5.2f} w={w:5.2f} h={h:5.2f}{extra}")
        t = _shape_text(sh).replace("\n", " ⏎ ")
        if t.strip():
            print(f"          「{t[:150]}」")


def _osa(script, timeout=600):
    return subprocess.run(["osascript", "-e", script], capture_output=True, text=True,
                          timeout=timeout)


def render(path, dpi=80):
    """用 Microsoft PowerPoint（Mac）出 PDF，再 pdftoppm 出 PNG → 睇真實版面（含 table auto-grow）。

    ⚠ PowerPoint 嘅 AppleScript `open` 只喺【冷啟動】先穩陣：app 已經開住時會 -9074。
      所以每次都先 quit 再 open；第一個 AppleEvent 會 -1712 timeout（其實開緊）→ 忽略，改為 poll。"""
    import time
    src = Path(path).resolve()
    out = src.parent / (src.stem + "_render")
    out.mkdir(exist_ok=True)
    pdf = out / (src.stem + ".pdf")
    if pdf.exists():
        pdf.unlink()
    _osa('tell application "Microsoft PowerPoint" to quit saving no', timeout=120)
    time.sleep(5)
    _osa(f'tell application "Microsoft PowerPoint" to open POSIX file "{src}"', timeout=600)
    for _ in range(60):          # poll：cold launch 慢，AppleEvent timeout 唔代表失敗
        r = _osa('tell application "Microsoft PowerPoint" to get count of presentations', 120)
        if r.stdout.strip().isdigit() and int(r.stdout.strip()) > 0:
            break
        time.sleep(2)
    else:
        print("✗ PowerPoint 開唔到檔"); return
    r = _osa(f'''with timeout of 540 seconds
tell application "Microsoft PowerPoint"
    save presentation 1 in POSIX file "{pdf}" as save as PDF
    close presentation 1 saving no
end tell
end timeout''', timeout=600)
    if not pdf.exists():
        print("✗ PowerPoint 出 PDF 失敗：", (r.stderr or "").strip()[:300]); return
    subprocess.run(["pdftoppm", "-png", "-r", str(dpi), str(pdf), str(out / "s")], check=False)
    print(f"✓ {pdf}\n✓ PNG → {out}/s-*.png")


# ── 自畫 preview（唔使 PowerPoint / LibreOffice）───────────────────────────
FONT_DIRS = ["/System/Library/Fonts/Supplemental", "/System/Library/Fonts",
             "C:/Windows/Fonts", str(Path.home() / "Library/Fonts")]
FONT_FILES = {  # (cjk?, bold?) → 候選檔名
    (False, False): ["Arial.ttf", "arial.ttf", "Helvetica.ttc"],
    (False, True): ["Arial Bold.ttf", "arialbd.ttf", "Arial.ttf"],
    (True, False): ["STHeiti Light.ttc", "msyh.ttc", "PingFang.ttc", "Arial Unicode.ttf"],
    (True, True): ["STHeiti Medium.ttc", "msyhbd.ttc", "PingFang.ttc", "Arial Unicode.ttf"],
}
_fc = {}


def _font(cjk, bold, px):
    from PIL import ImageFont
    key = (cjk, bold, px)
    if key in _fc:
        return _fc[key]
    for name in FONT_FILES[(cjk, bold)]:
        for d in FONT_DIRS:
            p = Path(d) / name
            if p.exists():
                try:
                    _fc[key] = ImageFont.truetype(str(p), px); return _fc[key]
                except Exception:
                    pass
    _fc[key] = ImageFont.load_default()
    return _fc[key]


def _runs(tf):
    """text_frame → [(para_align, [(text, size_pt, bold, rgb)])]。
    ⚠ 要行返 <a:br/>（硬換行）；python-pptx 嘅 p.runs 會 skip 咗佢哋。"""
    out = []
    for p in tf.paragraphs:
        rs, last = [], (12.0, False, (0, 0, 0))
        for el in p._p:
            tag = el.tag.split("}")[-1]
            if tag == "br":
                rs.append(("\n",) + last); continue
            if tag != "r":
                continue
            r = next((x for x in p.runs if x._r is el), None)
            if r is None or not r.text:
                continue
            try:
                col = r.font.color.rgb
                rgb = (col[0], col[1], col[2]) if col else (0, 0, 0)
            except Exception:
                rgb = (0, 0, 0)
            last = (r.font.size.pt if r.font.size else 12.0, bool(r.font.bold), rgb)
            rs.append((r.text,) + last)
        out.append((str(p.alignment), rs))
    return out


def _draw_text(dr, runs, x, y, w, scale, *, align="LEFT", valign_h=None):
    """畫 wrap 文字（逐 run 接住排，中英文分開量度）。回結束 y。"""
    for _al, rs in runs:
        if not rs:
            y += 6 * scale / 72.0 * 72
            continue
        size = max(r[1] for r in rs)
        lh = size * 1.3 * scale
        line, lw = [], 0.0
        maxw = w * 72 * scale
        for text, sz, bold, rgb in rs:
            for ch in text:
                if ch == "\n":                       # <a:br/> 硬換行
                    _flush(dr, line, x, y, maxw, scale, align)
                    y += lh; line, lw = [], 0.0
                    continue
                cjk = L._is_cn(ch)
                f = _font(cjk, bold, max(6, int(round(sz * scale))))
                cw = f.getlength(ch)
                if lw + cw > maxw and line:
                    _flush(dr, line, x, y, maxw, scale, align)
                    y += lh; line, lw = [], 0.0
                line.append((ch, f, rgb)); lw += cw
        if line:
            _flush(dr, line, x, y, maxw, scale, align)
            y += lh
    return y


def _flush(dr, line, x, y, maxw, scale, align):
    lw = sum(f.getlength(c) for c, f, _ in line)
    cx = x + (maxw - lw if align.startswith("RIGHT") else
              (maxw - lw) / 2 if align.startswith("CENTER") else 0)
    for c, f, rgb in line:
        dr.text((cx, y), c, font=f, fill=rgb)
        cx += f.getlength(c)


def preview(path, dpi=110, only=None):
    """自畫每版 PNG（唔使 Office）：rect fill、table（含估算 row 長高）、文字 run。
    ★ 表格用同 layout.row_h 一樣嘅估算 → 睇到嘅就係 audit 檢查嘅嘢。"""
    from PIL import Image, ImageDraw
    prs = Presentation(str(path))
    W, H = _in(prs.slide_width), _in(prs.slide_height)
    out = Path(path).resolve().parent / (Path(path).stem + "_preview")
    out.mkdir(exist_ok=True)
    sc = dpi / 72.0                                  # pt → px
    for i, slide in enumerate(prs.slides, 1):
        if only and i not in only:
            continue
        img = Image.new("RGB", (int(W * dpi), int(H * dpi)), "white")
        dr = ImageDraw.Draw(img)
        for sh in slide.shapes:
            x, y, w, h = _in(sh.left) * dpi, _in(sh.top) * dpi, _in(sh.width) * dpi, _in(sh.height) * dpi
            if sh.has_table:
                tbl = sh.table
                widths = [_in(c.width) for c in tbl.columns]
                cy = y
                for r in tbl.rows:
                    sizes = [run.font.size.pt for cell in r.cells
                             for p in cell.text_frame.paragraphs for run in p.runs if run.font.size]
                    size = max(sizes) if sizes else 10.0
                    txts, ws = [], []
                    for ci, cell in enumerate(r.cells):
                        if getattr(cell, "is_spanned", False):
                            continue
                        n = cell.span_width if getattr(cell, "is_merge_origin", False) else 1
                        txts.append(cell.text); ws.append(sum(widths[ci:ci + n]))
                    rh = max(_in(r.height), L.row_h(txts, ws, size)) * dpi
                    cx = x
                    for ci, cell in enumerate(r.cells):
                        if getattr(cell, "is_spanned", False):
                            continue          # merge origin 已經推進晒 cx，唔可以再加
                        n = cell.span_width if getattr(cell, "is_merge_origin", False) else 1
                        cw = sum(widths[ci:ci + n]) * dpi
                        try:
                            fc = cell.fill.fore_color.rgb
                            fill = (fc[0], fc[1], fc[2])
                        except Exception:
                            fill = (255, 255, 255)
                        dr.rectangle([cx, cy, cx + cw, cy + rh], fill=fill, outline=(191, 191, 191))
                        al = str(cell.text_frame.paragraphs[0].alignment or "LEFT")
                        _draw_text(dr, _runs(cell.text_frame), cx + 2, cy + 2,
                                   (cw - 6) / dpi, sc, align=al)
                        cx += cw
                    cy += rh
                if cy > H * dpi:                     # 表爆版 → 畫紅界
                    dr.line([0, H * dpi - 2, W * dpi, H * dpi - 2], fill=(220, 0, 0), width=3)
                continue
            if sh.shape_type is not None and not sh.has_text_frame:
                continue
            try:
                if sh.fill.type is not None and sh.fill.type == 1:
                    fc = sh.fill.fore_color.rgb
                    dr.rectangle([x, y, x + w, y + h], fill=(fc[0], fc[1], fc[2]))
            except Exception:
                pass
            if sh.has_text_frame and sh.text_frame.text.strip():
                al = str(sh.text_frame.paragraphs[0].alignment or "LEFT")
                _draw_text(dr, _runs(sh.text_frame), x + 1, y + 1, (w - 2) / dpi, sc, align=al)
        dr.rectangle([0, 0, W * dpi - 1, H * dpi - 1], outline=(170, 170, 170))
        img.save(out / f"s-{i:02d}.png")
    print(f"✓ preview PNG → {out}/s-*.png")


def main():
    args = [a for a in sys.argv[1:]]
    if not args:
        print(__doc__); return
    path = args[0]
    if "--render" in args:
        render(path); return
    if "--preview" in args:
        only = None
        if "--slide" in args:
            only = {int(x) for x in args[args.index("--slide") + 1].split(",")}
        preview(path, only=only); return
    if "--slide" in args:
        dump_slide(path, int(args[args.index("--slide") + 1])); return
    audit(path)


if __name__ == "__main__":
    main()
