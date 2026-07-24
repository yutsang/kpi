#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_pdf_content.py — 抽 PDF 逐頁文字（優先 text layer；某頁 text 太少就 tesseract OCR fallback）。
輸出去 report_content/（gitignored）：逐頁 {stem}_pNN.txt + 合併 {stem}.md，方便 paste 返嚟 cross-check。

用法：python scripts/report/extract_pdf_content.py mgm_2025_report.pdf [out_dir]
需要：pdfplumber（text layer）；OCR fallback 需 tesseract + pdftoppm（chi_tra/chi_sim/eng）。
⚠ 輸出同 PDF 都係 confidential — 已 gitignore，切勿 commit。
"""
import subprocess
import sys
import tempfile
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("✗ pip install pdfplumber"); sys.exit(1)

OCR_LANGS = "chi_tra+chi_sim+eng"
CJK_MIN = 0.45       # text layer CJK 比例低過咁多 = 壞編碼亂碼 → 改用 OCR


def _cjk_ratio(s):
    """真中文文字 CJK 佔非空白字元比例應好高；亂碼（壞 CID 編碼）會極低。"""
    if not s:
        return 0.0
    cjk = sum(1 for c in s if "一" <= c <= "鿿")
    nonspace = sum(1 for c in s if not c.isspace())
    return cjk / max(nonspace, 1)


def _ocr_page(pdf_path, page_no, dpi=200):
    """pdftoppm 出 PNG → tesseract OCR（chi_tra+chi_sim+eng）。失敗回空字串。"""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return ""
    with tempfile.TemporaryDirectory() as td:
        base = str(Path(td) / "pg")
        try:
            subprocess.run(["pdftoppm", "-png", "-r", str(dpi), "-f", str(page_no),
                            "-l", str(page_no), pdf_path, base],
                           check=True, capture_output=True)
        except Exception:
            return ""
        pngs = sorted(Path(td).glob("pg*.png"))
        if not pngs:
            return ""
        try:
            return pytesseract.image_to_string(Image.open(pngs[0]), lang=OCR_LANGS).strip()
        except Exception:
            return ""


def main():
    if len(sys.argv) < 2:
        print("用法：python scripts/report/extract_pdf_content.py <pdf> [out_dir]"); return
    pdf_path = sys.argv[1]
    stem = Path(pdf_path).stem
    out_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("report_content")
    out_dir.mkdir(parents=True, exist_ok=True)

    combined = [f"# {stem} — 逐頁內容（text layer CJK 太低則 OCR）\n"]
    n_ocr = 0
    with pdfplumber.open(pdf_path) as pdf:
        n = len(pdf.pages)
        for i, page in enumerate(pdf.pages, 1):
            tl = (page.extract_text() or "").strip()
            src = "text-layer"
            # text layer CJK 比例低 = 壞編碼亂碼 → OCR（本 PDF 正屬此類）
            if _cjk_ratio(tl) < CJK_MIN:
                ocr = _ocr_page(pdf_path, i)
                if _cjk_ratio(ocr) >= _cjk_ratio(tl):
                    txt, src, n_ocr = ocr, "ocr", n_ocr + 1
                else:
                    txt = tl
            else:
                txt = tl
            (out_dir / f"{stem}_p{i:02d}.txt").write_text(txt, encoding="utf-8")
            combined.append(f"\n\n===== 第 {i} 頁 / {n}（{src}，{len(txt)} 字）=====\n\n{txt}")
            print(f"  p{i:02d}/{n}：{src}  {len(txt)} 字  cjk={_cjk_ratio(txt):.2f}")

    (out_dir / f"{stem}.md").write_text("\n".join(combined), encoding="utf-8")
    print(f"\n✓ {n} 頁 → {out_dir}/（{stem}.md 合併 + 逐頁 _pNN.txt）；OCR fallback {n_ocr} 頁")
    print(f"  （confidential，已 gitignore，切勿 commit）")


if __name__ == "__main__":
    main()
