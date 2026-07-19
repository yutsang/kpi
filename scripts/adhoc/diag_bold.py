# -*- coding: utf-8 -*-
"""
Diagnostic: inspect bold formatting in source xlsx vs aligned output.

Usage (Windows):
  python scripts/adhoc/diag_bold.py <source_xlsx> [<aligned_xlsx>]

Prints for each non-empty cell in sheet 1:
  row | col | value[:35] | cell.font.bold | in_rich_lookup | rt_runs_bold
"""
import sys, io, zipfile, xml.etree.ElementTree as ET
from pathlib import Path
from copy import copy

PASSWORD = "dicj_kpmg"
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def decrypt(path: Path) -> io.BytesIO:
    import msoffcrypto
    buf = io.BytesIO()
    with open(path, "rb") as f:
        off = msoffcrypto.OfficeFile(f)
        off.load_key(password=PASSWORD)
        off.decrypt(buf)
    buf.seek(0)
    return buf


def load(path: Path):
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        src = path
    except Exception:
        src = decrypt(path)
        wb = openpyxl.load_workbook(src, data_only=True)
    return wb, src


def build_rich_lookup(src):
    """Returns {plain_text: [(run_text, b_in_rpr)]} for si entries with runs."""
    if isinstance(src, io.BytesIO):
        src.seek(0)
    result = {}
    try:
        with zipfile.ZipFile(src, 'r') as zf:
            if 'xl/sharedStrings.xml' not in zf.namelist():
                return {}
            root = ET.parse(zf.open('xl/sharedStrings.xml')).getroot()
    except Exception as e:
        return {"_error": str(e)}

    for si in root.findall(f'{{{NS}}}si'):
        rs = si.findall(f'{{{NS}}}r')
        if not rs:
            continue
        runs = []
        plain = ""
        for r in rs:
            t_el = r.find(f'{{{NS}}}t')
            txt = (t_el.text or '') if t_el is not None else ''
            plain += txt
            rpr = r.find(f'{{{NS}}}rPr')
            b_el = rpr.find(f'{{{NS}}}b') if rpr is not None else None
            if b_el is not None:
                b_val = b_el.get('val', '1') != '0'
            else:
                b_val = None  # not specified in run
            runs.append((txt[:20], b_val))
        result[plain] = runs
    return result


def inspect(path: Path, label: str):
    print(f"\n{'='*70}")
    print(f"  {label}: {path.name}")
    print(f"{'='*70}")

    wb, src = load(path)
    ws = wb.worksheets[0]
    rich = build_rich_lookup(src)

    print(f"  Sheet: {ws.title} | Rich lookup entries: {len(rich)}")
    print(f"  {'R':>3} {'C':>3}  {'font.bold':>10}  {'in_lookup':>9}  {'run b_vals':20}  value[:40]")
    print(f"  {'-'*3} {'-'*3}  {'-'*10}  {'-'*9}  {'-'*20}  {'-'*40}")

    shown = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            v = str(cell.value)
            fb = cell.font.bold if cell.font else '?'
            in_lk = v in rich
            run_b = ""
            if in_lk:
                runs = rich[v]
                run_b = str([b for _, b in runs])[:20]
            print(f"  {cell.row:>3} {cell.column:>3}  {str(fb):>10}  {str(in_lk):>9}  {run_b:<20}  {v[:40]!r}")
            shown += 1
            if shown >= 60:
                print("  ... (truncated at 60 cells)")
                return

    # summary stats
    all_bold = [c.font.bold for row in ws.iter_rows() for c in row
                if c.value is not None and c.font]
    from collections import Counter
    print(f"\n  font.bold summary: {Counter(all_bold)}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diag_bold.py <source.xlsx> [<aligned.xlsx>]")
        sys.exit(1)

    src_path = Path(sys.argv[1])
    inspect(src_path, "SOURCE")

    if len(sys.argv) >= 3:
        out_path = Path(sys.argv[2])
        inspect(out_path, "ALIGNED OUTPUT")
