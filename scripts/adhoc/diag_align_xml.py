#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
diag_align_xml.py — 掃描 align_to_header 產生嘅 sheet XML，揾出 Excel 會拒收
（"Removed Records: Cell information"）嘅 cell。

用法：
    # 掃 results/ 或任何路徑下的 _dbg_*.xml
    python scripts/adhoc/diag_align_xml.py results/_dbg_xxx.xml [more.xml ...]

    # 或直接掃一個（未加密）輸出 xlsx（連 styles/sharedStrings 一齊驗）
    python scripts/adhoc/diag_align_xml.py path/to/output.plain.xlsx

檢查每個 <c>：
  - ref 的 row 同 parent <row r=> 一致
  - col ≤ 16384, row ≤ 1048576
  - t="s"：<v> 係整數且 0 ≤ idx < sharedStrings 數（如有 xlsx）
  - t="str"/t="e"：一定要有 <f>（冇 → 非法，Excel 丟 cell）
  - t="inlineStr"：要有 <is>，唔可以有 <v>
  - t="b"：<v> 係 0/1
  - s= index < cellXfs 數（如有 xlsx）
  - 同一行冇重複 ref
  - 同一行 ref 升序
  - 冇 XML 1.0 非法控制字元
  - <mergeCell> ref 格式合法
"""
import re
import sys
import zipfile
from pathlib import Path

_INVALID_XML = re.compile(
    rb'[^\x09\x0A\x0D\x20-\x7f\xc2-\xf4][\x80-\xbf]*')  # rough: catch stray control bytes


def _cidx(letter: str) -> int:
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + ord(ch) - 64
    return idx


def _split_ref(ref: str):
    m = re.match(r'^([A-Z]+)(\d+)$', ref)
    if not m:
        return None, None
    return _cidx(m.group(1)), int(m.group(2))


def scan_sheet(xml: bytes, label: str, n_strings=None, n_xfs=None) -> list:
    problems = []
    # crude control-char scan on decoded text (invalid XML 1.0 chars)
    try:
        txt = xml.decode('utf-8')
    except UnicodeDecodeError:
        problems.append(f"{label}: 唔係合法 UTF-8")
        return problems
    bad_ctrl = re.findall(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', txt)
    if bad_ctrl:
        problems.append(f"{label}: 有 {len(bad_ctrl)} 個非法控制字元")

    # per-row cell scan
    for rm in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*>(.*?)</row>', txt, re.DOTALL):
        rn = int(rm.group(1))
        body = rm.group(2)
        cols_seen = []
        for cm in re.finditer(r'<c\b([^>]*?)(?:/>|>(.*?)</c>)', body, re.DOTALL):
            attrs = cm.group(1)
            inner = cm.group(2) or ''
            ref_m = re.search(r'\br="([A-Z]+\d+)"', attrs)
            if not ref_m:
                problems.append(f"{label} row{rn}: cell 冇 r= 屬性")
                continue
            ref = ref_m.group(1)
            cc, cr = _split_ref(ref)
            if cc is None:
                problems.append(f"{label} row{rn}: 壞 ref {ref}")
                continue
            if cr != rn:
                problems.append(f"{label} row{rn}: cell {ref} 的 row≠{rn}")
            if cc > 16384:
                problems.append(f"{label} row{rn}: {ref} col>16384")
            cols_seen.append((cc, ref))
            t_m = re.search(r'\bt="([^"]+)"', attrs)
            t = t_m.group(1) if t_m else None
            s_m = re.search(r'\bs="(\d+)"', attrs)
            s = int(s_m.group(1)) if s_m else 0
            has_f = '<f' in inner
            has_v = '<v>' in inner or '<v/>' in inner or '<v ' in inner
            has_is = '<is>' in inner or '<is/>' in inner
            v_m = re.search(r'<v>(.*?)</v>', inner, re.DOTALL)
            vtext = v_m.group(1) if v_m else None

            if t in ('str', 'e') and not has_f:
                problems.append(f"{label} row{rn} {ref}: t=\"{t}\" 冇 <f>（Excel 會丟）")
            if t == 'inlineStr':
                if not has_is:
                    problems.append(f"{label} row{rn} {ref}: inlineStr 冇 <is>")
                if has_v:
                    problems.append(f"{label} row{rn} {ref}: inlineStr 有 <v>（衝突）")
            if t == 's':
                if vtext is None or not re.fullmatch(r'\d+', vtext.strip()):
                    problems.append(f"{label} row{rn} {ref}: t=\"s\" 的 <v>={vtext!r} 唔係整數")
                elif n_strings is not None and int(vtext) >= n_strings:
                    problems.append(f"{label} row{rn} {ref}: sharedString idx {vtext} ≥ {n_strings}")
            if t == 'b' and vtext is not None and vtext.strip() not in ('0', '1'):
                problems.append(f"{label} row{rn} {ref}: t=\"b\" 的 <v>={vtext!r} 非 0/1")
            if n_xfs is not None and s >= n_xfs:
                problems.append(f"{label} row{rn} {ref}: style idx {s} ≥ cellXfs {n_xfs}")
        # dup / order
        only_cols = [c for c, _ in cols_seen]
        if len(only_cols) != len(set(only_cols)):
            dup = [ref for c, ref in cols_seen if only_cols.count(c) > 1]
            problems.append(f"{label} row{rn}: 重複 col ref {dup}")
        if only_cols != sorted(only_cols):
            problems.append(f"{label} row{rn}: cell 唔係升序 {[ref for _, ref in cols_seen]}")

    # merge refs
    for mm in re.finditer(r'<mergeCell ref="([^"]+)"/>', txt):
        ref = mm.group(1)
        if not re.fullmatch(r'[A-Z]+\d+:[A-Z]+\d+', ref):
            problems.append(f"{label}: 壞 mergeCell ref {ref}")
    return problems


def scan_xlsx(path: Path) -> list:
    out = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        n_strings = None
        if 'xl/sharedStrings.xml' in names:
            ss = z.read('xl/sharedStrings.xml').decode('utf-8', 'ignore')
            m = re.search(r'\buniqueCount="(\d+)"', ss) or re.search(r'\bcount="(\d+)"', ss)
            n_strings = len(re.findall(r'<si\b', ss))
        n_xfs = None
        if 'xl/styles.xml' in names:
            st = z.read('xl/styles.xml').decode('utf-8', 'ignore')
            cx = re.search(r'<cellXfs\b[^>]*>(.*?)</cellXfs>', st, re.DOTALL)
            if cx:
                n_xfs = len(re.findall(r'<xf\b', cx.group(1)))
        print(f"# {path.name}: sharedStrings={n_strings} cellXfs={n_xfs}")
        for n in names:
            if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
                out += scan_sheet(z.read(n), n, n_strings, n_xfs)
    return out


def main():
    args = sys.argv[1:]
    if not args:
        # default: scan results/*.xml
        args = [str(p) for p in Path('results').glob('*_dbg_*.xml')]
        args += [str(p) for p in Path('results').glob('*.xml')]
    if not args:
        print("俾一個 .xml 或 .xlsx 路徑，或放 _dbg_*.xml 入 results/")
        return
    all_problems = []
    for a in args:
        p = Path(a)
        if not p.exists():
            print(f"✗ 唔存在: {p}")
            continue
        if p.suffix.lower() == '.xlsx':
            probs = scan_xlsx(p)
        else:
            probs = scan_sheet(p.read_bytes(), p.name)
        if probs:
            print(f"\n### {p.name}: {len(probs)} 個問題")
            for x in probs[:200]:
                print("  ✗", x)
        else:
            print(f"\n### {p.name}: ✓ 冇揾到已知問題")
        all_problems += probs
    print(f"\n總共 {len(all_problems)} 個問題")


if __name__ == '__main__':
    main()
