#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
align_to_header.py — 對齊 source_1 全部檔到 表頭.xlsx 標準格式（34 欄）。

新架構（v2）：直接 copy source_1 ZIP，ZIP/XML level 改欄序 + 植入計算值。
好處：
  - sharedStrings 原封不動 → rich text / bold 完美保留
  - drawing（Confidential 等 shape）自動繼承，唔需要 inject
  - 唔再需要 template openpyxl workbook 做輸出底

Windows 跑：
    set PYTHONIOENCODING=utf-8
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace ^
        --only "source_1\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx" --preview
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace ^
        --only "source_1\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx" --with-overlay
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace --all --with-overlay
"""
from __future__ import annotations

import argparse
import io
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

PASSWORD = "dicj_kpmg"
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
HDR_SCAN = 12
GATE_NOADJ = True
ENCRYPT_OUT = True
INSERT_XSGZ_FB_COL = True

# Primary OOXML namespace
_XNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_XNS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_XNS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'

# Register common OOXML namespaces so ElementTree preserves them on write
ET.register_namespace('', _XNS)
ET.register_namespace('r', _XNS_R)
ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
ET.register_namespace('xr', 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
ET.register_namespace('xr2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2')
ET.register_namespace('xr3', 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3')
# Additional namespaces that appear in some XLSX files
ET.register_namespace('x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main')
ET.register_namespace('x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main')
ET.register_namespace('x15ac', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac')
ET.register_namespace('xm', 'http://schemas.microsoft.com/office/excel/2006/main')


def _register_all_ns(xml_bytes: bytes) -> None:
    """Scan raw XML bytes for xmlns: declarations and register all prefixes.

    Must be called before ET.parse() so ElementTree preserves original
    namespace prefixes on write (otherwise unknown prefixes become ns0:, ns1:,
    which makes Excel reject the file with a 'found a problem' dialog).
    """
    for prefix, uri in re.findall(rb'xmlns:([A-Za-z][A-Za-z0-9_.-]*)="([^"]+)"', xml_bytes):
        try:
            ET.register_namespace(prefix.decode('ascii', 'ignore'), uri.decode('ascii', 'ignore'))
        except Exception:
            pass
    for uri in re.findall(rb'(?<![A-Za-z:])xmlns="([^"]+)"', xml_bytes):
        try:
            ET.register_namespace('', uri.decode('ascii', 'ignore'))
        except Exception:
            pass


# ── Column index utilities ────────────────────────────────────────────────────
def _cidx(letter: str) -> int:
    """'AB' → 28"""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + ord(ch) - 64
    return idx


def _cletter(n: int) -> str:
    """28 → 'AB'"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_of(ref: str) -> int:
    """'AB5' → 28"""
    m = re.match(r'^([A-Z]+)', ref or '')
    return _cidx(m.group(1)) if m else 0


def _row_of_ref(ref: str) -> int:
    """'AB5' → 5"""
    m = re.search(r'(\d+)$', ref or '')
    return int(m.group(1)) if m else 0


def _col_of_cell(cell_b: bytes) -> int:
    """Parse the actual column index from a <c r="AB5" ...> byte chunk."""
    m = re.search(rb'\br="([A-Z]+)\d+"', cell_b)
    return _cidx(m.group(1).decode('ascii')) if m else 0


def _s_of_cell(cell_b: bytes) -> "str | None":
    """Parse the style index (s="..") from a <c ...> byte chunk."""
    m = re.search(rb'\bs="(\d+)"', cell_b)
    return m.group(1).decode('ascii') if m else None


def _cref(col: int, row: int) -> str:
    return _cletter(col) + str(row)


# ── Byte-level XML primitives (no ET serialization → no namespace mangling) ───

# Chars illegal in XML 1.0 (Excel drops cells containing them)
_INVALID_XML = re.compile(
    '[^\x09\x0A\x0D\x20-퟿-�\U00010000-\U0010FFFF]')

def _xml_esc(text: str) -> str:
    text = _INVALID_XML.sub('', str(text))
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def _xml_esc_attr(val: str) -> str:
    val = _INVALID_XML.sub('', str(val))
    return val.replace('&', '&amp;').replace('"', '&quot;').replace('<', '&lt;')

def _b_inline_c(col: int, row: int, text: str, s: str = '0') -> bytes:
    t = _xml_esc(str(text) if text is not None else '')
    return (f'<c r="{_cref(col,row)}" t="inlineStr" s="{s}">'
            f'<is><t xml:space="preserve">{t}</t></is></c>').encode('utf-8')

def _b_num_c(col: int, row: int, value: float, s: str = '0') -> bytes:
    v = str(int(round(value))) if abs(value - round(value)) < 1e-9 else repr(value)
    return f'<c r="{_cref(col,row)}" s="{s}"><v>{v}</v></c>'.encode('utf-8')

def _b_reref(cell_b: bytes, col: int, row: int) -> bytes:
    """Update r="OLD" → r="NEW" in raw cell bytes."""
    return re.sub(rb'\br="[A-Z]+\d+"', ('r="' + _cref(col, row) + '"').encode(), cell_b, count=1)

def _b_strip_f(cell_b: bytes) -> bytes:
    """Remove <f> formula elements. If the cell was a formula-string (t="str")
    or formula-error (t="e") result, convert it to an inline string so Excel
    keeps the cached value — a bare t="str"/t="e" with no <f> is INVALID and
    Excel drops it ('Removed Records: Cell information')."""
    cell_b = re.sub(rb'<f\b[^>]*/>', b'', cell_b)
    cell_b = re.sub(rb'<f\b[^>]*>.*?</f>', b'', cell_b, flags=re.DOTALL)
    if b't="str"' in cell_b or b't="e"' in cell_b:
        cell_b = cell_b.replace(b't="str"', b't="inlineStr"')
        cell_b = cell_b.replace(b't="e"', b't="inlineStr"')
        cell_b = re.sub(rb'<v>(.*?)</v>',
                        rb'<is><t xml:space="preserve">\1</t></is>',
                        cell_b, flags=re.DOTALL)
        cell_b = re.sub(rb'<v\s*/>', b'<is><t></t></is>', cell_b)
    return cell_b

def _b_row(rn: int, row_el, cells: list) -> bytes:
    """Build <row r="N" safe-attrs>cells</row>. Drops namespaced + spans attrs."""
    parts = []
    if row_el is not None:
        for k, v in row_el.attrib.items():
            if k in ('r', 'spans') or k.startswith('{'):
                continue
            parts.append(f'{k}="{_xml_esc_attr(str(v))}"')
    attrs = (' ' + ' '.join(parts)) if parts else ''
    return (f'<row r="{rn}"{attrs}>').encode('utf-8') + b''.join(cells) + b'</row>'

# Self-closing form MUST come first so `<c .../>` isn't swallowed into the next
# `<c ...>...</c>`. Using `[^>]*?` (non-greedy) so `/>` binds to THIS tag.
_RE_ROW_CHUNK  = re.compile(rb'<row\b[^>]*?/>|<row\b[^>]*?>.*?</row>', re.DOTALL)
_RE_CELL_CHUNK = re.compile(rb'<c\b[^>]*?/>|<c\b[^>]*?>.*?</c>', re.DOTALL)

def _b_extract_rows(sd_bytes: bytes) -> 'dict[int, bytes]':
    """Extract {row_num: raw_bytes} from raw sheetData bytes."""
    result: dict = {}
    for m in _RE_ROW_CHUNK.finditer(sd_bytes):
        r_m = re.search(rb'\br="(\d+)"', m.group())
        if r_m:
            result[int(r_m.group(1))] = m.group()
    return result

def _b_extract_cells(row_b: bytes) -> 'dict[int, bytes]':
    """Extract {col_idx: raw_bytes} from raw row bytes."""
    result: dict = {}
    for m in _RE_CELL_CHUNK.finditer(row_b):
        r_m = re.search(rb'\br="([A-Z]+)\d+"', m.group())
        if r_m:
            result[_cidx(r_m.group(1).decode('ascii'))] = m.group()
    return result

def _b_section(xml_bytes: bytes, tag: bytes) -> 'tuple[int,int] | None':
    """Return (start, end) offsets of <tag>...</tag> or <tag/> in xml_bytes."""
    m = re.search(rb'<' + tag + rb'\b[^>]*>.*?</' + tag + rb'>', xml_bytes, re.DOTALL)
    if m:
        return m.start(), m.end()
    m = re.search(rb'<' + tag + rb'\b[^>]*/>', xml_bytes)
    return (m.start(), m.end()) if m else None

def _b_replace(xml_bytes: bytes, tag: bytes, replacement: bytes) -> bytes:
    span = _b_section(xml_bytes, tag)
    if span:
        return xml_bytes[:span[0]] + replacement + xml_bytes[span[1]:]
    return xml_bytes

def _b_remove(xml_bytes: bytes, tag: bytes) -> bytes:
    return _b_replace(xml_bytes, tag, b'')


def _remap_merge(ref: str, col_map: dict) -> "str | None":
    """'G5:H15' → remapped via col_map; None if either col not in map."""
    parts = ref.split(':')
    if len(parts) != 2:
        return None
    c1, r1 = _col_of(parts[0]), _row_of_ref(parts[0])
    c2, r2 = _col_of(parts[1]), _row_of_ref(parts[1])
    if c1 == c2:
        nc = col_map.get(c1)
        return f'{_cref(nc, r1)}:{_cref(nc, r2)}' if nc else None
    nc1, nc2 = col_map.get(c1), col_map.get(c2)
    return f'{_cref(nc1, r1)}:{_cref(nc2, r2)}' if (nc1 and nc2) else None


# ── 正規化 ────────────────────────────────────────────────────────────────────
_PUNC = str.maketrans({
    "（": "(", "）": ")", "：": ":", "，": ",", "、": ",", "／": "/",
    "　": "", " ": "", " ": "", "\"": "", "“": "", "”": "",
    "‘": "",
})


def _s(x) -> str:
    return "" if x is None else str(x).strip().replace("\r", "")


def nkey(x) -> str:
    return _s(x).replace("\n", "").translate(_PUNC)


_VARIANTS = {
    nkey("畢馬威的分析"): nkey("KPMG分析"),
}
_RE_REPLY = re.compile(r".+的回覆$")
_RE_ASK   = re.compile(r"^KPMG希望進一步向.+瞭解的事項$")


def canon_sub(raw: str) -> str:
    k = nkey(raw)
    k = re.sub(r"\([^()]*\)$", "", k)
    if k in _VARIANTS:
        return _VARIANTS[k]
    if _RE_REPLY.match(k):
        return nkey("跨司工作組的回覆")
    if _RE_ASK.match(k):
        return nkey("KPMG希望進一步向跨司工作組瞭解的事項")
    return k


# ── 表頭目標 schema（34 欄）────────────────────────────────────────────────────
G1   = "與跨司工作組的第一輪問題諮詢"
G2   = "與跨司工作組的第二輪問題諮詢"
GT   = "與承批公司溝通潛在調整事項"

TARGET_SCHEMA = [
    {"grp": "",   "sub": "是否該司局範疇的項目",                              "rule": ("copy_any", "是否該司局範疇的項目")},
    {"grp": G1,   "sub": "是否有希望諮詢的問題",                              "rule": ("copy", G1, "是否有希望諮詢的問題")},
    {"grp": G1,   "sub": "問題狀態",                                          "rule": ("copy", G1, "問題狀態")},
    {"grp": G1,   "sub": "KPMG提出日期",                                      "rule": ("copy", G1, "KPMG提出日期")},
    {"grp": G1,   "sub": "KPMG分析",                                          "rule": ("copy", G1, "KPMG分析")},
    {"grp": G1,   "sub": "承批公司管理層解釋",                                "rule": ("copy", G1, "承批公司管理層解釋")},
    {"grp": G1,   "sub": "KPMG希望進一步向跨司工作組瞭解的事項",             "rule": ("copy", G1, "KPMG希望進一步向跨司工作組瞭解的事項")},
    {"grp": G1,   "sub": "跨司工作組的回覆",                                  "rule": ("copy", G1, "跨司工作組的回覆")},
    {"grp": G2,   "sub": "問題狀態",                                          "rule": ("copy", G2, "問題狀態")},
    {"grp": G2,   "sub": "KPMG提出日期",                                      "rule": ("copy", G2, "KPMG提出日期")},
    {"grp": G2,   "sub": "KPMG分析",                                          "rule": ("copy", G2, "KPMG分析")},
    {"grp": G2,   "sub": "承批公司管理層解釋",                                "rule": ("copy", G2, "承批公司管理層解釋")},
    {"grp": G2,   "sub": "KPMG希望進一步向跨司工作組瞭解的事項",             "rule": ("copy", G2, "KPMG希望進一步向跨司工作組瞭解的事項")},
    {"grp": G2,   "sub": "跨司工作組的回覆",                                  "rule": ("copy", G2, "跨司工作組的回覆")},
    {"grp": GT,   "sub": "畢馬威關注事項",                                    "rule": ("copy", GT, "畢馬威關注事項")},
    {"grp": GT,   "sub": "承批公司的反饋意見",                                "rule": ("copy", GT, "承批公司的反饋意見")},
    {"grp": GT,   "sub": "是否需進一步與跨司工作組溝通",                      "rule": ("yn_followup",)},
    {"grp": GT,   "sub": "需溝通關注事項",                                    "rule": ("blank",)},
    {"grp": GT,   "sub": "該關注事項涉及調整金額",                            "rule": ("abs_total",)},
    {"grp": GT,   "sub": "跨司工作組主責部門針對該關注事項已給的反饋意見",   "rule": ("copy", GT, "跨司工作組的反饋意見")},
    {"grp": GT,   "sub": "KPMG需與跨司工作組進一步確認的問題",               "rule": ("copy", GT, "畢馬威的分析")},
    {"grp": GT,   "sub": "跨司工作組最新反饋意見",                            "rule": ("blank",)},
    {"grp": "畢馬威審查意見",     "sub": "建議調整金額",                      "rule": ("seed", "潛在調整合計")},
    {"grp": "畢馬威審查意見",     "sub": "調整原因",                          "rule": ("enum",)},
    {"grp": "畢馬威審查意見",     "sub": "建議調整後金額",                    "rule": ("seed", "調整後投資金額")},
    {"grp": "跨司工作組審閱意見", "sub": "項目分析意見",                      "rule": ("blank",)},
    {"grp": "跨司工作組審閱意見", "sub": "建議接納之調整後金額",              "rule": ("seed", "跨司工作組確認投資金額")},
    {"grp": "",   "sub": "對比畢馬威審查意見與跨司工作組審閱意見是否一致",   "rule": ("blank",)},
]

ADJ_NON_TYPE = {nkey(x) for x in
                ["申報投資金額", "潛在調整合計", "調整後投資金額", "跨司工作組確認投資金額"]}
OVERLAY_SUBS = [nkey(x) for x in
                ["是否需進一步與跨司工作組溝通", "需溝通關注事項", "該關注事項涉及調整金額",
                 "跨司工作組主責部門針對該關注事項已給的反饋意見",
                 "KPMG需與跨司工作組進一步確認的問題", "跨司工作組最新反饋意見",
                 "畢馬威關注事項", "承批公司的反饋意見"]]
NO_MERGE_SUBS = {nkey("是否需進一步與跨司工作組溝通")}


# ── I/O ──────────────────────────────────────────────────────────────────────
def _is_encrypted(path: Path) -> bool:
    try:
        with open(path, "rb") as f:
            return f.read(4) == b"\xd0\xcf\x11\xe0"
    except Exception:
        return False


def load_wb(path: Path):
    """Load for openpyxl extraction only (data_only; encrypted → decrypt)."""
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception:
        import msoffcrypto
        buf = io.BytesIO()
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=PASSWORD)
            off.decrypt(buf)
        buf.seek(0)
        return openpyxl.load_workbook(buf, data_only=True)


def load_raw(path: Path) -> "tuple[io.BytesIO | Path, bool]":
    """Returns (src, was_encrypted). BytesIO if encrypted, else Path."""
    if _is_encrypted(path):
        import msoffcrypto
        buf = io.BytesIO()
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=PASSWORD)
            off.decrypt(buf)
        buf.seek(0)
        return buf, True
    return path, False


def encrypt_file(plain: Path, enc: Path, log) -> bool:
    import subprocess, sys
    base = [str(plain), str(enc), "-e", "-p", PASSWORD]
    for cmd in ([sys.executable, "-m", "msoffcrypto"] + base,
                ["msoffcrypto-tool"] + base):
        try:
            r = subprocess.run(cmd, capture_output=True, text=True)
        except FileNotFoundError:
            continue
        if r.returncode == 0 and enc.exists() and enc.stat().st_size > 0:
            return True
    log("  ⚠ 加密失敗 → 輸出未加密；pip install -U msoffcrypto-tool")
    return False


# ── 正規化 / 數字 ─────────────────────────────────────────────────────────────
def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _s(v).replace(",", "").replace("，", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s)
    return None


def fmt_amt(x: float) -> str:
    return str(int(x)) if abs(x - round(x)) < 1e-9 else f"{x:.2f}".rstrip("0").rstrip(".")


# ── 表頭解析 ─────────────────────────────────────────────────────────────────
def group_map(ws, grow: int, maxcol: int) -> dict:
    gm: dict = {}
    merged = {}
    for m in ws.merged_cells.ranges:
        if m.min_row <= grow <= m.max_row and m.max_col > m.min_col:
            v = _s(ws.cell(m.min_row, m.min_col).value)
            for c in range(m.min_col, m.max_col + 1):
                merged[c] = v
    for c in range(1, maxcol + 1):
        gm[c] = nkey(merged.get(c, _s(ws.cell(grow, c).value)))
    return gm


def detect_sub_row(ws, maxrow: int, maxcol: int) -> int:
    known = {nkey(r["sub"]) for r in TARGET_SCHEMA} | {nkey("非博彩投資項目性質")}
    best, best_hit = 1, -1
    for r in range(1, min(HDR_SCAN, maxrow) + 1):
        vals = {canon_sub(ws.cell(r, c).value) for c in range(1, maxcol + 1)}
        hit = len(known & vals)
        if hit > best_hit:
            best_hit, best = hit, r
    return best


def find_anchor_col(ws, subrow: int, maxcol: int) -> "int | None":
    tgt = nkey("是否該司局範疇的項目")
    for c in range(1, maxcol + 1):
        if canon_sub(ws.cell(subrow, c).value) == tgt:
            return c
    return None


# ── 項目抽取 ─────────────────────────────────────────────────────────────────
class Project:
    __slots__ = ("r0", "r1", "left", "by_gs", "adj", "seed", "seq", "override")

    def __init__(self, r0, r1):
        self.r0, self.r1 = r0, r1
        self.left: list[list] = []
        self.by_gs: dict = {}
        self.adj: list[tuple] = []
        self.seed: dict = {}
        self.seq: str = ""
        self.override: dict = {}

    def has_adj(self) -> bool:
        if self.adj:
            return True
        t = _adj_total(self)
        return t is not None and abs(t) > 1e-9


def merged_val(ws, c: int, r0: int, r1: int):
    for r in range(r0, r1 + 1):
        v = ws.cell(r, c).value
        if _s(v) != "":
            return v
    return None


def project_spans(ws, anchor: int, data0: int, maxrow: int) -> list:
    spans = []
    covered = set()
    for m in ws.merged_cells.ranges:
        if m.min_col <= anchor <= m.max_col and m.min_row >= data0 and m.max_row > m.min_row:
            spans.append((m.min_row, m.max_row))
            covered.update(range(m.min_row, m.max_row + 1))
    for r in range(data0, maxrow + 1):
        if r in covered:
            continue
        if _s(ws.cell(r, anchor).value) != "":
            spans.append((r, r))
    return sorted(spans)


_RE_ITEM_LABEL    = "投資項目序號及名稱"
_RE_ITEM_LABEL_NK = nkey(_RE_ITEM_LABEL)
_RE_SEQ_CODE      = re.compile(r"^(?:[A-Za-z]{0,5}\d+|項目\d+)\s*[-–]")


def _derive_seq(first: list) -> str:
    for i, x in enumerate(first):
        if nkey(x).startswith(_RE_ITEM_LABEL_NK):
            tail = re.sub(r"^.*?投資項目序號及名稱[:：]?\s*", "", x).strip()
            if tail:
                return tail
            for y in first[i + 1:]:
                if _s(y):
                    return _s(y)
    for x in first:
        if _RE_SEQ_CODE.match(x):
            return x
    return first[1] if len(first) > 1 else (first[0] if first else "")


def find_label_col(ws, subrow: int, anchor: int, maxrow: int) -> "int | None":
    for r in range(subrow + 1, min(subrow + 80, maxrow) + 1):
        for c in range(1, max(anchor, 2)):
            if nkey(ws.cell(r, c).value).startswith(_RE_ITEM_LABEL):
                return c
    return None


def _row_blank(ws, r: int, maxcol: int) -> bool:
    return all(_s(ws.cell(r, c).value) == "" for c in range(1, maxcol + 1))


def _is_band_row(ws, r: int, maxcol: int) -> bool:
    for c in range(1, min(maxcol, 6) + 1):
        k = nkey(ws.cell(r, c).value)
        if k.startswith("非博彩投資項目性質") or k == "資料要求":
            return True
    return False


def spans_by_label(ws, label_col: int, data0: int, maxrow: int, maxcol: int) -> list:
    starts = [r for r in range(data0, maxrow + 1)
              if nkey(ws.cell(r, label_col).value).startswith(_RE_ITEM_LABEL)]
    spans = []
    for i, s in enumerate(starts):
        e = starts[i + 1] - 1 if i + 1 < len(starts) else maxrow
        while e > s and (_row_blank(ws, e, maxcol) or _is_band_row(ws, e, maxcol)):
            e -= 1
        spans.append((s, e))
    return spans


def extract(ws, log) -> tuple:
    maxrow, maxcol = ws.max_row or 0, ws.max_column or 0
    subrow = detect_sub_row(ws, maxrow, maxcol)
    grow   = max(1, subrow - 1)
    gm     = group_map(ws, grow, maxcol)
    anchor = find_anchor_col(ws, subrow, maxcol)
    if anchor is None:
        return [], subrow, 0, gm, maxcol, {}, maxrow
    col_gs: dict = {}
    for c in range(1, maxcol + 1):
        col_gs[c] = (gm.get(c, ""), canon_sub(ws.cell(subrow, c).value))
    left_cols  = list(range(1, anchor))
    right_cols = list(range(anchor, maxcol + 1))
    data0      = subrow + 1
    label_col  = find_label_col(ws, subrow, anchor, maxrow)
    am = [(m.min_row, m.max_row) for m in ws.merged_cells.ranges
          if m.min_col <= anchor <= m.max_col and m.min_row >= data0 and m.max_row > m.min_row]
    if label_col:
        spans = spans_by_label(ws, label_col, data0, maxrow, maxcol)
        taller = sum(1 for (s, e) in spans
                     for (a0, a1) in am if a0 == s and (e - s) > (a1 - a0))
        if taller:
            log(f"      （項目分段用標籤行：{len(spans)} 個；{taller} 個高過 anchor 合併）")
    else:
        spans = project_spans(ws, anchor, data0, maxrow)
        log("      （揾唔到標籤欄 → 用 anchor 合併分段）")
    projs: list = []
    for r0, r1 in spans:
        p = Project(r0, r1)
        for r in range(r0, r1 + 1):
            p.left.append([ws.cell(r, c).value for c in left_cols])
        first  = [_s(x) for x in (p.left[0] if p.left else [])]
        p.seq  = _derive_seq(first)
        for c in right_cols:
            grp, sub = col_gs[c]
            val = merged_val(ws, c, r0, r1)
            if sub in {nkey("潛在調整合計"), nkey("調整後投資金額"),
                       nkey("跨司工作組確認投資金額"), nkey("申報投資金額")}:
                p.seed[sub] = val
            elif _is_adj_type(grp, sub, ws.cell(subrow, c).value):
                n = num(val)
                if n is not None and abs(n) > 1e-9:
                    p.adj.append((_s(ws.cell(subrow, c).value), n))
            else:
                if val is not None:
                    p.by_gs[(grp, sub)] = val
        projs.append(p)
    if label_col:
        label_starts = {s for s, _ in spans}
        orphan = [a0 for (a0, a1) in am if a0 not in label_starts
                  and not any(s <= a0 <= e for (s, e) in spans)]
        log(f"      項目({len(projs)}): {[p.seq for p in projs]}")
        if orphan:
            log(f"      ⚠ anchor 有合併起點但唔喺標籤項目內: rows {orphan}")
    return projs, subrow, anchor, gm, maxcol, col_gs, maxrow


# ── 值計算 ────────────────────────────────────────────────────────────────────
def _is_adj_type(grp: str, sub: str, raw: str) -> bool:
    if "潛在調整" not in grp or "溝通" in grp:
        return False
    return sub not in ADJ_NON_TYPE


def _adj_total(p: Project):
    n = num(p.seed.get(nkey("潛在調整合計")))
    if n is not None:
        return n
    after  = num(p.seed.get(nkey("調整後投資金額")))
    before = num(p.seed.get(nkey("申報投資金額")))
    if after is not None and before is not None:
        return after - before
    return None


def build_enum(p: Project) -> str:
    lines = []
    for i, (typ, amt) in enumerate(p.adj, 1):
        lines.append(f"{i}、{typ}：{fmt_amt(abs(amt))}萬澳門元")
    if not lines:
        t = _adj_total(p)
        if t is not None and abs(t) > 1e-9:
            lines.append(f"1、投資金額調整：{fmt_amt(abs(t))}萬澳門元")
    return "\n".join(lines)


def cell_value(tpl, c: int, p: Project):
    _grp, sub = tpl.col_gs[c]
    if sub in p.override and _s(p.override[sub]) != "":
        return p.override[sub]
    return resolve(tpl.rules[c], p)


def resolve(rule, p: Project):
    kind = rule[0]
    if kind == "blank":
        return None
    if kind == "yn_followup":
        return "否"
    if kind == "copy_any":
        sub = canon_sub(rule[1])
        for (g, s), v in p.by_gs.items():
            if s == sub:
                return v
        return None
    if kind == "copy":
        want_sub = canon_sub(rule[2])
        for (g, s), v in p.by_gs.items():
            if s == want_sub and _grp_match(g, rule[1]):
                return v
        return None
    if kind == "enum":
        if GATE_NOADJ and not p.has_adj():
            return None
        return build_enum(p) or None
    if kind == "abs_total":
        if GATE_NOADJ and not p.has_adj():
            return None
        t = _adj_total(p)
        return abs(t) if t is not None else None
    if kind == "seed":
        return p.seed.get(nkey(rule[1]))
    return None


_RE_ROUND1 = re.compile(r"(第一|首|1)輪")
_RE_ROUND2 = re.compile(r"(第二|2)輪")


def _round_of(s: str) -> str:
    if _RE_ROUND2.search(s):
        return "2"
    if _RE_ROUND1.search(s):
        return "1"
    return ""


def _grp_match(got: str, want: str) -> bool:
    g, w = nkey(got), nkey(want)
    rg, rw = _round_of(g), _round_of(w)
    if rg and rw and rg != rw:
        return False
    return w in g or g in w or g.startswith(w[:6])


# ── Template（schema + 表頭 labels；唔再儲 style）────────────────────────────
class Template:
    def __init__(self, hf: Path, log):
        wb = load_wb(hf)
        ws = wb.active
        self.maxcol = ws.max_column or 0
        self.subrow = detect_sub_row(ws, ws.max_row or 0, self.maxcol)
        self.grow   = max(1, self.subrow - 1)
        self.gm     = group_map(ws, self.grow, self.maxcol)
        self.anchor = find_anchor_col(ws, self.subrow, self.maxcol)
        self.col_gs = {c: (self.gm.get(c, ""), canon_sub(ws.cell(self.subrow, c).value))
                       for c in range(1, self.maxcol + 1)}
        self.rules  = self._attach_rules(ws, log)
        # Header labels for output: (row, tgt_col) → display text
        self.hdr_labels: dict = {}
        for r in range(1, self.subrow + 1):
            for c in range(1, self.maxcol + 1):
                v = ws.cell(r, c).value
                if v is not None:
                    self.hdr_labels[(r, c)] = v
        # Header merges as tuples (min_row, min_col, max_row, max_col)
        self.hdr_merges: list = [
            (m.min_row, m.min_col, m.max_row, m.max_col)
            for m in ws.merged_cells.ranges if m.max_row <= self.subrow
        ]
        wb.close()

    def _attach_rules(self, ws, log) -> dict:
        rules: dict = {}
        for c in range(1, (self.anchor or 1)):
            rules[c] = ("__left__", c)
        for c in range(self.anchor or 1, self.maxcol + 1):
            grp, sub = self.col_gs[c]
            if not sub:
                sub = canon_sub(ws.cell(self.grow, c).value)
            match = None
            for row in TARGET_SCHEMA:
                if canon_sub(row["sub"]) == sub and (row["grp"] == "" or _grp_match(grp, row["grp"])):
                    match = row["rule"]
                    break
            if match is None:
                match = ("blank",)
                log(f"      ⚠ 表頭欄 {get_column_letter(c)} ({grp}|{sub}) 冇對應 rule → 留空")
            rules[c] = match
        return rules

    def insert_blank_col(self, after_sub: str, grp: str, sub: str, log=None) -> None:
        """插一條 output-only 空白欄喺 after_sub 後面。"""
        tgt = canon_sub(after_sub)
        P   = None
        for c in range(self.anchor or 1, self.maxcol + 1):
            if self.col_gs[c][1] == tgt:
                P = c + 1
                break
        if P is None:
            if log:
                log(f"      ⚠ 插欄揾唔到『{after_sub}』→ 唔插新欄")
            return
        old_max = self.maxcol
        for c in range(old_max, P - 1, -1):
            self.col_gs[c + 1] = self.col_gs.pop(c)
            self.rules[c + 1]  = self.rules.pop(c)
            if c in self.gm:
                self.gm[c + 1] = self.gm.pop(c)
        self.col_gs[P] = (nkey(grp), canon_sub(sub))
        self.rules[P]  = ("blank",)
        self.gm[P]     = nkey(grp)
        # Shift hdr_labels
        new_labels: dict = {}
        for (r, c), v in self.hdr_labels.items():
            new_labels[(r, c + 1) if c >= P else (r, c)] = v
        new_labels[(self.subrow, P)] = sub
        self.hdr_labels = new_labels
        # Shift hdr_merges
        new_merges = []
        for (mr1, mc1, mr2, mc2) in self.hdr_merges:
            if mc2 < P:
                new_merges.append((mr1, mc1, mr2, mc2))
            elif mc1 >= P:
                new_merges.append((mr1, mc1 + 1, mr2, mc2 + 1))
            else:
                new_merges.append((mr1, mc1, mr2, mc2 + 1))
        self.hdr_merges = new_merges
        self.maxcol = old_max + 1
        if log:
            log(f"      ＋插空白欄『{sub}』→ 第 {P} 欄（{get_column_letter(P)}）")


# ── 欄對位圖（audit/verify 繼續用）──────────────────────────────────────────
def build_col_plan(tpl: Template, src_anchor: int, src_col_gs: dict, src_maxcol: int) -> dict:
    plan: dict = {}
    for c in range(1, tpl.anchor or 1):
        sc = src_anchor - (tpl.anchor - c)
        plan[c] = ("src", sc) if sc >= 1 else ("blank",)
    for c in range(tpl.anchor or 1, tpl.maxcol + 1):
        rule = tpl.rules[c]
        if rule[0] in ("copy", "copy_any"):
            want_sub = canon_sub(rule[1] if rule[0] == "copy_any" else rule[2])
            want_grp = None if rule[0] == "copy_any" else rule[1]
            sc = next((scol for scol, (g, s) in src_col_gs.items()
                       if s == want_sub and (want_grp is None or _grp_match(g, want_grp))), None)
            plan[c] = ("src", sc) if sc else ("derive", rule)
        else:
            plan[c] = ("derive", rule)
    return plan


def unmapped_source_cols(tpl: Template, src_anchor: int, src_col_gs: dict) -> list:
    referenced = []
    for c in range(tpl.anchor or 1, tpl.maxcol + 1):
        rule = tpl.rules[c]
        if rule[0] == "copy_any":
            referenced.append((None, canon_sub(rule[1])))
        elif rule[0] == "copy":
            referenced.append((rule[1], canon_sub(rule[2])))
    out = []
    for sc, (g, s) in sorted(src_col_gs.items()):
        if sc < src_anchor or not s:
            continue
        if s in ADJ_NON_TYPE or _is_adj_type(g, s, s):
            continue
        if not any((rg is None or _grp_match(g, rg)) and s == rs for rg, rs in referenced):
            out.append((sc, g, s))
    return out


def warn_unmapped_source(tpl: Template, src_anchor: int, src_col_gs: dict, log) -> int:
    cols = unmapped_source_cols(tpl, src_anchor, src_col_gs)
    for sc, g, s in cols:
        log(f"      ⚠ source 欄 {get_column_letter(sc)} ({g}|{s}) 冇對應目標 → data 唔會帶過")
    return len(cols)


# ── XML-level 欄對位 ──────────────────────────────────────────────────────────
def build_xml_col_map(tpl: Template, src_anchor: int, src_col_gs: dict):
    """
    Returns:
      col_map:     {src_col_idx: tgt_col_idx}  for "move" cells
      derive_rules: {tgt_col_idx: rule}         for derived/computed cols
    """
    col_map: dict = {}
    derive_rules: dict = {}

    # Left half: right-align against anchor offset
    for tgt_c in range(1, tpl.anchor or 1):
        src_c = src_anchor - (tpl.anchor - tgt_c)
        if src_c >= 1:
            col_map[src_c] = tgt_c

    # Right half: match by (group, sub)
    for tgt_c in range(tpl.anchor or 1, tpl.maxcol + 1):
        rule = tpl.rules[tgt_c]
        if rule[0] in ("copy", "copy_any"):
            want_sub = canon_sub(rule[1] if rule[0] == "copy_any" else rule[2])
            want_grp = None if rule[0] == "copy_any" else rule[1]
            src_c = next(
                (sc for sc, (g, s) in src_col_gs.items()
                 if s == want_sub and (want_grp is None or _grp_match(g, want_grp))),
                None)
            if src_c:
                col_map[src_c] = tgt_c
            else:
                derive_rules[tgt_c] = rule
        else:
            derive_rules[tgt_c] = rule

    return col_map, derive_rules


def build_hdr_col_map(tpl: Template, src_col_gs: dict) -> dict:
    """
    {tgt_col: src_col} by matching (group, sub) for ALL target columns
    (copy AND derive). Used only to borrow the source's real cell styling
    (fill / border / font) for the header row and for computed cells — so
    derive columns (建議調整金額 etc.) don't fall back to a blank style.
    """
    m: dict = {}
    used: set = set()
    for tgt_c in range(1, tpl.maxcol + 1):
        g, s = tpl.col_gs.get(tgt_c, ("", ""))
        if not s:
            continue
        cand = None
        for sc, (sg, ss) in src_col_gs.items():
            if sc in used or ss != s:
                continue
            if g == "" or _grp_match(sg, g):
                cand = sc
                break
        if cand:
            m[tgt_c] = cand
            used.add(cand)
    return m


def precompute_xml_values(tpl: Template, projs: list, col_map: dict,
                          derive_rules: dict, no_merge_tgt_cols: set) -> dict:
    """
    Returns {(row, tgt_col): value} for all cells to INSERT (not move).
    Includes:
      - derived columns (enum, seed, abs_total, yn_followup, blank)
      - copy columns overridden by source_2
    NO_MERGE_SUBS columns: fill ALL rows in the project span (not just r0).
    """
    result: dict = {}
    rev_col_map = {tgt: src for src, tgt in col_map.items()}

    for p in projs:
        # Derived columns
        for tgt_c, rule in derive_rules.items():
            v = cell_value(tpl, tgt_c, p)
            if v is None:
                continue
            if tgt_c in no_merge_tgt_cols:
                for r in range(p.r0, p.r1 + 1):
                    result[(r, tgt_c)] = v
            else:
                result[(p.r0, tgt_c)] = v

        # Copy columns with source_2 override (override takes precedence over moved cell)
        for tgt_c in range(tpl.anchor or 1, tpl.maxcol + 1):
            if tgt_c in derive_rules:
                continue
            _grp, sub = tpl.col_gs.get(tgt_c, ("", ""))
            if sub in p.override and _s(p.override[sub]) != "":
                v = p.override[sub]
                if tgt_c in no_merge_tgt_cols:
                    for r in range(p.r0, p.r1 + 1):
                        result[(r, tgt_c)] = v
                else:
                    result[(p.r0, tgt_c)] = v

    return result


# ── Core XML sheet transform ──────────────────────────────────────────────────
def transform_sheet_xml(
    sheet_bytes: bytes,
    col_map: dict,
    tgt_total_cols: int,
    src_subrow: int,
    tpl: Template,
    computed: dict,
    proj_spans: list,
    no_merge_tgt_cols: set,
    hdr_col_map: "dict | None" = None,
) -> bytes:
    """
    Reorder columns in the sheet XML.

    KEY DESIGN: ET is used for parsing/reading ONLY. Output is produced by
    patching the original raw bytes (replacing <sheetData>, <mergeCells>,
    <dimension> sections with newly generated bytes). The root <worksheet>
    element — including ALL namespace declarations — is kept verbatim from
    the source, so no namespace prefix mangling can occur.
    """
    NS = _XNS
    _register_all_ns(sheet_bytes)
    root = ET.fromstring(sheet_bytes)          # read-only parse

    sheet_data_el = root.find(f'{{{NS}}}sheetData')
    if sheet_data_el is None:
        return sheet_bytes                      # nothing to transform

    grow, subrow = tpl.grow, tpl.subrow
    rev_map = {tgt: src for src, tgt in col_map.items()}
    hdr_col_map = hdr_col_map or {}
    # Prefer the (group,sub)-based header map (covers derive columns) and fall
    # back to col_map's reverse (copy columns) so every target column can borrow
    # its source styling.
    style_src = {t: hdr_col_map.get(t, rev_map.get(t)) for t in range(1, tgt_total_cols + 1)}

    # Computed values by row
    computed_by_row: dict = {}
    for (r, tgt_c), v in computed.items():
        computed_by_row.setdefault(r, {})[tgt_c] = v

    # ET row index (for attribute access in header rebuilding)
    et_rows: dict = {int(r.get('r', '0')): r
                     for r in sheet_data_el.findall(f'{{{NS}}}row')}

    # Raw row bytes extracted from source sheetData
    sd_span = _b_section(sheet_bytes, b'sheetData')
    sd_raw  = sheet_bytes[sd_span[0]:sd_span[1]] if sd_span else b''
    src_row_b = _b_extract_rows(sd_raw)

    # Default style: first non-zero s= from any data row
    def_s = '0'
    for r_el in et_rows.values():
        for c_el in r_el.findall(f'{{{NS}}}c'):
            s = c_el.get('s')
            if s and s != '0':
                def_s = s
                break
        if def_s != '0':
            break

    all_rnums = sorted(set(et_rows.keys()) | {grow, subrow})
    new_rows: list = []          # list[bytes]

    for rn in all_rnums:
        row_el  = et_rows.get(rn)
        raw_row = src_row_b.get(rn, b'')

        # ── Pre-header rows: verbatim ──────────────────────────────────────
        if rn < grow:
            if raw_row:
                new_rows.append(raw_row)
            continue

        # ── Header rows: rebuild from tpl.hdr_labels ──────────────────────
        if rn in (grow, subrow):
            # Map this header row's source cells by column for style borrowing
            hdr_src_style: dict = {}
            if row_el is not None:
                for c_el in row_el.findall(f'{{{NS}}}c'):
                    hdr_src_style[_col_of(c_el.get('r', ''))] = c_el.get('s', def_s)
            cells: list = []
            for tgt_c in range(1, tgt_total_cols + 1):
                label = tpl.hdr_labels.get((rn, tgt_c))
                if label is None:
                    continue
                # Borrow style from the source header cell of the same (grp,sub)
                src_c  = style_src.get(tgt_c)
                s_attr = hdr_src_style.get(src_c, def_s) if src_c else def_s
                cells.append(_b_inline_c(tgt_c, rn, str(label), s_attr))
            new_rows.append(_b_row(rn, row_el, cells))
            continue

        # ── Data rows ─────────────────────────────────────────────────────
        if not raw_row and row_el is None:
            continue

        src_cells = _b_extract_cells(raw_row)
        # Key by the cell's ACTUAL column parsed from bytes (not the intended
        # tgt_c), so that even if a reref anomaly leaves the original column,
        # two cells never collide on the same column in the output. Also drops
        # anything beyond the target schema width.
        by_col: dict = {}    # actual_col(int) → bytes

        for src_c, cell_b in src_cells.items():
            tgt_c = col_map.get(src_c)
            if tgt_c is None or tgt_c > tgt_total_cols:
                continue
            cell_b = _b_strip_f(cell_b)
            cell_b = _b_reref(cell_b, tgt_c, rn)
            ac = _col_of_cell(cell_b)
            if ac < 1 or ac > tgt_total_cols:
                continue
            by_col[ac] = cell_b

        # Apply computed / override values — these WIN over any moved cell that
        # landed on the same column (source may already contain derive columns).
        for tgt_c, value in computed_by_row.get(rn, {}).items():
            if tgt_c < 1 or tgt_c > tgt_total_cols:
                continue
            # Keep the border/fill: reuse the style of the cell already there,
            # else the source cell of the same (grp,sub) in this row, else def_s.
            cs = _s_of_cell(by_col.get(tgt_c, b''))
            if cs is None:
                src_c = style_src.get(tgt_c)
                if src_c is not None and src_c in src_cells:
                    cs = _s_of_cell(src_cells[src_c])
            cs = cs or def_s
            if isinstance(value, (int, float)):
                by_col[tgt_c] = _b_num_c(tgt_c, rn, float(value), cs)
            elif value is not None:
                by_col[tgt_c] = _b_inline_c(tgt_c, rn, str(value), cs)
            else:
                by_col.pop(tgt_c, None)

        if not by_col:
            continue

        ordered = [by_col[c] for c in sorted(by_col)]
        new_rows.append(_b_row(rn, row_el, ordered))

    # ── Build new <sheetData> bytes ────────────────────────────────────────────
    new_sd = b'<sheetData>' + b''.join(new_rows) + b'</sheetData>'

    # ── Build new <mergeCells> bytes ───────────────────────────────────────────
    merges: list = []
    merges_el = root.find(f'{{{NS}}}mergeCells')
    if merges_el is not None:
        for mc in merges_el.findall(f'{{{NS}}}mergeCell'):
            ref   = mc.get('ref', '')
            parts = ref.split(':')
            if len(parts) != 2:
                continue
            r1 = _row_of_ref(parts[0]); r2 = _row_of_ref(parts[1])
            if r1 < grow and r2 < grow:
                merges.append(ref)                  # pre-grow: keep as-is
            elif r1 in (grow, subrow) or r2 in (grow, subrow):
                pass                                # header: rebuilt below
            else:
                new_ref = _remap_merge(ref, col_map)
                if new_ref:
                    # Only keep left-half merges (both endpoints < anchor).
                    # Right-half merges are rebuilt from scratch by project-span
                    # merges below; keeping them here would create duplicates →
                    # Excel removes both and reports "Removed Records: Merge cells".
                    nc1 = _col_of(new_ref.split(':')[0])
                    nc2 = _col_of(new_ref.split(':')[1])
                    if nc1 < (tpl.anchor or 1) and nc2 < (tpl.anchor or 1):
                        merges.append(new_ref)

    # Header merges from template schema
    for (mr1, mc1, mr2, mc2) in tpl.hdr_merges:
        if mc2 > mc1 or mr2 > mr1:
            merges.append(f'{_cref(mc1, mr1)}:{_cref(mc2, mr2)}')

    # Vertical project-span merges for right-half columns
    for r0, r1 in proj_spans:
        if r1 <= r0:
            continue
        for tgt_c in range(tpl.anchor or 1, tgt_total_cols + 1):
            if tgt_c in no_merge_tgt_cols:
                continue
            merges.append(f'{_cref(tgt_c, r0)}:{_cref(tgt_c, r1)}')

    new_mc = b''
    if merges:
        mc_inner = ''.join(f'<mergeCell ref="{r}"/>' for r in merges)
        new_mc = f'<mergeCells count="{len(merges)}">{mc_inner}</mergeCells>'.encode('utf-8')

    # ── Dimension ref ─────────────────────────────────────────────────────────
    max_row_out = max(all_rnums) if all_rnums else 1
    new_dim_ref = f'A1:{_cref(tgt_total_cols, max_row_out)}'.encode()

    # ── Patch original bytes (root element stays verbatim) ────────────────────
    out = sheet_bytes
    # dimension (usually self-closing, before sheetData)
    out = re.sub(rb'(<dimension\b[^>]+ref=")[^"]*(")', rb'\g<1>' + new_dim_ref + rb'\2', out)
    # sheetData
    out = _b_replace(out, b'sheetData', new_sd)
    # mergeCells
    if new_mc:
        if _b_section(out, b'mergeCells'):
            out = _b_replace(out, b'mergeCells', new_mc)
        else:
            out = out.replace(b'</sheetData>', b'</sheetData>' + new_mc, 1)
    else:
        out = _b_remove(out, b'mergeCells')
    # autoFilter (col refs invalid after reorder)
    out = _b_remove(out, b'autoFilter')
    # sheetProtection (allow editing aligned output)
    out = _b_remove(out, b'sheetProtection')

    return out


# ── ZIP-level transform ───────────────────────────────────────────────────────
def transform_xlsx_zip(
    src: "io.BytesIO | Path",
    sheets_info: dict,
    tpl: Template,
    out_path: Path,
    log,
) -> None:
    """
    src:         decrypted source (BytesIO or plain Path)
    sheets_info: {sheet_name: (col_map, computed, proj_spans, no_merge_tgt_cols, src_subrow, hdr_col_map)}
    Writes output xlsx to out_path.
    """
    if isinstance(src, io.BytesIO):
        src.seek(0)

    # Read all ZIP entries
    with zipfile.ZipFile(src, 'r') as zin:
        files: dict = {n: zin.read(n) for n in zin.namelist()}

    # Map sheet names → XML paths via workbook.xml + rels
    wb_xml   = ET.fromstring(files['xl/workbook.xml'])
    rels_raw = files.get('xl/_rels/workbook.xml.rels', b'<Relationships/>')
    rels_xml = ET.fromstring(rels_raw)

    rID_to_path: dict = {}
    for rel in rels_xml.findall(f'{{{_XNS_PKG}}}Relationship'):
        rID    = rel.get('Id', '')
        target = rel.get('Target', '')
        # Target is relative to xl/; normalise to full ZIP path
        if target.startswith('/'):
            path = target.lstrip('/')
        else:
            path = 'xl/' + target
        rID_to_path[rID] = path

    name_to_rID: dict = {}
    for sh in wb_xml.findall(f'.//{{{_XNS}}}sheet'):
        n   = sh.get('name', '')
        rID = sh.get(f'{{{_XNS_R}}}id', '')
        name_to_rID[n] = rID

    # Transform each main sheet
    for sn, (col_map, computed, proj_spans, no_merge, src_subrow, hdr_col_map) in sheets_info.items():
        rID        = name_to_rID.get(sn, '')
        sheet_path = rID_to_path.get(rID, '')
        if not sheet_path or sheet_path not in files:
            log(f"  ⚠ sheet {sn!r}: XML path {sheet_path!r} 唔喺 ZIP → 跳過")
            continue
        log(f"  ── sheet {sn!r}: XML-level 對齊（{len(proj_spans)} 個項目）")
        files[sheet_path] = transform_sheet_xml(
            files[sheet_path],
            col_map=col_map,
            tgt_total_cols=tpl.maxcol,
            src_subrow=src_subrow,
            tpl=tpl,
            computed=computed,
            proj_spans=proj_spans,
            no_merge_tgt_cols=no_merge,
            hdr_col_map=hdr_col_map,
        )
        if os.environ.get("KPI_DUMP_XML"):
            safe = re.sub(r'[^\w]', '_', sn)
            dbg = out_path.parent / f"_dbg_{out_path.stem}_{safe}.xml"
            dbg.write_bytes(files[sheet_path])
            log(f"     [KPI_DUMP_XML] 寫 {dbg.name}")

    # Remove calcChain.xml (formula refs invalid after column reorder).
    # Also remove its entries from [Content_Types].xml and workbook rels,
    # otherwise Excel reports "found a problem" because it expects the file.
    if 'xl/calcChain.xml' in files:
        files.pop('xl/calcChain.xml')
        ct_key = '[Content_Types].xml'
        if ct_key in files:
            files[ct_key] = re.sub(
                rb'<Override[^>]*calcChain[^>]*/>\s*', b'', files[ct_key])
        rels_key = 'xl/_rels/workbook.xml.rels'
        if rels_key in files:
            files[rels_key] = re.sub(
                rb'<Relationship[^>]*calcChain[^>]*/>\s*', b'', files[rels_key])

    # Write output
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(out_path, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)


# ── source_2 overlay ──────────────────────────────────────────────────────────
def find_overlay_file(root: Path, scope: str, company: str) -> "Path | None":
    base = root / "source_2"
    if not base.exists():
        return None
    cands = []
    for p in base.rglob("*"):
        if p.is_dir() or p.name.startswith("~$") or p.suffix.lower() not in EXCEL_EXT:
            continue
        rel = p.relative_to(base).as_posix()
        if scope and scope in rel and company and company in rel:
            cands.append(p)
    return sorted(cands, key=lambda x: len(x.as_posix()))[0] if cands else None


def _concern_sum(ws, p: Project, col_gs: dict) -> "float | None":
    col_amt = next((c for c, (g, s) in col_gs.items()
                    if s == nkey("該關注事項涉及調整金額")), None)
    if col_amt is None:
        return None
    total, prev, found = 0.0, None, False
    for r in range(p.r0, p.r1 + 1):
        a = num(ws.cell(r, col_amt).value)
        if a is None or a == prev:
            continue
        prev = a
        total += abs(a)
        found = True
    return total if found else None


def _seqkey(seq: str) -> str:
    s = re.sub(r"^.*?投資項目序號及名稱[:：]?\s*", "", _s(seq)).strip()
    m = re.match(r"^([A-Za-z]{0,5}\d+|項目\d+)\s*[-–]", s)
    if m:
        return nkey(m.group(1))
    m2 = re.match(r"^\s*(\d+)", s)
    return m2.group(1) if m2 else nkey(s)


def build_overlay(path: Path, log) -> dict:
    try:
        wb = load_wb(path)
    except Exception as e:
        log(f"      ⚠ overlay 開唔到 {path.name}: {e}")
        return {}
    out: dict = {}
    log(f"      overlay 檔 sheets: {wb.sheetnames}")
    for sn in wb.sheetnames:
        ws = wb[sn]
        projs, subrow, anchor, gm, maxcol, col_gs, maxrow = extract(ws, log)
        if not anchor or not projs:
            log(f"      overlay sheet {sn!r}: anchor={anchor} → 跳過")
            continue
        found = 0
        for p in projs:
            d: dict = {}
            for (g, s), v in p.by_gs.items():
                if s not in OVERLAY_SUBS or _s(v) == "":
                    continue
                if s == nkey("該關注事項涉及調整金額"):
                    continue
                d[s] = v
            cs = _concern_sum(ws, p, col_gs)
            if cs is not None:
                d[nkey("該關注事項涉及調整金額")] = cs
            if d:
                out[_seqkey(p.seq)] = d
                found += 1
        log(f"      overlay sheet {sn!r}: {len(projs)} 項目, {found} 個有覆蓋值")
    wb.close()
    return out


def apply_overlay(projs: list, overlay: dict, tpl: Template, log) -> None:
    hits = 0
    for p in projs:
        d = overlay.get(_seqkey(p.seq))
        if not d:
            continue
        hits += 1
        for s, v in d.items():
            p.override[s] = v
    log(f"      overlay 命中 {hits}/{len(projs)} 個項目")


# ── Preview ───────────────────────────────────────────────────────────────────
def preview_sheet(sn, projs: list, tpl: Template, log) -> None:
    log(f"\n  ── sheet {sn!r}：{len(projs)} 個項目 ──")
    for p in projs[:6]:
        log(f"\n    ▸ 項目「{p.seq}」 rows {p.r0}-{p.r1}（{p.r1-p.r0+1}行）")
        log(f"       左半首行: {[_s(x) for x in (p.left[0] if p.left else [])][:6]}")
        if p.adj:
            log(f"       調整類型: " + "; ".join(f"{t}={fmt_amt(a)}" for t, a in p.adj))
        t = _adj_total(p)
        log(f"       seed: 申報={_s(p.seed.get(nkey('申報投資金額')))} "
            f"合計={_s(p.seed.get(nkey('潛在調整合計')))} "
            f"調整後={_s(p.seed.get(nkey('調整後投資金額')))} "
            f"→_adj_total={fmt_amt(t) if t is not None else 'None'}")
        for c in range(tpl.anchor or 1, tpl.maxcol + 1):
            g, s = tpl.col_gs[c]
            if s in {nkey("調整原因"), nkey("該關注事項涉及調整金額"),
                     nkey("建議調整金額"), nkey("建議調整後金額"),
                     nkey("建議接納之調整後金額"), nkey("需溝通關注事項"),
                     nkey("是否需進一步與跨司工作組溝通"), nkey("畢馬威關注事項"),
                     nkey("承批公司的反饋意見"),
                     nkey("跨司工作組主責部門針對該關注事項已給的反饋意見"),
                     nkey("KPMG需與跨司工作組進一步確認的問題"),
                     nkey("跨司工作組最新反饋意見")}:
                v = cell_value(tpl, c, p)
                if _s(v):
                    tag  = " [source_2]" if s in p.override and _s(p.override[s]) != "" else ""
                    disp = _s(v).replace("\n", " ⏎ ")
                    log(f"       → [{get_column_letter(c)} {s}]{tag} = {disp[:120]}")
    if len(projs) > 6:
        log(f"    …（另 {len(projs)-6} 個項目略）")


# ── 每檔處理 ─────────────────────────────────────────────────────────────────
SCOPE_HINTS = ["旅遊局", "博監局", "經濟局", "文化局", "體育局", "郵電局", "其他範疇", "治安警"]


def infer_scope_company(rel: str) -> tuple:
    parts  = Path(rel).parts
    scope  = next((s for s in SCOPE_HINTS if any(s in pp for pp in parts)), "")
    m      = re.search(r"([A-Za-z]{2,})[-\-]", Path(rel).name)
    company = m.group(1) if m else ""
    return scope, company


def is_attachment_sheet(name: str) -> bool:
    return any(k in name for k in ("附件", "承批公司附件", "attach", "Attach"))


def is_junk_sheet(name: str) -> bool:
    return name.strip().upper().startswith("UPSLIDE")


def is_dup_file(rel: str) -> bool:
    stem = Path(rel).stem
    return any(k in stem for k in (" - Copy", "-Copy", "- Copy", "副本", "複本"))


def process_file(root: Path, rel: str, tpl: Template, out_dir: Path,
                 preview: bool, with_overlay: bool, log) -> None:
    src = root / rel
    log(f"\n{'='*74}\n# {rel}")
    if not src.exists():
        log("  ✗ 唔存在"); return

    # Load with openpyxl for extraction
    wb = load_wb(src)
    scope, company = infer_scope_company(rel)

    overlay: dict = {}
    if with_overlay:
        ofile = find_overlay_file(root, scope, company)
        if ofile:
            log(f"  overlay 檔: {ofile.relative_to(root).as_posix()}")
            overlay = build_overlay(ofile, log)
            if not overlay:
                log("  ⚠ overlay 抽唔到任何覆蓋值")
        else:
            log(f"  （冇 source_2 match scope={scope} company={company}）")

    sheets_info: dict = {}   # sheet_name → (col_map, computed, proj_spans, no_merge, subrow, hdr_col_map)

    for sn in wb.sheetnames:
        ws = wb[sn]
        if is_junk_sheet(sn):
            log(f"  ── sheet {sn!r}：垃圾 sheet → 保留原樣（ZIP 直抄）")
            continue
        if is_attachment_sheet(sn):
            log(f"  ── sheet {sn!r}：附件 → 保留原樣（ZIP 直抄）")
            continue
        projs, subrow, anchor, gm, maxcol, col_gs, maxrow = extract(ws, log)
        if not anchor or not projs:
            log(f"  ── sheet {sn!r}：揾唔到 anchor 或冇項目 → 保留原樣")
            continue
        warn_unmapped_source(tpl, anchor, col_gs, log)
        if overlay:
            apply_overlay(projs, overlay, tpl, log)
        if preview:
            preview_sheet(sn, projs, tpl, log)
            continue
        col_map, derive_rules = build_xml_col_map(tpl, anchor, col_gs)
        hdr_col_map = build_hdr_col_map(tpl, col_gs)
        no_merge = {c for c in range(tpl.anchor or 1, tpl.maxcol + 1)
                    if tpl.col_gs.get(c, ("", ""))[1] in NO_MERGE_SUBS}
        computed = precompute_xml_values(tpl, projs, col_map, derive_rules, no_merge)
        proj_spans = [(p.r0, p.r1) for p in projs]
        sheets_info[sn] = (col_map, computed, proj_spans, no_merge, subrow, hdr_col_map)

    wb.close()

    if preview:
        return
    if not sheets_info:
        log("  ⚠ 冇 sheet 需要對齊，skip")
        return

    # Load raw bytes for ZIP-level transform
    raw_src, _was_enc = load_raw(src)
    out_path = out_dir / rel
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if ENCRYPT_OUT:
        tmp = out_path.with_suffix(".plain.xlsx")
        transform_xlsx_zip(raw_src, sheets_info, tpl, tmp, log)
        if encrypt_file(tmp, out_path, log):
            tmp.unlink(missing_ok=True)
            log(f"  ✓ 寫入 {out_path.relative_to(root).as_posix()}（已加密 dicj_kpmg）")
        else:
            shutil.move(str(tmp), str(out_path))
            log(f"  ✓ 寫入 {out_path.relative_to(root).as_posix()}（未加密）")
    else:
        transform_xlsx_zip(raw_src, sheets_info, tpl, out_path, log)
        log(f"  ✓ 寫入 {out_path.relative_to(root).as_posix()}")


def copy_asis(root: Path, rel: str, out_dir: Path, log) -> None:
    src, dst = root / rel, out_dir / rel
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    log(f"  ⧉ 照抄 {rel}")


def iter_source1(root: Path):
    base = root / "source_1"
    for p in sorted(base.rglob("*")):
        if p.is_dir() or p.name.startswith("~$"):
            continue
        if p.suffix.lower() in EXCEL_EXT:
            yield p.relative_to(root).as_posix()


def _quiet(*_a, **_k):
    pass


# ── Audit ────────────────────────────────────────────────────────────────────
TOL_ADJ = 0.5


def audit(root: Path, tpl: Template, log) -> None:
    files = list(iter_source1(root))
    fA, fB, fC, fE, fAtt, ok = [], [], [], [], [], []
    tpl_left = [tpl.col_gs[c][1] for c in range(1, tpl.anchor or 1)]
    log(f"# AUDIT（read-only）：source_1 共 {len(files)} 檔  "
        f"/ template 左半 {(tpl.anchor or 1) - 1} 欄, anchor {get_column_letter(tpl.anchor) if tpl.anchor else '?'}")
    log(f"# template 左半名: {tpl_left}\n")
    for rel in files:
        if any(pp.lower() == "ss" for pp in Path(rel).parts):
            log(f"▸ {rel}   [ss/ → 照抄]")
            continue
        if is_dup_file(rel):
            log(f"▸ {rel}   [副本 → 跳過]")
            continue
        src = root / rel
        enc = "加密" if _is_encrypted(src) else "無密碼"
        try:
            wb = load_wb(src)
        except Exception as e:
            log(f"▸ {rel}   ✗ 開唔到: {e}")
            fA.append(f"{rel} (開唔到)")
            continue
        log(f"▸ {rel}   [{enc}]  sheets={wb.sheetnames}")
        scope, company = infer_scope_company(rel)
        for sn in wb.sheetnames:
            ws = wb[sn]
            if is_junk_sheet(sn):
                log(f"    · {sn!r}  垃圾 → 跳過")
                continue
            if is_attachment_sheet(sn):
                log(f"    · {sn!r}  附件 → best-effort")
                fAtt.append(f"{rel}::{sn}")
                continue
            projs, subrow, anchor, gm, maxcol, col_gs, maxrow = extract(ws, _quiet)
            older = any(y in sn for y in ("2024", "2023", "2022"))
            if not anchor or not projs:
                log(f"    · {sn!r}  ✗ 冇 anchor／冇項目{'（舊年份）' if older else ''}")
                fA.append(f"{rel}::{sn}")
                continue
            leftn  = anchor - 1
            tleft  = (tpl.anchor or 1) - 1
            offset = (tpl.anchor or 1) - anchor
            unm    = unmapped_source_cols(tpl, anchor, col_gs)
            adjn   = sum(1 for _sc, (g, s) in col_gs.items() if _is_adj_type(g, s, s))
            flags  = []
            if leftn != tleft:
                flags.append(f"左半{leftn}欄≠template{tleft}(E)"); fE.append(f"{rel}::{sn}")
            if unm:
                flags.append(f"漏欄{len(unm)}(B)"); fB.append(f"{rel}::{sn}")
            if older:
                flags.append("舊年份(C)"); fC.append(f"{rel}::{sn}")
            status = "  ".join(flags) if flags else "✓ 全對齊"
            log(f"    · {sn!r}  anchor={get_column_letter(anchor)}  左半{leftn}欄(offset{offset:+d})"
                f"  項目{len(projs)}  調整類型欄{adjn}  →  {status}")
            if leftn != tleft:
                src_left = [col_gs[c][1] for c in range(1, anchor)]
                log(f"        左半(source): {src_left}")
            for sc, g, s in unm:
                log(f"        ⚠ 漏: 欄{get_column_letter(sc)} ({g} | {s})")
            if not flags:
                ok.append(f"{rel}::{sn}")
        of = find_overlay_file(root, scope, company)
        log(f"    overlay: {('✓ ' + of.relative_to(root).as_posix()) if of else '✗ 冇對應檔'}")
        wb.close()
    log("\n" + "=" * 64 + "\n# A–F 風險匯總")
    log(f"A 對齊唔到    : {len(fA)}")
    for x in fA: log(f"    - {x}")
    log(f"B 漏欄        : {len(fB)}")
    for x in fB: log(f"    - {x}")
    log(f"C 舊年份      : {len(fC)}")
    log(f"E 左半欄數差  : {len(fE)}")
    log(f"F 附件        : {len(fAtt)}")
    log(f"✓ 全對齊      : {len(ok)}")


# ── Verify ────────────────────────────────────────────────────────────────────
def _project_number_checks(p: Project, tol: float) -> tuple:
    idi, mpi = [], []
    total  = num(p.seed.get(nkey("潛在調整合計")))
    before = num(p.seed.get(nkey("申報投資金額")))
    after  = num(p.seed.get(nkey("調整後投資金額")))
    sigma  = sum(a for _t, a in p.adj) if p.adj else None
    if total is not None and sigma is not None and abs(total - sigma) > tol:
        idi.append(f"潛在調整合計={fmt_amt(total)} ≠ Σ逐項={fmt_amt(sigma)}")
    adj_id = total if total is not None else sigma
    if after is not None and before is not None and adj_id is not None:
        exp = before + adj_id
        if abs(after - exp) > tol:
            idi.append(f"調整後={fmt_amt(after)} ≠ 申報{fmt_amt(before)}+調整{fmt_amt(adj_id)}")
    g_adj = num(resolve(("seed", "潛在調整合計"), p))
    if g_adj is not None and total is not None and abs(g_adj - total) > tol:
        mpi.append(f"建議調整金額 寫={fmt_amt(g_adj)} ≠ 潛在調整合計={fmt_amt(total)}")
    g_after = num(resolve(("seed", "調整後投資金額"), p))
    if g_after is not None and after is not None and abs(g_after - after) > tol:
        mpi.append(f"建議調整後金額 寫={fmt_amt(g_after)} ≠ 調整後={fmt_amt(after)}")
    g_abs   = resolve(("abs_total",), p)
    tot_abs = _adj_total(p)
    if g_abs is not None and tot_abs is not None and abs(g_abs - abs(tot_abs)) > tol:
        mpi.append(f"該關注事項涉及調整金額 寫={fmt_amt(g_abs)} ≠ |調整總額|={fmt_amt(abs(tot_abs))}")
    enum = build_enum(p)
    if enum:
        esum = sum(float(x) for x in re.findall(r"[:：]\s*(-?\d+(?:\.\d+)?)\s*萬", enum))
        exp_e = (sum(abs(a) for _t, a in p.adj) if p.adj
                 else (abs(tot_abs) if tot_abs is not None else None))
        if exp_e is not None and abs(esum - exp_e) > tol:
            mpi.append(f"調整原因列舉加總={fmt_amt(esum)} ≠ 應有={fmt_amt(exp_e)}")
    return idi, mpi


def verify(root: Path, tpl: Template, tol: float, log) -> None:
    files = list(iter_source1(root))
    n_proj = n_idfail = n_mpfail = 0
    G_before = G_after = G_total = 0.0
    log(f"# VERIFY（read-only）：source_1 共 {len(files)} 檔  容差 {fmt_amt(tol)} 萬\n")
    for rel in files:
        if any(pp.lower() == "ss" for pp in Path(rel).parts) or is_dup_file(rel):
            continue
        try:
            wb = load_wb(root / rel)
        except Exception as e:
            log(f"▸ {rel}   ✗ 開唔到: {e}")
            continue
        printed = False
        for sn in wb.sheetnames:
            ws = wb[sn]
            if is_junk_sheet(sn) or is_attachment_sheet(sn):
                continue
            projs, subrow, anchor, gm, maxcol, col_gs, maxrow = extract(ws, _quiet)
            if not anchor or not projs:
                continue
            s_before = s_after = s_total = 0.0
            fails = []
            for p in projs:
                n_proj += 1
                idi, mpi = _project_number_checks(p, tol)
                if idi:
                    n_idfail += 1
                if mpi:
                    n_mpfail += 1
                if idi or mpi:
                    fails.append((p, idi, mpi))
                s_before += num(p.seed.get(nkey("申報投資金額"))) or 0
                s_after  += num(p.seed.get(nkey("調整後投資金額"))) or 0
                s_total  += num(p.seed.get(nkey("潛在調整合計"))) or 0
            G_before += s_before; G_after += s_after; G_total += s_total
            if not printed:
                log(f"▸ {rel}"); printed = True
            log(f"    · {sn!r}  項目{len(projs)}  Σ申報={fmt_amt(s_before)} "
                f"Σ調整後={fmt_amt(s_after)} Σ合計={fmt_amt(s_total)}"
                + (f"  ⚠ 對唔到 {len(fails)} 個" if fails else "  ✓ 全對得返"))
            for p, idi, mpi in fails[:25]:
                for msg in idi:
                    log(f"        ✗[識別] 「{p.seq}」: {msg}")
                for msg in mpi:
                    log(f"        ✗[映射] 「{p.seq}」: {msg}")
        wb.close()
    log("\n" + "=" * 64 + "\n# 對數匯總")
    log(f"項目總數              : {n_proj}")
    log(f"識別/算式對唔返       : {n_idfail}")
    log(f"我哋映射寫錯數        : {n_mpfail}")
    log(f"Σ申報投資金額        : {fmt_amt(G_before)} 萬")
    log(f"Σ調整後投資金額      : {fmt_amt(G_after)} 萬")
    log(f"Σ潛在調整合計        : {fmt_amt(G_total)} 萬")
    log(f"Σ調整後-Σ申報        : {fmt_amt(G_after - G_before)} 萬"
        f"  (應≈Σ合計{fmt_amt(G_total)}，差{fmt_amt(G_after - G_before - G_total)})")


# ── main ─────────────────────────────────────────────────────────────────────
def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root",        default="ad-hoc/workspace")
    ap.add_argument("--header-file", default="表頭.xlsx")
    ap.add_argument("--only",        help="淨處理呢一個相對 root 嘅檔")
    ap.add_argument("--all",         action="store_true", help="批 source_1 全部檔")
    ap.add_argument("--audit",       action="store_true", help="read-only 巡查（唔寫檔）")
    ap.add_argument("--verify",      action="store_true", help="read-only 對數（唔寫檔）")
    ap.add_argument("--tol",         type=float, default=TOL_ADJ)
    ap.add_argument("--preview",     action="store_true", help="唔寫 xlsx，只吐文字")
    ap.add_argument("--with-overlay",action="store_true", help="套 source_2 per-範疇覆蓋")
    ap.add_argument("--fill-zero-adj",action="store_true")
    ap.add_argument("--no-encrypt",  action="store_true")
    ap.add_argument("--no-extra-col",action="store_true")
    ap.add_argument("--out",         default="_aligned")
    a = ap.parse_args()
    global GATE_NOADJ, ENCRYPT_OUT
    GATE_NOADJ  = not a.fill_zero_adj
    ENCRYPT_OUT = not a.no_encrypt

    root    = Path(a.root)
    out_dir = root / a.out
    lines: list = []

    def log(s=""):
        lines.append(s); print(s)

    hf = root / a.header_file
    if not hf.exists():
        raise SystemExit(f"✗ 冇表頭: {hf}")
    tpl = Template(hf, log)
    log(f"# 表頭 template：{tpl.maxcol} 欄, 子表頭 row{tpl.subrow}, "
        f"anchor {get_column_letter(tpl.anchor) if tpl.anchor else '?'}")

    if a.audit:
        audit(root, tpl, log)
        rp = root / "_align_audit.txt"
        rp.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n✓ audit 寫入 {rp.resolve()}")
        return

    if a.verify:
        verify(root, tpl, a.tol, log)
        rp = root / "_align_verify.txt"
        rp.write_text("\n".join(lines), encoding="utf-8")
        print(f"\n✓ verify 寫入 {rp.resolve()}")
        return

    if INSERT_XSGZ_FB_COL and not a.no_extra_col:
        tpl.insert_blank_col("承批公司的反饋意見", GT, "跨司工作組的反饋意見", log)

    if a.only:
        targets = [a.only]
    elif a.all:
        targets = list(iter_source1(root))
    else:
        raise SystemExit("俾 --only <rel> 或 --all")

    for rel in targets:
        parts = Path(rel).parts
        if any(pp.lower() == "ss" for pp in parts):
            if not a.preview:
                copy_asis(root, rel, out_dir, log)
            continue
        if a.all and is_dup_file(rel):
            log(f"  ⏭ 跳過副本: {rel}")
            continue
        try:
            process_file(root, rel, tpl, out_dir, a.preview, a.with_overlay, log)
        except Exception as e:
            import traceback
            log(f"  ✗ 出錯: {type(e).__name__}: {e}")
            log(traceback.format_exc())

    rp = root / ("_align_preview.txt" if a.preview else "_align_log.txt")
    rp.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n✓ log 寫入 {rp.resolve()}")


if __name__ == "__main__":
    main()
