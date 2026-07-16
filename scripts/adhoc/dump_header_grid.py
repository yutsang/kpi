#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dump_header_grid.py — 把幾個檔嘅「完整表頭層級」逐欄吐出嚟（含合併表頭 / leading 欄）。

配合 inspect_headers.py：inspect 見到 表頭 左邊 6 欄係空（多行合併），呢個 dump
逐欄印出 row1..8 嘅堆疊值，睇到真正欄名同分組，方便砌 source→表頭 mapping。

Windows 跑：
    set PYTHONIOENCODING=utf-8
    python scripts\\adhoc\\dump_header_grid.py --root ad-hoc\\workspace

預設 dump：表頭 + 2 個代表 source_1（其他範疇-MGM、博監局-SJM）+ source_2 master tracker。
可自訂： --files "表頭.xlsx" "source_1\\旅遊局\\Wynn-....xlsx"
純讀，唔改檔。
"""
from __future__ import annotations

import argparse
import io
import sys
from pathlib import Path

import openpyxl

PASSWORD = "dicj_kpmg"
DEFAULT_FILES = [
    "表頭.xlsx",
    r"source_1\其他範疇\MGM-投資計劃執行情況表二（其他範疇）.xlsx",
    r"source_1\博監局\SJM-投資計劃執行情況表二（博監局）.xlsx",
    r"source_2\2025年年度投資計劃執行情況審查專項工作關注事項 - 承批公司回覆-跨司工作组反馈意见-待进一步確認事項.xlsx",
]


def load_wb(path: Path):
    try:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        try:
            import msoffcrypto
        except ImportError:
            raise RuntimeError("加密檔要 pip install msoffcrypto-tool")
        buf = io.BytesIO()
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=PASSWORD)
            off.decrypt(buf)
        buf.seek(0)
        return openpyxl.load_workbook(buf, read_only=True, data_only=True)


def col_letter(i: int) -> str:
    s = ""
    while i > 0:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _norm(x) -> str:
    return "" if x is None else str(x).strip().replace("\n", " ").replace("\r", " ")


def dump_sheet(ws, hdr_rows: int, data_rows: int, out):
    grid = []
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=hdr_rows + data_rows, values_only=True)):
        grid.append([_norm(c) for c in row])
    ncols = max((len(r) for r in grid), default=0)
    # trim 全空尾欄
    while ncols > 0 and all(len(r) < ncols or r[ncols - 1] == "" for r in grid):
        ncols -= 1
    out(f"    欄數(非空): {ncols}")
    for c in range(ncols):
        head_stack = [grid[r][c] for r in range(min(hdr_rows, len(grid))) if c < len(grid[r]) and grid[r][c]]
        data_vals = [grid[r][c] for r in range(hdr_rows, len(grid)) if c < len(grid[r]) and grid[r][c]]
        if not head_stack and not data_vals:
            continue
        hs = "  ↑  ".join(head_stack) if head_stack else "(空)"
        dv = f"   e.g. {data_vals[0]}" if data_vals else ""
        out(f"    [{c+1:>2} {col_letter(c+1):>2}] {hs}{dv}")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="ad-hoc/workspace")
    ap.add_argument("--files", nargs="*", default=None, help="相對 root 嘅檔（預設 4 個代表）")
    ap.add_argument("--hdr-rows", type=int, default=8, help="當頭幾行係表頭區")
    ap.add_argument("--data-rows", type=int, default=3, help="順手 dump 幾行 data 做例")
    ap.add_argument("--all-sheets", action="store_true", help="每個 sheet 都 dump（預設只第一個）")
    ap.add_argument("--pair", nargs=2, metavar=("範疇", "公司"),
                    help="自動揾 source_1/source_2 內檔名含呢兩個關鍵字嘅檔一齊 dump（睇 join key）")
    a = ap.parse_args()
    root = Path(a.root)
    files = list(a.files) if a.files else list(DEFAULT_FILES)
    if a.pair:
        kw_scope, kw_company = a.pair
        for src in ("source_1", "source_2"):
            base = root / src
            if not base.exists():
                continue
            for p in sorted(base.rglob("*")):
                if p.is_dir() or p.name.startswith("~$"):
                    continue
                if p.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
                    continue
                relstr = p.relative_to(root).as_posix()
                if kw_scope in relstr and kw_company in relstr:
                    files.append(p.relative_to(root).as_posix())
    lines: list[str] = []

    def out(s=""):
        lines.append(s)
        print(s)

    for rel in files:
        p = root / rel
        out("\n" + "=" * 78)
        out(f"# {rel}")
        if not p.exists():
            out(f"  ✗ 唔存在: {p}")
            continue
        try:
            wb = load_wb(p)
        except Exception as e:
            out(f"  ✗ 開唔到: {type(e).__name__}: {e}")
            continue
        out(f"  sheets: {wb.sheetnames}")
        targets = wb.sheetnames if a.all_sheets else [wb.sheetnames[0]]
        for sn in targets:
            out(f"\n  ── sheet: {sn!r} ──")
            dump_sheet(wb[sn], a.hdr_rows, a.data_rows, out)
        wb.close()

    rp = root / "_header_grid.txt"
    rp.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n✓ 寫入 {rp.resolve()}（貼返嚟俾我）")


if __name__ == "__main__":
    main()
