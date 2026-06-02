"""Auto-inspect data/table_2/2025_*_{Gaming,Non-gaming}.xlsx files.

Usage (Windows or Mac, repo root):
  python -m kpi.tools.table2_inspect

Reads every xlsx in data/table_2/ matching `2025_<Company>_<Gaming|Non-gaming>.xlsx`,
samples each sheet (header + 8 rows), and writes a single review file at
  data/table_2/_inspect_output.xlsx

with one sheet per source file. Stdout shows a compact summary so you can paste
the result back without opening the Excel.

Goal: replace manual paste for the sample template — let me see all 12 file
structures at once.
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

from kpi.lib.conf import ROOT
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()


SRC_DIR = ROOT / "data" / "table_2"
OUT = SRC_DIR / "_inspect_output.xlsx"
FNAME_RE = re.compile(r"^2025_([A-Za-z]+)_(Gaming|Non-gaming)\.xlsx$")

SAMPLE_ROWS = 60          # raw rows to keep per sheet (表二 vertical forms need ~50+ rows per project block)
MAX_SHEET_NAME = 28       # output sheet name length cap


def _safe_sheet_name(s: str) -> str:
    s = re.sub(r"[\\/*?:\[\]]", "_", s)
    return s[:MAX_SHEET_NAME]


def inspect_one(fpath: Path) -> tuple[str, list[dict]]:
    """Return (label, list_of_sheet_summaries)."""
    m = FNAME_RE.match(fpath.name)
    label = f"{m.group(1)}_{m.group(2)}" if m else fpath.stem
    summaries = []
    try:
        xl = pd.ExcelFile(fpath, engine="openpyxl")
    except Exception as e:
        print(f"  [error] {fpath.name}: {e}")
        return label, summaries
    for sheet in xl.sheet_names:
        # Don't skip 表一-only sheets — some sheets contain BOTH 表一 and 表二 data stacked.
        # We'll mark each row with a probable section label so the user can see the split.
        try:
            # header=None → preserve original cell layout so we can see title rows + real header row
            df_raw = pd.read_excel(fpath, sheet_name=sheet, dtype=str, header=None,
                                   nrows=SAMPLE_ROWS, engine="openpyxl")
        except Exception as e:
            summaries.append({"sheet": sheet, "error": str(e), "df": None})
            continue
        # Count non-null cells per row (helps guess which row is the real header)
        nn = df_raw.notna().sum(axis=1).tolist()
        # Two scans:
        #   markers: rows containing 表一/表二 (section break, anchors with "1." etc.)
        #   labels:  any row whose cells contain a Chinese "：" label
        #            (captures the full vertical-form field list)
        markers = []
        labels = []
        for i, row in df_raw.iterrows():
            cells = [str(v).strip() for v in row.tolist() if pd.notna(v)]
            joined = " | ".join(cells)
            tags = []
            for keyword in ("表一", "表二", "投資項目序號"):
                if any(keyword in c for c in cells):
                    tags.append(keyword)
            if tags:
                markers.append((int(i) + 1, tags, joined[:120]))
            # Capture label rows — short cells ending with "：" likely field labels
            label_cells = [c for c in cells if c.endswith("：") and len(c) < 50]
            if label_cells:
                labels.append((int(i) + 1, label_cells[0], joined[:160]))
        # Guess: the row in first 10 with the most non-null cells
        scan = nn[:10]
        header_row = int(max(range(len(scan)), key=lambda i: scan[i])) if scan else 0
        # Re-read with detected header, sample data rows
        try:
            df_h = pd.read_excel(fpath, sheet_name=sheet, dtype=str, header=header_row,
                                 nrows=header_row + 50, engine="openpyxl")
        except Exception:
            df_h = None
        summaries.append({
            "sheet": sheet,
            "n_cols": len(df_raw.columns),
            "nn_per_row": nn,
            "header_row": header_row,            # 0-indexed
            "header_cols": list(df_h.columns) if df_h is not None else [],
            "df_raw": df_raw,
            "df_with_header": df_h.head(50) if df_h is not None else None,
            "markers": markers,
            "labels": labels,
        })
    return label, summaries


def main() -> None:
    if not SRC_DIR.exists():
        print(f"Not found: {SRC_DIR}")
        return
    files = sorted(p for p in SRC_DIR.glob("2025_*_*.xlsx") if p.name != "_sample_template.xlsx"
                   and not p.name.startswith("_"))
    if not files:
        print(f"No matching files in {SRC_DIR}")
        print("Expected pattern: 2025_<Company>_<Gaming|Non-gaming>.xlsx")
        return

    print(f"Found {len(files)} files in {SRC_DIR}\n")

    SRC_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUT, engine="xlsxwriter") as writer:
        # Index sheet first
        idx_rows = []
        for fp in files:
            label, summaries = inspect_one(fp)
            print(f"── {fp.name}  →  label: {label} ──")
            for s in summaries:
                if "error" in s:
                    print(f"  [{s['sheet']}] read error: {s['error']}")
                    idx_rows.append({"file": fp.name, "label": label, "sheet": s["sheet"],
                                     "n_cols": "", "header_row": "", "status": "error",
                                     "header_cols": ""})
                    continue
                hr = s["header_row"]    # 0-indexed
                cols = s["header_cols"]
                print(f"  [{s['sheet']}]  header_row={hr+1} (1-indexed)  n_cols={s['n_cols']}")
                print(f"    columns: {cols}")
                if s.get("markers"):
                    print(f"    section markers (row → keywords):")
                    for row_no, tags, preview in s["markers"][:20]:
                        print(f"      row {row_no:>3}  {tags}  | {preview}")
                if s.get("labels"):
                    # Distinct field labels seen in this sheet (one occurrence each, keep order)
                    seen = []
                    for _, lbl, _ in s["labels"]:
                        if lbl not in seen:
                            seen.append(lbl)
                    print(f"    field labels ({len(seen)} distinct):")
                    for lbl in seen:
                        print(f"      • {lbl}")
                idx_rows.append({
                    "file": fp.name, "label": label, "sheet": s["sheet"],
                    "n_cols": s["n_cols"], "header_row": hr + 1,
                    "header_cols": " | ".join(str(c) for c in cols),
                    "nn_per_row": str(s["nn_per_row"]),
                    "status": "ok",
                })
                # Sheet A: raw layout (preserves row positions)
                out_raw = _safe_sheet_name(f"{label}__{s['sheet']}_raw")
                s["df_raw"].to_excel(writer, sheet_name=out_raw, index=False, header=False)
                # Sheet B: parsed with detected header
                if s["df_with_header"] is not None:
                    out_hdr = _safe_sheet_name(f"{label}__{s['sheet']}_h")
                    s["df_with_header"].to_excel(writer, sheet_name=out_hdr, index=False)
            print()
        pd.DataFrame(idx_rows).to_excel(writer, sheet_name="_index", index=False)

    print(f"\nWrote: {OUT.relative_to(ROOT)}")
    print("Send this file back (or paste the stdout above) so I can design the normalize schema.")


if __name__ == "__main__":
    main()
