"""Inspect the RAW SOURCE sheets behind the 2023 prebuilds (galaxy sap tabs + mgm tabs) so the
build scripts can be re-mapped to REAL columns. The audit found whole blocks get blank account_code
+ constant account_desc (→ signature collapse) and NG polluted by 項目類別 / wrong-block. To fix
the mapping we need to SEE each source tab's actual columns.

For each (entity, file, tab) it auto-detects the header row (densest of first 8 rows), then dumps
per column: non-blank %, distinct count, 3 sample values — and flags likely roles:
  [GL?]   = looks like a GL/account code (mostly digits, 4+ chars)
  [DESC?] = free-text description (high distinct, alpha)
  [NG?]   = NG0–NG11 codes or DICJ theme names
  [AMT?]  = numeric amount

Run (Windows):  python scripts/inspect_23_sources.py [--pw dicj_kpmg]
Output: prints + results/inspect_23_sources.txt
"""
from __future__ import annotations
import argparse, io, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
NG_THEME = ("NG", "博彩", "吸引", "會議", "娛樂", "體育", "文化", "健康", "主題", "美食", "社區", "海上", "其他", "藝術")

# (entity, filename, [tabs]) — tabs the prebuild reads. None tab = inspect all sheet names.
SRC = [
    ("galaxy", "galaxy_sap.xlsx", ["SAP Record-non gaming", "PMS", "VMS", "VR Record",
                                   "Simulation&Pre-Opening", "Corporate support", "Other Source Detail"]),
    ("mgm",    "OPEX.xlsx",       ["WD#1", "WD#2", "WD#3", "PM", "Data#4", "Data#5", "Data#6"]),
    ("mgm",    "CAPEX.xlsx",      None),
    ("mgm",    "MGM-gaming.xlsx", None),
]


def find_file(name):
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def open_book(fp, pw):
    if pw:
        try:
            import msoffcrypto
            buf = io.BytesIO()
            with open(fp, "rb") as f:
                of = msoffcrypto.OfficeFile(f); of.load_key(password=pw); of.decrypt(buf)
            buf.seek(0); return pd.ExcelFile(buf)
        except Exception:
            pass
    return pd.ExcelFile(fp)


def auto_header(xl, sheet):
    raw = xl.parse(sheet, header=None, nrows=8, dtype=object)
    best, bestn = 0, -1
    for i in range(min(8, len(raw))):
        nn = raw.iloc[i].notna().sum()
        if nn > bestn: best, bestn = i, nn
    return best


def role(s):
    nb = s[s.ne("")]
    if not len(nb): return ""
    digit = nb.str.fullmatch(r"-?[\d,]+\.?\d*").mean()
    isng = nb.isin(["NG0", "NG1", "NG2", "NG3", "NG4", "NG5", "NG6", "NG7", "NG8", "NG9", "NG10", "NG11"]).mean()
    theme = nb.apply(lambda v: any(t in str(v) for t in NG_THEME)).mean()
    nun = nb.nunique()
    glish = nb.str.fullmatch(r"\d{4,}").mean()
    tags = []
    if isng > 0.5 or theme > 0.5: tags.append("NG?")
    if glish > 0.4: tags.append("GL?")
    elif digit > 0.8: tags.append("AMT?")
    if nun > max(20, 0.3 * len(nb)) and digit < 0.3: tags.append("DESC?")
    return " ".join(f"[{t}]" for t in tags)


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--pw", default=None); a = ap.parse_args()
    L = ["# inspect_23_sources — real columns of the 2023 source tabs (for re-mapping the builds)"]
    seen = {}
    for ent, fname, tabs in SRC:
        fp = seen.get(fname) or find_file(fname)
        if not fp:
            L.append(f"\n## {ent} / {fname}: X not found"); continue
        seen[fname] = fp
        try:
            xl = open_book(fp, a.pw)
        except Exception as e:
            L.append(f"\n## {ent} / {fname}: X open failed: {e}"); continue
        use = tabs or xl.sheet_names
        L.append(f"\n{'='*72}\n## {ent} / {fname}  (sheets: {xl.sheet_names})")
        for sheet in use:
            if sheet not in xl.sheet_names:
                L.append(f"\n  -- tab {sheet!r} NOT in book"); continue
            try:
                hdr = auto_header(xl, sheet)
                df = xl.parse(sheet, header=hdr, dtype=object)
            except Exception as e:
                L.append(f"\n  -- {sheet}: read failed {e}"); continue
            L.append(f"\n  ▸ {sheet}  (auto header_row={hdr}, {len(df):,} rows, {len(df.columns)} cols)")
            for c in df.columns:
                s = df[c].astype("string").fillna("").str.strip()
                nb = s.ne("").mean() * 100
                nun = s[s.ne("")].nunique()
                samp = " | ".join(map(str, pd.Series(s[s.ne('')].unique())[:3]))
                L.append(f"      {str(c)[:32]:32s} nb{nb:5.0f}% uniq{nun:>6} {role(s):14s} {samp[:70]}")
    out = ROOT / "results" / "inspect_23_sources.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
