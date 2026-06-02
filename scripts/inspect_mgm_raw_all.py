"""Thorough STRUCTURAL inspect of ALL MGM raw files (f1–f4 named in conf prebuild_sources).

Why: row-level H mapping (account_code/description → 建設/設備/廣告/人工/comp/…) can only come
from the actual raw ledger detail — NOT the summary sheet (which only has WD1-4 aggregates per
項目, hence everything 待拆). Those raw files aren't on my side, so this dumps their real
structure — every sheet, every column (dtype / %filled / #unique / samples), detected
account-code & 項目 & WD & amount columns, and a few full sample rows — so we can wire the
account_code → H rules to the right columns and tie to golden at row level.

Run on Windows (data/mgm/raw/ lives there). Writes a concise report you can paste back:
  python scripts/inspect_mgm_raw_all.py
  python scripts/inspect_mgm_raw_all.py --rows 4 --maxcells 18
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent


def _trunc(x, n=22):
    s = str(x).replace("\n", " ").replace("\r", " ").strip()
    return (s[:n] + "…") if len(s) > n else s


def _looks_account(sv):
    """Fraction of values that look like a GL account code (5-8 digits, maybe a suffix)."""
    hit = sv.str.fullmatch(r"\d{5,8}(\.\d+)?").fillna(False).mean()
    return float(hit)


def _looks_amount(col):
    num = pd.to_numeric(col, errors="coerce")
    return float(num.notna().mean()), float(num.fillna(0).sum())


def _classify_col(name, col):
    nm = str(name).lower()
    sv = col.astype("string").fillna("").str.strip()
    nonblank = (sv != "").mean()
    nunq = sv.nunique(dropna=True)
    tags = []
    if "項目" in str(name) or "project" in nm or "item" in nm or "task" in nm:
        tags.append("PROJ?")
    if any(k in nm for k in ("session", "範疇", "ng", "scope", "投資")):
        tags.append("NG/SESS?")
    if re.search(r"wd\s*\d|patron|payroll|cogs|allocation", nm):
        tags.append("WD?")
    if any(k in nm for k in ("account", "ledger", "科目", "gl")) or _looks_account(sv) > 0.5:
        tags.append("ACCT?")
    if any(k in nm for k in ("desc", "memo", "摘要", "narr", "name")):
        tags.append("DESC?")
    afrac, asum = _looks_amount(col)
    if afrac > 0.7 and nunq > 5:
        tags.append(f"AMT?Σ={asum:,.0f}")
    samples = " | ".join(_trunc(v) for v in sv[sv != ""].drop_duplicates().head(2))
    return nonblank, nunq, " ".join(tags), samples


def dump_file(path: Path, rows: int, maxcells: int, out):
    def w(s=""):
        print(s); out.append(s)
    w(f"\n{'='*92}\nFILE: {path.name}\n{'='*92}")
    if not path.exists():
        w(f"  X MISSING ({path})"); return
    try:
        sheets = pd.read_excel(path, sheet_name=None, header=None, dtype=str)
    except Exception as e:
        w(f"  X read failed: {e}"); return
    for sname, raw in sheets.items():
        nr, nc = raw.shape
        w(f"\n--- sheet '{sname}'  ({nr:,} rows × {nc} cols) ---")
        # guess header row = first row in the top 12 that is mostly non-blank text
        hdr = 0
        for i in range(min(12, nr)):
            r = raw.iloc[i].astype("string").fillna("").str.strip()
            if (r != "").mean() > 0.5 and r.str.contains(r"[A-Za-z一-鿿]").mean() > 0.4:
                hdr = i; break
        w(f"  header guess = row {hdr}: " +
          " | ".join(f"[{j}]{_trunc(v,16)}" for j, v in enumerate(raw.iloc[hdr].tolist()[:maxcells])))
        body = raw.iloc[hdr + 1:].reset_index(drop=True)
        names = raw.iloc[hdr].tolist()
        w(f"  {'idx':>3} {'column':22} {'%fill':>6} {'#uniq':>7}  tags / samples")
        for j in range(min(nc, maxcells)):
            col = body.iloc[:, j] if j < body.shape[1] else pd.Series([], dtype=str)
            nb, nq, tags, samp = _classify_col(names[j], col)
            w(f"  {j:>3} {_trunc(names[j],22):22} {nb*100:5.0f}% {nq:>7}  {tags}  {samp}")
        w(f"  sample {rows} data rows:")
        for _, r in body.head(rows).iterrows():
            w("    " + " | ".join(_trunc(v, 14) for v in r.tolist()[:maxcells]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=4, help="sample data rows per sheet")
    ap.add_argument("--maxcells", type=int, default=16, help="max columns shown per row")
    ap.add_argument("--raw", default="data/mgm/raw")
    args = ap.parse_args()

    cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    srcs = cfg.get("prebuild_sources") or {}
    files = [(k, ROOT / args.raw / v) for k, v in srcs.items() if v]
    out: list[str] = []
    print(f"MGM raw inspect — {len(files)} files from conf prebuild_sources:")
    for k, p in files:
        print(f"  {k}: {p.name}  {'(missing)' if not p.exists() else ''}")
    for k, p in files:
        dump_file(p, args.rows, args.maxcells, out)

    rep = ROOT / "results" / "mgm_raw_structure.txt"
    rep.parent.mkdir(parents=True, exist_ok=True)
    rep.write_text("\n".join(out), encoding="utf-8")
    print(f"\n→ wrote {rep}  ({len(out)} lines) — paste 返呢個檔畀我。")


if __name__ == "__main__":
    main()
