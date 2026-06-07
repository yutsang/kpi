"""Inspect Galaxy 2023 取數 structure for wiring.

Galaxy's 'Cover' sheet has col I 'Cover Amount' = sum of 9 component columns (J..R),
each drawn from a different source tab — some across 2 files, some with the 取數 column
given by letter, some not starting on row 1. This dumps, per source:
  - the named 取數 column total (+ by-letter fallback) so each source reconciles to its
    Cover component,
  - the detected subproject key column (B3.50xxx pattern) + per-subproject totals,
  - candidate columns for ng_code / ng_label / account / account_desc / project /
    company / desc (so we can map source → OUR taxonomy),
  - a raw header=None preview (catches header-offset tabs like 'Corporate support').

Goal = a per-subproject one-to-one mapping across the 9 sheets BEFORE any wiring.

Run (Windows):
  python scripts/inspect_galaxy_23.py
  python scripts/inspect_galaxy_23.py --sap galaxy_sap.xlsx --book galaxy_raw_databook.xlsx
Output: prints + results/galaxy_2023_inspect.txt
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# subproject key like "B3.50123" / "B3.5001" — letter, digits, dot, >=2 digits
SUBPROJ_RE = re.compile(r"[A-Za-z]\d+\.\d{2,}")

# candidate-column keywords per logical slot (dumped per tab)
SLOTS = {
    "ng_code":      ["ng code", "ng_code", "ng no", "ng11", "領域編", "範疇編"],
    "ng_label":     ["ng label", "ng name", "nature of ng", "投資領域", "投資方向", "範疇", "category", "性質"],
    "account_code": ["account code", "gl account", "g/l", "科目代", "cost element", "ledger account", "entry account", "account no"],
    "account_desc": ["account desc", "account name", "gl desc", "科目名", "cost element desc", "ledger account desc", "long text", "description of account"],
    "project":      ["project name", "project desc", "投資項目", "項目名", "项目名", "project code", "project no", "initiative"],
    "subproject":   ["sub project", "subproject", "sub-project", "子項目", "wbs", "b3", "internal order", "io "],
    "company":      ["company", "entity", "公司", "legal", "property", "casino", "venue"],
    "desc":         ["description", "memo", "narration", "text", "remark", "摘要", "details", "particular"],
    "amount":       ["amount", "reported amt", "reported amount", "investment", "mop", "金額", "金额"],
}

# the 9 Cover components, in column order J..R (Cover Amount = col I = Σ of these)
COVER_COMPONENTS = [
    ("J", "OPEX SAP Record Amount"),
    ("K", "CAPEX SAP Record Amount"),
    ("L", "VR Record"),
    ("M", "Simulation"),
    ("N", "Pre Opening"),
    ("O", "VMS"),
    ("P", "Corporate support"),
    ("Q", "PMS"),
    ("R", "Other source detail"),
]

# per-source tab spec. header=None -> auto-detect; int -> 0-based override.
# amount = named 取數 col; letter = positional fallback; filter = (col_or_letter, value).
def sources(sap, book):
    return [
        dict(key="00_Cover",            file=sap,  sheet="Cover",                 header=None, amount="Cover Amount", letter="I", cover=True),
        dict(key="1_OPEX_SAP",          file=sap,  sheet="SAP Record-non gaming", header=None, amount="Reported Amount (MOP)"),
        dict(key="2a_CAPEX_SAP_Gaming", file=book, sheet="SAP Gaming",            header=None, amount="Reported Amount (MOP)"),
        dict(key="2b_CAPEX_SAP_NonGam", file=sap,  sheet="SAP Non Gaming",        header=None, amount="Reported Amount (MOP)"),
        dict(key="3_VR_Record",         file=sap,  sheet="VR Record",             header=None, amount="Amount", letter="K"),
        dict(key="4_Simulation",        file=sap,  sheet="Simulation&Pre-Opening",header=None, filter=("J", "Simulation")),
        dict(key="5a_PreOpening_filt",  file=sap,  sheet="Simulation&Pre-Opening",header=None, filter=("J", "Pre-Opening")),
        dict(key="5b_PreOpening_tab",   file=sap,  sheet="Pre Opening",           header=None),
        dict(key="6_VMS",               file=sap,  sheet="VMS",                   header=None, amount="Amount (Reported Amt) (MOP)", letter="K"),
        dict(key="7_Corporate_support", file=sap,  sheet="Corporate support",     header=18,   letter="R"),   # data starts excel row 20
        dict(key="8_PMS",               file=sap,  sheet="PMS",                   header=None, amount="Amount (MOP)", letter="M"),
        dict(key="9_Other_Source",      file=sap,  sheet="Other Source Detail",   header=None, amount="Investment", letter="L"),
    ]


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    hits = list((ROOT / "data").glob(f"*/raw/{name}")) + list((ROOT / "data").glob(f"*/raw/*{name}*"))
    return hits[0] if hits else None


def letter_to_idx(letter):
    idx = 0
    for ch in str(letter).strip().upper():
        if "A" <= ch <= "Z":
            idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def auto_header(xl_path, sheet):
    raw = pd.read_excel(xl_path, sheet_name=sheet, header=None, nrows=8, dtype=object)
    if not len(raw): return 0
    return max(range(min(6, len(raw))), key=lambda i: raw.iloc[i].notna().sum())


def col_match(df, kws):
    return [str(c) for c in df.columns if any(k.lower() in str(c).lower() for k in kws)]


def detect_subproj_col(df):
    """column whose values most look like B3.50xxx subproject keys."""
    best, best_frac = None, 0.0
    for c in df.columns:
        s = df[c].dropna().astype(str)
        if s.empty: continue
        frac = s.str.contains(SUBPROJ_RE).mean()
        if frac > best_frac:
            best, best_frac = c, frac
    return (best, best_frac) if best_frac >= 0.20 else (None, best_frac)


def num(series):
    return pd.to_numeric(series, errors="coerce")


def dump_source(L, spec, fp):
    key, sheet = spec["key"], spec["sheet"]
    L.append(f"\n{'='*78}\n## {key}  file={fp.name}  sheet={sheet!r}")
    try:
        hdr = spec["header"] if spec.get("header") is not None else auto_header(fp, sheet)
        df = pd.read_excel(fp, sheet_name=sheet, header=hdr, dtype=object)
    except Exception as e:
        L.append(f"  X read failed: {e}"); return
    L.append(f"  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
    L.append("  columns: " + " | ".join(str(c) for c in df.columns))

    # raw preview (header=None) — catches offset/multi-header tabs
    try:
        raw = pd.read_excel(fp, sheet_name=sheet, header=None, nrows=6, dtype=object)
        L.append("  -- raw 頭 6 行 (header=None, 前 14 欄) --")
        for i in range(len(raw)):
            L.append(f"    r{i}: " + " | ".join(str(raw.iloc[i, j])[:16] for j in range(min(14, raw.shape[1]))))
    except Exception:
        pass

    # optional Nature filter (Simulation / Pre-Opening)
    if spec.get("filter"):
        fcol_spec, fval = spec["filter"]
        fcol = df.columns[letter_to_idx(fcol_spec)] if len(fcol_spec) <= 2 and fcol_spec.isalpha() else fcol_spec
        if fcol in df.columns:
            before = len(df)
            df = df[df[fcol].astype(str).str.strip().str.lower() == fval.lower()]
            L.append(f"  -- filter {fcol!r} == {fval!r}: {before:,} → {len(df):,} 行 --")
        else:
            L.append(f"  ! filter col {fcol_spec!r} not found")

    # 取數 column total: by name, then by letter
    amt_series = None
    if spec.get("amount") and spec["amount"] in df.columns:
        amt_series = num(df[spec["amount"]])
        L.append(f"  -- 取數 (named {spec['amount']!r}): Σ={amt_series.sum():,.2f}  (n={amt_series.notna().sum():,}) --")
    elif spec.get("amount"):
        # fuzzy: contains
        cands = [c for c in df.columns if spec["amount"].lower() in str(c).lower()]
        if cands:
            amt_series = num(df[cands[0]])
            L.append(f"  -- 取數 (fuzzy {cands[0]!r}): Σ={amt_series.sum():,.2f}  (n={amt_series.notna().sum():,}) --")
    if spec.get("letter"):
        li = letter_to_idx(spec["letter"])
        if 0 <= li < len(df.columns):
            sL = num(df.iloc[:, li])
            L.append(f"  -- 取數 (letter {spec['letter']}=col[{li}] {str(df.columns[li])[:30]!r}): Σ={sL.sum():,.2f}  (n={sL.notna().sum():,}) --")
            if amt_series is None: amt_series = sL

    # subproject key column + per-subproject totals
    sp_col, frac = detect_subproj_col(df)
    if sp_col is not None:
        L.append(f"  -- subproject key col = {sp_col!r}  ({frac:.0%} 值似 B3.50xxx, {df[sp_col].nunique()} distinct) --")
        if amt_series is not None:
            g = (df.assign(_sp=df[sp_col].astype(str).str.strip(), _a=amt_series.fillna(0))
                   .groupby("_sp")["_a"].agg(amount="sum", n="size").reset_index()
                   .sort_values("amount", ascending=False, key=abs).head(40))
            for _, r in g.iterrows():
                L.append(f"      {r['amount']:>16,.0f}  n={int(r['n']):>4}  {r['_sp'][:48]}")
    else:
        L.append(f"  -- subproject key col: 未偵測到 (max {frac:.0%}) — 要人手指定 --")

    # candidate columns per slot
    L.append("  -- 候選欄 (raw col → slot) --")
    for slot, kws in SLOTS.items():
        cand = col_match(df, kws)
        if cand:
            L.append(f"    {slot:13s}: {' | '.join(cand[:6])}")

    # value dists for ng/account/project/company candidates
    seen = set()
    for slot in ("ng_label", "ng_code", "account_code", "account_desc", "project", "company"):
        for c in col_match(df, SLOTS[slot]):
            if c in seen: continue
            seen.add(c)
            vc = df[c].astype(str).str.strip().replace({"nan": "(空)"}).value_counts().head(15)
            L.append(f"    === {c!r} ({df[c].notna().sum():,} non-null, {df[c].nunique()} distinct) ===")
            for v, n in vc.items():
                L.append(f"        {n:>6,}  {str(v)[:54]}")


def dump_cover(L, spec, fp):
    """Cover: col I 'Cover Amount' = Σ of 9 components J..R, by subproject."""
    L.append(f"\n{'='*78}\n## {spec['key']}  file={fp.name}  sheet='Cover'  (取數總表)")
    try:
        hdr = auto_header(fp, "Cover")
        df = pd.read_excel(fp, sheet_name="Cover", header=hdr, dtype=object)
    except Exception as e:
        L.append(f"  X read failed: {e}"); return
    L.append(f"  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
    L.append("  columns: " + " | ".join(str(c) for c in df.columns))
    raw = pd.read_excel(fp, sheet_name="Cover", header=None, nrows=6, dtype=object)
    L.append("  -- raw 頭 6 行 (前 20 欄) --")
    for i in range(len(raw)):
        L.append(f"    r{i}: " + " | ".join(str(raw.iloc[i, j])[:14] for j in range(min(20, raw.shape[1]))))

    sp_col, frac = detect_subproj_col(df)
    L.append(f"  -- subproject key col = {sp_col!r} ({frac:.0%}) --")
    iI = letter_to_idx("I")
    cover_amt = num(df.iloc[:, iI]) if iI < len(df.columns) else None
    if cover_amt is not None:
        L.append(f"  -- I 'Cover Amount' (col[{iI}]={str(df.columns[iI])[:30]!r}): Σ={cover_amt.sum():,.2f} --")
    L.append("  -- 9 components (letter → header → Σ) --")
    comp_sum = 0.0
    for letter, label in COVER_COMPONENTS:
        idx = letter_to_idx(letter)
        if idx < len(df.columns):
            s = num(df.iloc[:, idx]); comp_sum += s.sum()
            L.append(f"    {letter} col[{idx}] {str(df.columns[idx])[:38]:38s} Σ={s.sum():>16,.0f}  (expect: {label})")
    L.append(f"    Σ(J..R)={comp_sum:,.0f}   vs  I={cover_amt.sum():,.0f}" if cover_amt is not None else "")

    # per-subproject: cover amount + each component (so we can join sources to it)
    if sp_col is not None and cover_amt is not None:
        comp_idx = {label: letter_to_idx(l) for l, label in COVER_COMPONENTS}
        out = df.assign(_sp=df[sp_col].astype(str).str.strip(), _cov=cover_amt.fillna(0))
        for label, idx in comp_idx.items():
            out[label] = num(df.iloc[:, idx]).fillna(0) if idx < len(df.columns) else 0
        agg = {"_cov": "sum", **{label: "sum" for label in comp_idx}}
        g = out.groupby("_sp").agg(agg).reset_index().sort_values("_cov", ascending=False, key=abs).head(60)
        L.append(f"\n  -- per-subproject (top 60 by Cover Amount) — sp | Cover | {' | '.join(l for _, l in COVER_COMPONENTS)} --")
        for _, r in g.iterrows():
            comps = "  ".join(f"{r[label]:>11,.0f}" for _, label in COVER_COMPONENTS)
            L.append(f"    {r['_sp'][:30]:30s} {r['_cov']:>14,.0f} | {comps}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sap", default="galaxy_sap.xlsx")
    ap.add_argument("--book", default="galaxy_raw_databook.xlsx")
    args = ap.parse_args()
    sap = find_file(args.sap); book = find_file(args.book)
    L = [f"# Galaxy 2023 取數 inspect", f"# sap={sap}  book={book}"]
    for f, tag in ((sap, args.sap), (book, args.book)):
        if f:
            try: L.append(f"# {tag} sheets = {pd.ExcelFile(f).sheet_names}")
            except Exception as e: L.append(f"# {tag} sheet list failed: {e}")
        else:
            L.append(f"# X {tag} NOT FOUND under data/*/raw/ — put it there or pass --sap/--book")

    for spec in sources(sap, book):
        fp = spec["file"]
        if not fp:
            L.append(f"\n## {spec['key']}  X file missing"); continue
        if spec.get("cover"):
            dump_cover(L, spec, fp)
        else:
            dump_source(L, spec, fp)

    out = ROOT / "results" / "galaxy_2023_inspect.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(str(x) for x in L), encoding="utf-8")
    print("\n".join(str(x) for x in L))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
