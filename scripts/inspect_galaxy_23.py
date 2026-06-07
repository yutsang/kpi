"""Inspect Galaxy 2023 取數 structure for wiring.

Galaxy's 'Cover' sheet (key = DICJ Code, a B-code subproject) carries NG11 Category +
Project + Descriptions + 項目類別 + Cover Amount = Σ of 9 component columns, each drawn
from a different source tab (some across 2 files, some encrypted, some header-offset).

This dumps, per source: named 取數 column total (+ by-letter fallback) so each source
reconciles to its Cover component; the subproject key (prefers 'DICJ Code'); per-subproject
totals; candidate cols for ng / account / project / company / desc; a raw header=None
preview. Plus a CAPEX-locator probe (lists every SAP/CIP-ish sheet's amount totals) to find
where the 1.556B 'CAPEX SAP Record Amount' lives.

galaxy_raw_databook.xlsx is password-protected — pass the password at runtime (NOT stored):
  python scripts/inspect_galaxy_23.py --pw <password>
Output: prints + results/galaxy_2023_inspect.txt
"""
from __future__ import annotations
import argparse, io, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SUBPROJ_RE = re.compile(r"[A-Za-z]\d+\.\d{1,}")   # B3.50 / B1.10 / A1.2 ...

SLOTS = {
    "ng_code":      ["ng code", "ngcode", "ng11 system", "ng11 category", "dicj code"],
    "ng_label":     ["ng11 category", "投資領域", "投資方向", "範疇", "項目類別", "nature"],
    "account_code": ["account code", "gl account", "g/l", "cost element", "科目"],
    "account_desc": ["account desc", "account description", "gl account desc", "科目名"],
    "project":      ["project", "投資項目", "項目名", "项目名", "descriptions", "description"],
    "company":      ["company", "entity", "公司", "property", "profit center", "outlet"],
    "desc":         ["description", "memo", "narration", "voucher discrop", "備注", "remark"],
    "amount":       ["amount", "reported amt", "reported amount", "investment", "金額", "金额", "base amount"],
}

COVER_COMPONENTS = [
    ("J", "OPEX SAP Record Amount"), ("K", "CAPEX SAP Record Amount"), ("L", "VR Record"),
    ("M", "Simulation"), ("N", "Pre Opening"), ("O", "VMS"), ("P", "Corporate support"),
    ("Q", "PMS"), ("R", "Other source detail"),
]

# sheets to probe for amount totals when hunting the CAPEX 1.556B (sap file + databook)
CAPEX_PROBE_SAP = ["SAP Record-non gaming", "SAP Record-gaming", "SAP Record KDC",
                   "SAP Non Gaming", "FA對應的CIP明細", "沒有SAP Record", "GICC間接成本調整"]


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    d = ROOT / "data"
    hits = list(d.glob(f"*/raw/{name}")) + list(d.rglob(name)) + list(d.rglob(f"*{name}*"))
    exact = [h for h in hits if h.name == name]
    return (exact or hits)[0] if (exact or hits) else None


def handle(fp, pw):
    """return a readable handle; decrypt with msoffcrypto if pw given (fresh each call)."""
    if not pw:
        return fp
    try:
        import msoffcrypto
    except ImportError:
        raise RuntimeError("need msoffcrypto-tool for the encrypted book: pip install msoffcrypto-tool")
    buf = io.BytesIO()
    with open(fp, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        of.load_key(password=pw)
        of.decrypt(buf)
    buf.seek(0)
    return buf


def sheet_names(fp, pw):
    return pd.ExcelFile(handle(fp, pw)).sheet_names


def read(fp, sheet, header, pw):
    return pd.read_excel(handle(fp, pw), sheet_name=sheet, header=header, dtype=object)


def letter_to_idx(letter):
    idx = 0
    for ch in str(letter).strip().upper():
        if "A" <= ch <= "Z": idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def auto_header(fp, sheet, pw):
    raw = pd.read_excel(handle(fp, pw), sheet_name=sheet, header=None, nrows=8, dtype=object)
    return max(range(min(6, len(raw))), key=lambda i: raw.iloc[i].notna().sum()) if len(raw) else 0


def col_match(df, kws):
    return [str(c) for c in df.columns if any(k.lower() in str(c).lower() for k in kws)]


def detect_subproj_col(df):
    # prefer an exact 'DICJ Code' column
    for c in df.columns:
        if str(c).strip().lower() == "dicj code":
            return c, df[c].dropna().astype(str).str.contains(SUBPROJ_RE).mean()
    best, best_frac = None, 0.0
    for c in df.columns:
        s = df[c].dropna().astype(str)
        if s.empty: continue
        frac = s.str.fullmatch(SUBPROJ_RE).mean()   # fullmatch → avoids composite '判断1'
        if frac > best_frac: best, best_frac = c, frac
    return (best, best_frac) if best_frac >= 0.30 else (None, best_frac)


def num(s): return pd.to_numeric(s, errors="coerce")


def dump_source(L, spec, fp, pw):
    key, sheet = spec["key"], spec["sheet"]
    use_pw = pw if spec.get("enc") else None
    L.append(f"\n{'='*78}\n## {key}  file={fp.name}  sheet={sheet!r}")
    try:
        hdr = spec["header"] if spec.get("header") is not None else auto_header(fp, sheet, use_pw)
        df = read(fp, sheet, hdr, use_pw)
    except Exception as e:
        L.append(f"  X read failed: {e}"); return
    L.append(f"  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
    L.append("  columns: " + " | ".join(str(c) for c in df.columns))
    try:
        raw = pd.read_excel(handle(fp, use_pw), sheet_name=sheet, header=None, nrows=6, dtype=object)
        L.append("  -- raw 頭 6 行 (前 14 欄) --")
        for i in range(len(raw)):
            L.append(f"    r{i}: " + " | ".join(str(raw.iloc[i, j])[:16] for j in range(min(14, raw.shape[1]))))
    except Exception: pass

    if spec.get("filter"):
        fcol_spec, fval = spec["filter"]
        fcol = df.columns[letter_to_idx(fcol_spec)] if len(str(fcol_spec)) <= 2 and str(fcol_spec).isalpha() else fcol_spec
        if fcol in df.columns:
            before = len(df); df = df[df[fcol].astype(str).str.strip().str.lower() == fval.lower()]
            L.append(f"  -- filter {fcol!r} == {fval!r}: {before:,} → {len(df):,} 行 --")
        else: L.append(f"  ! filter col {fcol_spec!r} not found")

    amt = None
    if spec.get("amount") and spec["amount"] in df.columns:
        amt = num(df[spec["amount"]]); L.append(f"  -- 取數 (named {spec['amount']!r}): Σ={amt.sum():,.2f}  (n={amt.notna().sum():,}) --")
    elif spec.get("amount"):
        c = [c for c in df.columns if spec["amount"].lower() in str(c).lower()]
        if c: amt = num(df[c[0]]); L.append(f"  -- 取數 (fuzzy {c[0]!r}): Σ={amt.sum():,.2f}  (n={amt.notna().sum():,}) --")
    if spec.get("letter"):
        li = letter_to_idx(spec["letter"])
        if 0 <= li < len(df.columns):
            sL = num(df.iloc[:, li]); L.append(f"  -- 取數 (letter {spec['letter']}=col[{li}] {str(df.columns[li])[:30]!r}): Σ={sL.sum():,.2f}  (n={sL.notna().sum():,}) --")
            if amt is None: amt = sL

    sp, frac = detect_subproj_col(df)
    if sp is not None:
        # drop grand-total rows (subproject blank) when summing per-key
        d2 = df.assign(_sp=df[sp].astype(str).str.strip())
        d2 = d2[d2["_sp"].str.fullmatch(SUBPROJ_RE).fillna(False)]
        L.append(f"  -- subproject key = {sp!r}  ({frac:.0%} match, {d2['_sp'].nunique()} distinct; total-rows dropped {len(df)-len(d2)}) --")
        if amt is not None:
            d2 = d2.assign(_a=amt.reindex(d2.index).fillna(0))
            g = d2.groupby("_sp")["_a"].agg(amount="sum", n="size").reset_index().sort_values("amount", ascending=False, key=abs).head(40)
            for _, r in g.iterrows(): L.append(f"      {r['amount']:>16,.0f}  n={int(r['n']):>5}  {r['_sp'][:40]}")
    else:
        L.append(f"  -- subproject key: 未偵測到 (max {frac:.0%}) — key 可能係 NG11 System Code,要用 Cover 對映 --")

    L.append("  -- 候選欄 (raw col → slot) --")
    for slot, kws in SLOTS.items():
        c = col_match(df, kws)
        if c: L.append(f"    {slot:13s}: {' | '.join(c[:6])}")
    seen = set()
    for slot in ("ng_label", "account_code", "account_desc", "company"):
        for c in col_match(df, SLOTS[slot]):
            if c in seen: continue
            seen.add(c)
            vc = df[c].astype(str).str.strip().replace({"nan": "(空)"}).value_counts().head(12)
            L.append(f"    === {c!r} ({df[c].notna().sum():,} non-null, {df[c].nunique()} distinct) ===")
            for v, n in vc.items(): L.append(f"        {n:>6,}  {str(v)[:50]}")


def dump_cover(L, fp, pw):
    L.append(f"\n{'='*78}\n## 00_Cover  file={fp.name}  sheet='Cover'  (取數總表, key=DICJ Code)")
    hdr = auto_header(fp, "Cover", None)
    df = read(fp, "Cover", hdr, None)
    L.append(f"  header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
    L.append("  columns: " + " | ".join(str(c) for c in df.columns))
    sp, frac = detect_subproj_col(df)
    L.append(f"  -- subproject key = {sp!r} ({frac:.0%}) --")
    iI = letter_to_idx("I")
    cover_amt = num(df.iloc[:, iI]) if iI < len(df.columns) else None
    if cover_amt is not None:
        L.append(f"  -- I 'Cover Amount' col[{iI}]={str(df.columns[iI])[:24]!r}: Σ={cover_amt.sum():,.0f} --")
    L.append("  -- 9 components (letter → header → Σ) --")
    comp_sum = 0.0
    for letter, label in COVER_COMPONENTS:
        idx = letter_to_idx(letter)
        if idx < len(df.columns):
            s = num(df.iloc[:, idx]); comp_sum += s.sum()
            L.append(f"    {letter} col[{idx}] {str(df.columns[idx])[:34]:34s} Σ={s.sum():>16,.0f}  (expect {label})")
    if cover_amt is not None:
        gap = cover_amt.sum() - comp_sum
        L.append(f"    Σ(J..R)={comp_sum:,.0f}   I={cover_amt.sum():,.0f}   gap={gap:,.0f}  (likely '沒有SAP Record一次分類')")
    # extra reconciliation buckets
    for name in ["沒有SAP Record一次分類", "Staff cost大致分類", "COMP大致分類", "其他成本", "ADJ", "Combined"]:
        if name in df.columns:
            s = num(df[name]); L.append(f"    bucket {name!r}: Σ={s.sum():,.0f}  (n={s.notna().sum()})")
    # V/H label columns to inspect
    for name in ["NG11 Category", "項目類別", "Descriptions"]:
        if name in df.columns:
            vc = df[name].astype(str).str.strip().replace({"nan": "(空)"}).value_counts().head(20)
            L.append(f"    === {name!r} ({df[name].notna().sum():,} non-null, {df[name].nunique()} distinct) ===")
            for v, n in vc.items(): L.append(f"        {n:>6,}  {str(v)[:50]}")
    # per-subproject: cover + 9 components (top 60)
    if sp is not None and cover_amt is not None:
        idxs = {label: letter_to_idx(l) for l, label in COVER_COMPONENTS}
        out = df.assign(_sp=df[sp].astype(str).str.strip(), _cov=cover_amt.fillna(0))
        out = out[out["_sp"].str.fullmatch(SUBPROJ_RE).fillna(False)]
        for label, idx in idxs.items():
            out[label] = num(df.iloc[:, idx]).reindex(out.index).fillna(0) if idx < len(df.columns) else 0
        g = out.groupby("_sp").agg({"_cov": "sum", **{l: "sum" for l in idxs}}).reset_index().sort_values("_cov", ascending=False, key=abs).head(60)
        L.append(f"\n  -- per-subproject top 60: sp | Cover | {' | '.join(l for _, l in COVER_COMPONENTS)} --")
        for _, r in g.iterrows():
            comps = " ".join(f"{r[label]:>10,.0f}" for _, label in COVER_COMPONENTS)
            L.append(f"    {r['_sp'][:8]:8s} {r['_cov']:>13,.0f} | {comps}")


def capex_probe(L, sap, book, pw):
    L.append(f"\n{'='*78}\n## CAPEX 定位 probe — 邊張 sheet 夾出 1.556B 'CAPEX SAP Record Amount'")
    for fp, sheets, lbl, p in ((sap, CAPEX_PROBE_SAP, "sap", None),
                               (book, None, "book(encrypted)", pw)):
        if fp is None: L.append(f"  {lbl}: file missing"); continue
        try: avail = sheet_names(fp, p)
        except Exception as e: L.append(f"  {lbl}: sheet list failed — {e}"); continue
        tgt = sheets if sheets else avail
        for sh in tgt:
            if sh not in avail:
                if sheets: L.append(f"  {lbl}:{sh!r} — 唔存在")
                continue
            try:
                h = auto_header(fp, sh, p); d = read(fp, sh, h, p)
            except Exception as e:
                L.append(f"  {lbl}:{sh!r} — read failed {e}"); continue
            amt_cols = [c for c in d.columns if any(k in str(c).lower() for k in ["reported amount", "amount", "capex", "investment", "金額", "金额"])]
            tots = []
            for c in amt_cols[:6]:
                s = num(d[c])
                if s.notna().sum(): tots.append(f"{str(c)[:26]}=Σ{s.sum():,.0f}")
            L.append(f"  {lbl}:{sh!r}  rows={len(d):,} hdr={h}  | " + " ; ".join(tots[:6]))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sap", default="galaxy_sap.xlsx")
    ap.add_argument("--book", default="galaxy_raw_databook.xlsx")
    ap.add_argument("--pw", default=None, help="password for the encrypted databook (runtime only, never stored)")
    args = ap.parse_args()
    sap, book = find_file(args.sap), find_file(args.book)
    L = ["# Galaxy 2023 取數 inspect", f"# sap={sap}  book={book}  pw={'set' if args.pw else 'none'}"]
    for f, tag, p in ((sap, args.sap, None), (book, args.book, args.pw)):
        if f:
            try: L.append(f"# {tag} sheets = {sheet_names(f, p)}")
            except Exception as e: L.append(f"# {tag} sheet list failed: {e}")
        else: L.append(f"# X {tag} NOT FOUND under data/*/raw/ (incl. raw/2023/)")

    SOURCES = [
        dict(key="1_OPEX_SAP",          file=sap,  sheet="SAP Record-non gaming", amount="Reported Amount (MOP)"),
        dict(key="3_VR_Record",         file=sap,  sheet="VR Record",             amount="Amount", letter="K"),
        dict(key="4_Simulation",        file=sap,  sheet="Simulation&Pre-Opening",filter=("J", "Simulation")),
        dict(key="5a_PreOpening_filt",  file=sap,  sheet="Simulation&Pre-Opening",filter=("J", "Pre-Opening")),
        dict(key="5b_PreOpening_tab",   file=sap,  sheet="Pre Opening"),
        dict(key="6_VMS",               file=sap,  sheet="VMS",                   amount="Amount (Reported Amt) (MOP)", letter="K"),
        dict(key="7_Corporate_support", file=sap,  sheet="Corporate support",     header=18),
        dict(key="8_PMS",               file=sap,  sheet="PMS",                   amount="Amount (MOP)", letter="M"),
        dict(key="9_Other_Source",      file=sap,  sheet="Other Source Detail",   amount="Investment Amount", letter="L"),
    ]

    if sap: dump_cover(L, sap, None)
    capex_probe(L, sap, book, args.pw)
    for spec in SOURCES:
        if not spec["file"]: L.append(f"\n## {spec['key']}  X file missing"); continue
        dump_source(L, spec, spec["file"], args.pw)

    out = ROOT / "results" / "galaxy_2023_inspect.txt"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("\n".join(str(x) for x in L), encoding="utf-8")
    print("\n".join(str(x) for x in L))
    print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
