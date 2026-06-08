"""Data-quality AUDIT of the prebuilt 2023 raws (galaxy/sjm/mgm) — the user reports the 23 data
has real errors: NG column polluted with capex/non-NG values, many zero rows, and far too few
distinct signatures vs row count. This dumps, PER SOURCE block, the fill rates + value sanity so
the errors are visible (not hidden behind '100% have a V/H label').

For each entity it reads the BUILT 23 raw xlsx (the file the pipeline ingests) and reports:
  - per Source: rows, Σ|amt|, zero-amount rows, account_code/account_desc/project/NG fill% + nunique
  - NG value SANITY: top distinct NG values; flags any that look like capex/opex/amounts (= pollution)
  - signature collapse: distinct (account_code|account_desc) vs rows — low ratio = collapse
  - blank-key rows: how many rows have BOTH account_code AND account_desc blank (→ one giant sig)

Run (Windows):  python scripts/audit_23_raw.py
Output: prints + results/audit_23_raw.txt
"""
from __future__ import annotations
import re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# entity → (raw xlsx rel path, sheet, {role: candidate column names})
ENTS = {
    "galaxy": dict(
        raw="data/galaxy/raw/galaxy_23_raw.xlsx", sheet="combine",
        amt=["Reported Amount(MOP)"], ac=["Account Code"], ad=["Account Description"],
        proj=["Project"], ng=["NG11 Category"], desc=["Description"], src=["Source"]),
    "sjm": dict(
        raw="data/sjm/raw/sjm_23_raw.xlsx", sheet="combine",
        amt=["Val/COArea Crcy"], ac=["Cost Element"], ad=["Cost element descr."],
        proj=["Project Name"], ng=["項目性質"], desc=["Cost element descr."], src=["Capex", "Source"]),
    "mgm": dict(
        raw="data/mgm/raw/mgm_23_raw.xlsx", sheet="combine",
        amt=["Debit minus Credit"], ac=["Ledger Account"], ad=["Ledger Hierarchy Level 5"],
        proj=["Project_code"], ng=["Section.1"], desc=["Ledger Hierarchy Level 5"], src=["Source"]),
}

# values that should NEVER appear in an NG column (= pollution signal)
NG_POLLUTION = re.compile(r"(?i)\b(capex|opex|capx|payroll|debit|credit|amount)\b|^\s*-?[\d,]+\.?\d*\s*$")
NG_VALID_HINT = ("NG", "博彩", "吸引", "會議", "娛樂", "體育", "文化", "健康", "主題", "美食",
                 "社區", "海上", "其他", "藝術", "養生", "遊樂", "演")


def find(df, names):
    for n in names:
        if n in df.columns: return n
    norm = {str(c).strip(): c for c in df.columns}
    for n in names:
        if str(n).strip() in norm: return norm[str(n).strip()]
    return None


def col(df, c): return df[c].astype("string").fillna("").str.strip() if c else pd.Series("", index=df.index)


def audit(L, ent, cfg):
    fp = ROOT / cfg["raw"]
    if not fp.exists():
        L.append(f"\n{'='*70}\n## {ent}: X {fp} NOT FOUND"); return
    df = pd.read_excel(fp, sheet_name=cfg["sheet"], dtype=object)
    cA, cD, cP, cN, cAmt = (find(df, cfg["ac"]), find(df, cfg["ad"]), find(df, cfg["proj"]),
                            find(df, cfg["ng"]), find(df, cfg["amt"]))
    cSrc = find(df, cfg["src"])
    a = pd.to_numeric(df[cAmt], errors="coerce").fillna(0.0) if cAmt else pd.Series(0.0, index=df.index)
    ac, ad, ng = col(df, cA), col(df, cD), col(df, cN)
    L.append(f"\n{'='*70}\n## {ent}  ({len(df):,} rows, Σ|amt|={a.abs().sum():,.0f})  cols={len(df.columns)}")
    L.append(f"   mapped: amount={cAmt!r} account_code={cA!r} account_desc={cD!r} project={cP!r} NG={cN!r} source={cSrc!r}")
    L.append(f"   ALL columns: {list(df.columns)}")

    # ── signature collapse ──
    sig = ac + "|" + ad
    both_blank = (ac.eq("") & ad.eq(""))
    L.append(f"\n   ▸ SIGNATURE health:")
    L.append(f"      distinct (account_code|account_desc) = {sig.nunique():,}  for {len(df):,} rows "
             f"(ratio {sig.nunique()/max(len(df),1)*100:.2f}%)")
    L.append(f"      account_code blank: {ac.eq('').mean()*100:5.1f}%   distinct ac={ac[ac.ne('')].nunique():,}")
    L.append(f"      account_desc blank: {ad.eq('').mean()*100:5.1f}%   distinct ad={ad[ad.ne('')].nunique():,}")
    L.append(f"      ⚠️ rows with BOTH ac+ad blank (collapse to 1 sig): {int(both_blank.sum()):,} "
             f"({a.abs()[both_blank].sum()/max(a.abs().sum(),1)*100:.1f}% of |amt|)")
    L.append(f"      zero-amount rows: {int((a==0).sum()):,}")

    # ── NG value sanity (pollution detection) ──
    ngvc = ng[ng.ne("")].value_counts()
    polluted = [v for v in ngvc.index if NG_POLLUTION.search(str(v)) or not any(h in str(v) for h in NG_VALID_HINT)]
    L.append(f"\n   ▸ NG ({cN}) health: blank {ng.eq('').mean()*100:.1f}%  distinct {ng[ng.ne('')].nunique()}")
    L.append(f"      top NG values (count):")
    for v, n in ngvc.head(25).items():
        flag = "  ⚠️POLLUTED?" if (NG_POLLUTION.search(str(v)) or not any(h in str(v) for h in NG_VALID_HINT)) else ""
        L.append(f"         {str(v)[:40]:40s} {n:>7,}{flag}")
    if polluted:
        poll_amt = a.abs()[ng.isin(polluted)].sum()
        L.append(f"      ⚠️ {len(polluted)} NG values look non-NG (capex/garbage/number) — "
                 f"{poll_amt/max(a.abs().sum(),1)*100:.1f}% of |amt|: {polluted[:12]}")

    # ── per-source breakdown ──
    if cSrc:
        L.append(f"\n   ▸ per-Source ({cSrc}):  rows | Σ|amt|% | ac_blank% | ad_blank% | NG_blank% | zero% | distinct_sig")
        for s, g in df.groupby(col(df, cSrc)):
            gi = g.index
            ga = a.abs()[gi]; gac, gad, gng = ac[gi], ad[gi], ng[gi]
            gsig = (gac + "|" + gad)
            gz = (a[gi] == 0)
            L.append(f"      {str(s)[:16]:16s} {len(g):>7,} {ga.sum()/max(a.abs().sum(),1)*100:5.1f}% "
                     f"ac{gac.eq('').mean()*100:5.0f}% ad{gad.eq('').mean()*100:5.0f}% "
                     f"NG{gng.eq('').mean()*100:5.0f}% z{gz.mean()*100:4.0f}% sig={gsig.nunique():,}")


def main():
    L = ["# audit_23_raw — data-quality of the prebuilt 2023 raws (NG pollution / sig collapse / zero rows)"]
    for ent, cfg in ENTS.items():
        try:
            audit(L, ent, cfg)
        except Exception as e:
            L.append(f"\n## {ent}: X audit failed: {e}")
    out = ROOT / "results" / "audit_23_raw.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
