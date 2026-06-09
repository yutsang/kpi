"""Cross-entity master-audit column line-up — what's populated vs blank, per entity × year.

The 大表 (build_master_audit JE_KEEP) maps each entity's raw cols to canonical names
(account_code / account_desc / description / vendor / project / ng11_category). But each entity
maps a DIFFERENT raw column, and some end up blank (e.g. SJM 23 description is blank). This loops
all 6 entities, reads each one's conf column mapping + tagged_rows, and reports, per year bucket,
the mapped raw col name + non-blank% + distinct + a sample — so we can see what lines up and what's
missing across 各家各年, and fix the mapping where a column is empty.

Run (Windows):  python scripts/inspect_master_columns.py
Output: prints + results/inspect_master_columns.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
try: import yaml
except Exception: yaml = None

ROOT = Path(__file__).resolve().parent.parent
KEYS = ["project", "ng11_category", "account_code", "account_desc", "description", "vendor"]
ENTITIES = [(1, "galaxy"), (2, "sjm"), (3, "wynn"), (4, "vml"), (5, "melco"), (6, "mgm")]


def _fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    for c in df.columns:
        if str(name).strip() and str(name).strip() in str(c): return c
    return None


def main():
    L = ["# inspect_master_columns — cross-entity 大表 column line-up (per entity × year)"]
    for code, alias in ENTITIES:
        confp = ROOT / "conf" / f"company_{code}" / "parameters.yml"
        trp = ROOT / "data" / alias / "interim" / f"company_{code}_tagged_rows.parquet"
        L.append(f"\n{'='*78}\n## {alias} (company_{code})")
        if not confp.exists() or yaml is None:
            L.append(f"   conf missing or no yaml: {confp}"); continue
        cfg = yaml.safe_load(confp.read_text(encoding="utf-8")) or {}
        cols = cfg.get("columns", {}) or {}
        # per-year ng11_category overrides
        yov = {str(ys.get("year")): (ys.get("columns_override") or {}) for ys in (cfg.get("yearly_sources") or [])}
        if not trp.exists():
            L.append(f"   tagged_rows missing: {trp}"); continue
        df = pd.read_parquet(trp)
        per = next((c for c in ("report_period", "report_year") if c in df.columns), None)
        buckets = sorted(df[per].astype(str).unique()) if per else ["(all)"]
        L.append(f"   tagged_rows rows={len(df):,}  buckets={buckets}")
        for key in KEYS:
            base = cols.get(key, "")
            raw = _fuzzy(df, base)
            L.append(f"   • {key:14s} conf={str(base)[:24]:24s} -> col={str(raw)[:24]:24s}")
            if not raw:
                L.append(f"       !! NOT FOUND in tagged_rows"); continue
            s = df[raw].astype("string").fillna("").str.strip()
            for b in buckets:
                m = (df[per].astype(str) == b) if per else pd.Series(True, index=df.index)
                sb = s[m]
                nb = sb.ne("").mean() * 100 if len(sb) else 0
                samp = " | ".join(map(str, sb[sb.ne("")].value_counts().head(2).index))
                flag = "  <== BLANK" if nb < 5 else ""
                L.append(f"       {b:9s} nb{nb:5.1f}% uniq{sb[sb.ne('')].nunique():>6}  {samp[:46]}{flag}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_master_columns.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
