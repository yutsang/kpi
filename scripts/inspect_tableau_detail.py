"""Diagnose why the 7 uniform detail cols (科目層級/科目明細/發票號/PO號/成本中心/WBS子項/憑證號) are
blank in tableau_combined — is it (A) conf audit_detail_cols not present on this machine, or (B) the
raw source columns are missing from kpi_report.parquet (so prep_tableau / build_master_audit can't
pull them)?

For each entity: read conf audit_detail_cols + the kpi_report.parquet, and for every mapped raw col
report whether it EXISTS in kpi_report and its non-blank%.

Run (Windows):  python scripts/inspect_tableau_detail.py
Output: prints + results/inspect_tableau_detail.txt
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
ENTITIES = [(1, "galaxy"), (2, "sjm"), (3, "wynn"), (4, "vml"), (5, "melco"), (6, "mgm")]
DETAIL = ["科目層級", "科目明細", "發票號", "PO號", "成本中心", "WBS子項", "憑證號"]


def _fuzzy(df, name):
    if not name: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def main():
    L = ["# inspect_tableau_detail — why detail cols blank? (conf missing vs raw col missing)"]
    for code, alias in ENTITIES:
        confp = ROOT / "conf" / f"company_{code}" / "parameters.yml"
        kp = ROOT / "data" / alias / "output" / f"company_{code}_kpi_report.parquet"
        L.append(f"\n{'='*72}\n## {alias} (company_{code})")
        if not confp.exists() or yaml is None:
            L.append("   conf missing / no yaml"); continue
        cfg = yaml.safe_load(confp.read_text(encoding="utf-8")) or {}
        adc = cfg.get("audit_detail_cols") or {}
        L.append(f"   conf audit_detail_cols: {'PRESENT ('+str(len(adc))+' keys)' if adc else '!!! MISSING — conf not copied here'}")
        if not kp.exists():
            L.append(f"   kpi_report missing: {kp}"); continue
        df = pd.read_parquet(kp)
        L.append(f"   kpi_report rows={len(df):,}  cols={len(df.columns)}")
        for name in DETAIL:
            raws = adc.get(name, "")
            raws = raws if isinstance(raws, list) else ([raws] if raws else [])
            if not raws:
                L.append(f"      {name:6s} — (not mapped in conf)"); continue
            parts = []
            for r in raws:
                src = _fuzzy(df, r)
                if src:
                    nb = df[src].astype("string").fillna("").str.strip().ne("").mean() * 100
                    parts.append(f"{r}=✓{nb:.0f}%")
                else:
                    parts.append(f"{r}=✗missing")
            L.append(f"      {name:6s} ← {' | '.join(parts)}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_tableau_detail.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
