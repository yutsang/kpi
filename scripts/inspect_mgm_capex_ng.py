"""MGM 23 — find a Project Code → NG bridge for the non-gaming CAPEX (currently 未分類).

build_mgm_23_raw.py reads CAPEX.xlsx 'JL details' keyed by raw 'Project Code' (855000-23030C…),
which is NOT in the Leadsheet NG map (keyed by 承批公司項目序號 項目N) → Section.1 blank → 14,973
rows / 2.65B 未分類. To assign NG we need a Project Code → 項目/NG link. This dumps:
  (A) every column in CAPEX.xlsx 'JL details' + non-blank% + sample distinct values — to spot any
      項目序號 / Section / 範疇 / 性質 / Department / Cost Center col that maps to NG
  (B) the Leadsheet 承批公司項目序號 → 項目性質 map (so we see the NG vocabulary + whether any CAPEX
      column value could join to it)

Run (Windows):  python scripts/inspect_mgm_capex_ng.py
Output: prints + results/inspect_mgm_capex_ng.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
HINTS = ["項目", "序號", "Section", "範疇", "性質", "NG", "Nature", "Department", "Cost Center",
         "Cost Centre", "Centre", "WBS", "Worktag", "Program", "Initiative", "Scheme"]


def find_file(name):
    d = ROOT / "data"
    hits = list(d.rglob(name)) + list(d.rglob(f"*{name}*"))
    return hits[0] if hits else None


def dump_sheet(L, fp, sheet):
    try:
        df = pd.read_excel(fp, sheet_name=sheet, header=0, dtype=object)
    except Exception as e:
        L.append(f"   read {sheet} failed: {e}"); return
    L.append(f"\n   sheet '{sheet}': {len(df):,} rows, {len(df.columns)} cols")
    for c in df.columns:
        s = df[c].astype("string").fillna("").str.strip()
        nb = s.ne("").mean() * 100
        nun = s[s.ne("")].nunique()
        top = " | ".join(map(str, s[s.ne("")].value_counts().head(4).index))
        flag = " <== NG-bridge?" if any(h.lower() in str(c).lower() for h in HINTS) else ""
        L.append(f"      {str(c)[:34]:34s} nb{nb:4.0f}% uniq{nun:>6}  {top[:70]}{flag}")


def main():
    L = ["# inspect_mgm_capex_ng — find Project Code → NG bridge for non-gaming CAPEX"]

    cap = find_file("CAPEX.xlsx")
    L.append(f"\n## (A) CAPEX.xlsx = {cap.relative_to(ROOT) if cap else 'NOT FOUND'}")
    if cap:
        try:
            xls = pd.ExcelFile(cap)
            L.append(f"   tabs: {xls.sheet_names}")
        except Exception as e:
            L.append(f"   open failed: {e}"); xls = None
        if xls:
            for sh in xls.sheet_names:
                dump_sheet(L, cap, sh)

    # Leadsheet — the file with 承批公司項目序號 + 項目性質
    lead = None
    for nm in ("Leadsheet.xlsx", "leadsheet.xlsx", "Lead Sheet.xlsx", "MGM-gaming.xlsx"):
        lead = find_file(nm)
        if lead:
            break
    L.append(f"\n## (B) Leadsheet candidate = {lead.relative_to(ROOT) if lead else 'NOT FOUND (search manually)'}")
    if lead:
        try:
            xls = pd.ExcelFile(lead)
            for sh in xls.sheet_names[:3]:
                df = pd.read_excel(lead, sheet_name=sh, header=0, dtype=object)
                cols = [c for c in df.columns if any(k in str(c) for k in ("序號", "項目性質", "項目", "Section"))]
                if cols:
                    L.append(f"   sheet '{sh}' NG cols: {cols}")
                    no = next((c for c in df.columns if "序號" in str(c)), None)
                    nat = next((c for c in df.columns if "項目性質" in str(c)), None)
                    if no and nat:
                        for _, r in df[[no, nat]].dropna().head(15).iterrows():
                            L.append(f"      {str(r[no])[:20]:20s} -> {str(r[nat])[:24]}")
        except Exception as e:
            L.append(f"   read failed: {e}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_mgm_capex_ng.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
