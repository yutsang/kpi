"""TEST: SJM combined-admin-comp GUEST_FULL_NAME → 表1 (JE) project / V correspondence.

Project team's method: each combined admin-comp GUEST_FULL_NAME appears inside 表1 'Description'.
Substring-match the guest in 表1 Description → the matched rows' project + vertical_id give the
comp's true V (the project team's own classification), replacing the unreliable keyword guess.

This script TESTS whether the match is one-to-one (each guest → exactly ONE distinct V) and writes
a detailed map results/sjm_guest_v_map.tsv. Run it, paste the headline; the inject (build_rows)
already adopts the unambiguous (1-distinct-V) matches automatically inside the pipeline.

Run:
  python scripts/sjm_guest_to_project.py
"""
from __future__ import annotations
import sys, csv, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
COM = "company_2"


def _col(df, *needles):
    for c in df.columns:
        cl = str(c).lower()
        if all(n.lower() in cl for n in needles):
            return c
    return None


def main():
    parquet = ROOT / f"data/sjm/output/{COM}_kpi_report.parquet"
    if not parquet.exists():
        print(f"X {parquet} missing — run `kedro run --pipeline=sjm` first"); return
    cfg = yaml.safe_load((ROOT / f"conf/{COM}/parameters.yml").read_text(encoding="utf-8"))
    cols = (cfg.get("columns") or {})
    df = pd.read_parquet(parquet)

    desc_col = next((c for c in df.columns if str(c).strip() == cols.get("description", "")), None) \
        or next((c for c in df.columns if str(c).strip().lower() == "description"), None) \
        or next((c for c in df.columns if "description" in str(c).lower() or "摘要" in str(c)), None)
    proj_col = next((c for c in ("Project Name", "Project", "FileName") if c in df.columns), None)
    vid_col = "vertical_id" if "vertical_id" in df.columns else None
    vlab_col = "vertical_label" if "vertical_label" in df.columns else None
    print(f"表1: rows={len(df):,}  description={desc_col!r}  project={proj_col!r}  vertical_id={vid_col!r}")
    if not desc_col or not vid_col:
        print("X cannot find Description / vertical_id column. 表1 columns =")
        print("   " + ", ".join(map(str, df.columns)))
        return

    keep = [desc_col, vid_col] + [c for c in (proj_col, vlab_col) if c]
    je = df[keep].copy()
    je[desc_col] = je[desc_col].astype(str)
    je = je[je[desc_col].str.strip().ne("") & je[desc_col].str.lower().ne("nan")].drop_duplicates()
    print(f"  unique (description, V) rows for matching: {len(je):,}")

    # ---- combined admin comp guests (25-flag=Y) ----
    ac = ROOT / "data/sjm/raw/Admin Comp summary v2.xlsx"
    if not ac.exists():
        print(f"X {ac} missing"); return
    xl = pd.ExcelFile(ac)
    comb_name = next((s for s in xl.sheet_names if "combined" in s.lower()), None)
    if not comb_name:
        print(f"X no 'combined' sheet — sheets={xl.sheet_names}"); return
    d = xl.parse(comb_name, header=1, dtype=str)
    print(f"  combined sheet {comb_name!r} columns =\n    {list(d.columns)}")
    c_guest = (_col(d, "guest_full") or _col(d, "guest") or _col(d, "full_name")
               or _col(d, "comp type") or _col(d, "event") or _col(d, "outlet") or _col(d, "name"))
    c_amt = _col(d, "amount") or _col(d, "amt")
    c_flag = _col(d, "包括在25") or _col(d, "25年")
    if not c_guest or not c_amt:
        print(f"X cannot find guest col (got {c_guest!r}) / amount col (got {c_amt!r}) — pick the real name from columns above")
        return
    if c_flag:
        d = d[d[c_flag].astype(str).str.strip().str.upper().eq("Y")]
    d["_amt"] = pd.to_numeric(d[c_amt], errors="coerce").fillna(0)
    d["_guest"] = d[c_guest].astype(str).str.strip()
    g = d.groupby("_guest")["_amt"].sum()
    gi = g.index.astype(str)                      # Index has no .ne — compare via ndarray
    g = g[(gi != "") & (gi.str.lower() != "nan")]
    tot = float(g.sum())
    print(f"  guest col = {c_guest!r}  amount col = {c_amt!r}  flag col = {c_flag!r}")
    print(f"  combined guests (25-flag=Y): {len(g):,} unique, Σ {tot:,.0f} MOP\n")

    rows, one, amb, none = [], 0, 0, 0
    a1 = aa = a0 = 0.0
    for i, (guest, amt) in enumerate(g.items(), 1):
        if i % 500 == 0:
            print(f"  ... {i}/{len(g)} guests", flush=True)
        amt = float(amt)
        if len(guest) < 3:
            none += 1; a0 += amt
            rows.append([guest, "TOO_SHORT", 0, 0, 0, "", "", "", round(amt, 2), ""]); continue
        try:
            hit = je[je[desc_col].str.contains(re.escape(guest), case=False, na=False)]
        except Exception:
            hit = je.iloc[0:0]
        vids = [v for v in hit[vid_col].astype(str).unique() if v and v != "nan"]
        projs = [p for p in (hit[proj_col].astype(str).unique() if proj_col else []) if p and p != "nan"]
        if not vids:
            none += 1; a0 += amt; status, tv, tvl, tp = "UNMATCHED", "", "", ""
        elif len(vids) == 1:
            one += 1; a1 += amt; status = "1-to-1"
            tv = vids[0]; tvl = (str(hit[vlab_col].iloc[0]) if vlab_col and len(hit) else ""); tp = projs[0] if projs else ""
        else:
            amb += 1; aa += amt; status = "AMBIGUOUS"
            tv = hit[vid_col].astype(str).value_counts().idxmax(); tvl = ""; tp = ";".join(projs[:3])
        rows.append([guest, status, len(hit), len(vids), len(projs), tv, tvl, tp, round(amt, 2), ";".join(vids[:6])])

    out = ROOT / "results"; out.mkdir(exist_ok=True)
    with (out / "sjm_guest_v_map.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["guest", "status", "n_desc_hits", "n_distinct_V", "n_projects",
                    "vertical_id", "vertical_label", "project(s)", "comp_amount", "all_vids"])
        w.writerows(sorted(rows, key=lambda r: -abs(r[8])))

    pct = lambda a: f"{a/tot*100:.1f}%" if tot else "0%"
    print("\n=== GUEST_FULL_NAME → 表1 Description match ===")
    print(f"  1-to-1    (exactly 1 V — adopted):  {one:>5} guests   {a1:>16,.0f} MOP  ({pct(a1)})")
    print(f"  AMBIGUOUS (>1 V — falls back kw):   {amb:>5} guests   {aa:>16,.0f} MOP  ({pct(aa)})")
    print(f"  UNMATCHED (0 hits — falls back kw): {none:>5} guests   {a0:>16,.0f} MOP  ({pct(a0)})")
    print(f"\n→ results/sjm_guest_v_map.tsv  (paste headline; AMBIGUOUS/large rows worth eyeballing)")


if __name__ == "__main__":
    main()
