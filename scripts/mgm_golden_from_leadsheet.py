"""MGM 2023 (+24) golden TSV from the project-team 'Leadsheet' (de-facto golden — MGM has no
formal golden TSV, but Leadsheet states per-項目 OPEX / CAPEX / Payroll / total, authoritative).

Why: MGM gaming capex (~1.337B) lives in a PIVOT with no Project Plan Task → cannot be rebuilt
bottom-up (same reason 2025 went golden-driven). So we take the Leadsheet per-項目 amounts as
truth and let build_mgm_golden.py distribute them by the pipeline's V/NG + H proportions.

Golden TSV layout (build_mgm_golden expects): 序號 \t 名稱 \t payroll \t capex \t opex \t total
(amounts in 萬元 → build_mgm_golden ×10000 = MOP). 序號 rows are all-digits ('項目X' → strip).

Run (Windows):
  python scripts/mgm_golden_from_leadsheet.py --year 2023
Output → results/mgm_golden_23.tsv
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# per-year: which file + sheet holds the Leadsheet, and column keyword sets
SRC = {
    "2023": dict(file="OPEX.xlsx", sheet="Leadsheet", header=0),
    "2024": dict(file=None, sheet=None, header=0),   # 2024 has no clean Leadsheet — TBD
}


def find_file(name, year):
    if not name: return None
    for p in [ROOT / "data" / "mgm" / "raw" / year / name, ROOT / "data" / "mgm" / "raw" / name]:
        if p.exists(): return p
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def pick(df, *kwsets):
    """first column whose lowered name contains ALL keywords of any kwset."""
    for c in df.columns:
        cl = str(c).lower()
        for kws in kwsets:
            if all(k.lower() in cl for k in kws):
                return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce").fillna(0.0)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", default="2023", choices=["2023", "2024"])
    a = ap.parse_args()
    cfg = SRC[a.year]
    fp = find_file(cfg["file"], a.year)
    if not fp:
        print(f"X {cfg['file']} not found under data/mgm/raw/{a.year}/ — 2024 Leadsheet TBD" if a.year == "2024"
              else f"X {cfg['file']} not found under data/mgm/raw/{a.year}/")
        return
    df = pd.read_excel(fp, sheet_name=cfg["sheet"], header=cfg["header"], dtype=object)
    print(f"Leadsheet {fp.name}!{cfg['sheet']}  rows={len(df)}  cols={len(df.columns)}")

    c_no   = pick(df, ["承批公司項目序號"], ["項目序號"], ["序號"])
    c_name = pick(df, ["項目名稱"])
    c_opex = pick(df, ["opex", "萬"], ["opex金額", "萬"])
    c_capx = pick(df, ["capex", "萬"], ["capex金額", "萬"])
    c_pay  = pick(df, ["payroll", "萬"], ["人工", "萬"])
    c_tot  = pick(df, ["总金额"], ["總金額"], ["实际投资金额", "总"], ["實際投資金額"])
    print(f"  cols → 序號={c_no!r} 名稱={c_name!r} opex={c_opex!r} capex={c_capx!r} payroll={c_pay!r} total={c_tot!r}")
    if not (c_no and c_opex and c_capx):
        print("  X missing key columns — inspect Leadsheet header"); return

    def seq(v):
        m = re.search(r"\d+", str(v))
        return m.group(0) if m else ""

    _pay = num(df[c_pay]) if c_pay else pd.Series(0.0, index=df.index)
    _cap = num(df[c_capx])
    _opx = num(df[c_opex])
    c_nat = pick(df, ["項目性質"])
    nat = df[c_nat].astype(str) if c_nat else pd.Series("", index=df.index)
    # 博彩 vs 非博彩 both reuse 項目1..10 → 博彩 序號 +200 (keeps pure-digit for parse_golden;
    # build_mgm_23_raw gaming tabs use 項目201..210 to match).
    def osq(no, n):
        d = seq(no)
        return (str(int(d) + 200) if "博彩" in str(n) else d) if d else ""
    seqs = [osq(no, n) for no, n in zip(df[c_no], nat)]
    # NOTE: '总金额' col is buggy (= payroll + capex + 2×opex). Derive total = p+c+o.
    out = pd.DataFrame({
        "序號": seqs,
        "名稱": (df[c_name].astype(str).str.replace(r"[\r\n]+", " ", regex=True).str.strip().values if c_name else ""),
        "payroll": _pay.values,
        "capex": _cap.values,
        "opex": _opx.values,
        "total": (_pay + _cap + _opx).values,
        # 7th col = 項目性質 (project-team theme) → build_mgm_golden maps it to V/NG directly,
        # bypassing the pipeline V_OTHER corruption (CAPEX/payroll rows have blank NG → V_OTHER).
        "theme": (nat.str.replace(r"[\r\n]+", " ", regex=True).str.strip().values if c_nat else ""),
    })
    out = out[out["序號"].str.contains(r"\d", na=False)]   # 序號 must have a digit (drops 合計/blank rows)
    # drop subtotal/total rows (名稱 contains 合計/小計/總)
    out = out[~out["名稱"].str.contains("合計|小計|總計|total", case=False, na=False)]
    res = ROOT / "results"; res.mkdir(exist_ok=True)
    fpo = res / f"mgm_golden_{a.year[2:]}.tsv"
    out.to_csv(fpo, sep="\t", index=False, header=False, encoding="utf-8-sig")
    print(f"\n✓ {len(out)} 項目 → {fpo.relative_to(ROOT)}  (7 cols incl 項目性質 theme)")
    print(f"  Σ(萬元) payroll={out['payroll'].sum():,.0f}  capex={out['capex'].sum():,.0f}  "
          f"opex={out['opex'].sum():,.0f}  total={out['total'].sum():,.0f}")
    print("  (expect ~ payroll 4,601 / capex 32,685 / opex 96,583 / total 140,418 萬)")
    print("\n  項目性質 (theme) distinct values — confirm these are byte-exact keys in build_mgm_golden THEME2V:")
    for v, n in out["theme"].astype(str).str.strip().replace("", "(blank)").value_counts().items():
        print(f"    {str(v)[:32]:32s} {n:>3} 項目")


if __name__ == "__main__":
    main()
