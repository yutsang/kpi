r"""Find WHICH raw column holds the DICJ Code in each entity — by VALUE overlap with the golden DICJ
set (robust to the column being named differently per entity/year). Then we map that column to root
`dicj_code` (conf columns.dicj_code + yearly_sources overrides) so it carries → kpi_report → Tableau.

Scans data/<ent>/interim/<com>_raw.parquet (all raw cols, pre-trim). Reports, per entity, the top
columns ranked by how many golden DICJ codes their values match (exact + 項目-zero-normalised), so
we can confirm 'most raw files have a dicj code' and wire it.

Run (Windows):  python scripts/inspect_dicj_cols.py
Output: prints + results/inspect_dicj_cols.txt  (paste back). Golden read-only, never committed.
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
GCAND = [ROOT / "HQ 投資方向_audit_0616.xlsx", ROOT / "data" / "HQ 投資方向_audit_0616.xlsx",
         ROOT / "results" / "HQ 投資方向_audit_0616.xlsx"]
TAB = "Database combine"
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
NAME2ALIAS = [
    (["galaxy", "銀河"], "galaxy"), (["wynn", "永利"], "wynn"), (["mgm", "美高梅"], "mgm"),
    (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams"], "melco"),
    (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
    (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "vml", "venetian"], "vml"),
]


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws):
            return a
    return None


def _col(df, cands):
    for c in cands:
        if c in df.columns:
            return c
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if str(c).strip().lower() in low:
            return low[str(c).strip().lower()]
    return None


def _ndicj(s):
    """項目001→項目1; strip spaces; keep code shape (A1.1 / CA001 unchanged)."""
    s = str(s).strip()
    m = re.match(r"^(項目)0*(\d.*)$", s)
    return m.group(1) + m.group(2) if m else s


def main():
    L = ["# inspect_dicj_cols — 反找 raw 邊條欄裝 DICJ (by value overlap w/ golden)"]
    gpath = next((p for p in GCAND if p.exists()), None)
    if not gpath:
        L.append("!! golden 揾唔到"); _w(L); return
    g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    c_ent, c_dicj = _col(g, ["承批公司"]), _col(g, ["DICJ Code", "DICJ"])
    g["_a"] = g[c_ent].map(_alias)

    for alias, com in ENT.items():
        L.append(f"\n{'='*66}\n## {alias}")
        gd_raw = set(str(d).strip() for d in g[g["_a"] == alias][c_dicj].dropna().unique() if str(d).strip())
        gd_norm = set(_ndicj(d) for d in gd_raw)
        L.append(f"   golden DICJ: {len(gd_raw)} codes  sample={sorted(gd_raw)[:5]}")
        p = ROOT / "data" / alias / "interim" / f"{com}_raw.parquet"
        if not p.exists():
            L.append(f"   (no raw parquet: {p.name})"); continue
        df = pd.read_parquet(p, columns=None)
        n = len(df)
        hits = []
        for col in df.columns:
            s = df[col].astype("string").fillna("").str.strip()
            nb = int(s.ne("").sum())
            if nb == 0:
                continue
            vals = set(v for v in s.unique() if v)
            ov_raw = len(vals & gd_raw)
            ov_norm = len(set(_ndicj(v) for v in vals) & gd_norm)
            ov = max(ov_raw, ov_norm)
            if ov >= 2:                                  # at least 2 golden codes appear in this col
                hits.append((ov, ov_raw, ov_norm, col, nb, len(vals)))
        hits.sort(reverse=True)
        if not hits:
            L.append("   ▶ 冇任何欄 match golden DICJ → raw 真係冇 DICJ（要靠名 fill, build_dicj_fill）")
        else:
            L.append(f"   ▶ 候選 DICJ 欄 (match golden codes / {len(gd_raw)}):")
            for ov, ovr, ovn, col, nb, nv in hits[:5]:
                tag = "" if ovr >= ovn else " (項目零-normalise 後先 match → 我哋值有前導零/格式差)"
                L.append(f"      {str(col)[:32]:32s} match={ov:>4}  填{nb:>8,}/{n:,} ({nb/n*100:.0f}%)  {nv} distinct{tag}")
            best = hits[0][3]
            L.append(f"   → 建議 map columns.dicj_code = {best!r}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_dicj_cols.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
