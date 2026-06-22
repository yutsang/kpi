r"""inspect_golden_gaming_names —— 揾「同一 DICJ 嘅 博彩 vs 非博彩 係兩個唔同項目名」嘅 case。

懷疑：golden_name_map 用 dicj-idxmax（攞最大額名）。一個 DICJ 如果博彩(NG0,capex巨大)+非博彩
都有，idxmax 攞咗博彩個名 → 非博彩行(如 NG1 路演)冚錯博彩項目名。
查 golden「Database combine」：per (承批公司, DICJ Code) × 是否博彩 → 項目名稱，列出兩邊名唔同嘅。

Run（Windows）:  python scripts\inspect_golden_gaming_names.py
出 results\inspect_golden_gaming_names.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "inspect_golden_gaming_names.txt"
GCAND = [ROOT / "HQ 投資方向_audit_0616.xlsx", ROOT / "data" / "HQ 投資方向_audit_0616.xlsx"]


def main():
    L = ["# inspect_golden_gaming_names —— 同一DICJ 博彩vs非博彩 唔同名 (idxmax 回填錯源)"]
    gp = next((p for p in GCAND if p.exists()), None) or next(iter(ROOT.rglob("HQ 投資方向*.xlsx")), None)
    if not gp:
        L.append("  ⚠ golden 揾唔到"); _w(L); return
    g = pd.read_excel(gp, sheet_name="Database combine", dtype=str)
    g.columns = [str(c).strip() for c in g.columns]
    gcol = next((c for c in g.columns if "是否博彩" in c), None)
    dcol = next((c for c in g.columns if c.strip() in ("DICJ Code", "DICJ")), None)
    ncol = next((c for c in g.columns if "項目名稱" in c or "项目名称" in c), None)
    acomp = next((c for c in g.columns if "承批" in c), None)
    acol = next((c for c in g.columns if c.strip() == "Amount"), None)
    L.append(f"  欄: 是否博彩={gcol!r} DICJ={dcol!r} 名={ncol!r} 承批={acomp!r}")
    if not (gcol and dcol and ncol and acomp):
        L.append("  ⚠ 缺欄"); _w(L); return
    g["_g"] = g[gcol].astype(str).str.strip()
    g["_d"] = g[dcol].astype(str).str.strip()
    g["_n"] = g[ncol].astype(str).str.strip()
    g["_a"] = g[acomp].astype(str).str.strip()
    g["_amt"] = pd.to_numeric(g[acol].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if acol else 0.0

    L.append(f"\n  是否博彩 distinct values: {sorted(g['_g'].dropna().unique())[:8]}")

    # per (承批, DICJ)：是否博彩 → 名
    n_conflict = 0
    for (a, d), sub in g.groupby(["_a", "_d"]):
        names_by_g = sub.groupby("_g")["_n"].agg(lambda s: s.value_counts().index[0] if len(s) else "")
        distinct = set(v for v in names_by_g.values if str(v).strip() not in ("", "nan"))
        if len(distinct) > 1:   # 同 DICJ 唔同名
            n_conflict += 1
            if n_conflict <= 40:
                idxmax_name = sub.loc[sub["_amt"].abs().idxmax(), "_n"]   # 而家回填會攞呢個
                L.append(f"\n  {a[:14]:<14} DICJ={d[:12]:<12}  (idxmax回填→「{str(idxmax_name)[:24]}」)")
                for gv, sg in sub.groupby("_g"):
                    nm = sg.groupby("_n")["_amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
                    top = " | ".join(f"{str(k)[:22]}({v:,.0f})" for k, v in nm.head(2).items())
                    L.append(f"       是否博彩={str(gv)[:8]:<8}: {top}")
    L.append(f"\n  >>> 同DICJ 博彩/非博彩 唔同名 嘅 case 總數 = {n_conflict}（呢啲非博彩行而家可能冚咗博彩名）")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
