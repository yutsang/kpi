r"""掃 entity 嘅 raw + parquet，揾有冇 comp/staff marking 欄係我哋無 carry（galaxy/wynn 對唔上 golden）。

user 信項目組一定 mark 好 comp/人工（有 interns 專責）。melco/vml/sjm 用 comp_type/is_labor mark 到，
但 galaxy/wynn comp_type 多數空。呢個掃 raw 各 sheet 欄名，列含 comp/房/客房/會場/票/人工/staff/內部資源
等關鍵字嘅欄 + 非空比例 + 樣本值 → 睇係咪有條 mark 欄我哋漏咗 map。

Run（Windows）:  python scripts\inspect_raw_marks.py galaxy
                 python scripts\inspect_raw_marks.py wynn
出 results\inspect_raw_marks_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = sys.argv[1] if len(sys.argv) > 1 else "galaxy"
COM = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}[ENT]
KW = ["comp", "Comp", "COMP", "房", "客房", "room", "Room", "ROOM", "hotel", "Hotel", "酒店",
      "會場", "场", "場地", "venue", "Venue", "票", "ticket", "Ticket", "贈", "門票",
      "人工", "staff", "Staff", "labor", "labour", "Labour", "薪", "payroll", "Payroll", "工資",
      "內部", "internal", "Internal", "資源", "resource", "Resource", "分攤", "ADR", "房晚",
      "免費", "in kind", "in Kind", "食品", "飲料", "F&B", "FnB"]
L: list[str] = [f"# inspect_raw_marks — {ENT} 揾 comp/staff mark 欄"]


def _hit(c) -> bool:
    s = str(c)
    return any(k in s for k in KW)


def _fill(series):
    s = series.astype(str).str.strip()
    nb = ~s.isin(["", "nan", "None", "NaN", "<NA>", "0", "0.0"])
    samp = [x for x in s[nb].unique().tolist()[:8]]
    return int(nb.sum()), len(series), samp


def _scan(tag, df):
    cols = [c for c in df.columns if _hit(c)]
    if not cols:
        return
    L.append(f"\n{'='*70}\n## {tag}  ({len(df):,} rows)")
    for c in cols:
        nb, tot, samp = _fill(df[c])
        L.append(f"   「{c}」 非空 {nb:,}/{tot:,}  例: " + " | ".join(str(x)[:24] for x in samp))


def main():
    for p in (ROOT / f"data/{ENT}/output/{COM}_kpi_report.parquet",
              ROOT / f"data/{ENT}/interim/{COM}_raw.parquet"):
        if p.exists():
            try: _scan(f"parquet {p.name}", pd.read_parquet(p))
            except Exception as e: L.append(f"\n讀 {p.name} 失敗: {e}")
    wdir = ROOT / f"data/{ENT}"
    xs = sorted([x for x in wdir.rglob("*.xlsx") if "~$" not in x.name]) if wdir.exists() else []
    for x in xs:
        try: xl = pd.ExcelFile(x)
        except Exception: continue
        for sh in xl.sheet_names:
            try: _scan(f"{x.name} [{sh}]", xl.parse(sh, nrows=3000, dtype=str))
            except Exception: pass
    out = ROOT / "results" / f"inspect_raw_marks_{ENT}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
