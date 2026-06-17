r"""Diagnose project_code quality + LOCATE each entity's true fine-grained sub-code.

項目組要 4 層：dicj_code(碼) / 項目名稱(golden名) / project_code(細碼) / project名(細名).
而家 project_code 亂：有 None、有等同 dicj_code(roll-up 漏)、真細碼(SP00022/A003)時常埋喺個名。

掃每家 kpi_report.parquet 全部欄，揾邊條最似「細碼」(高覆蓋、高 distinct、唔等於 dicj、合 code pattern)，
同埋 project/subproject 名開頭有冇可抽嘅 leading code → 我據此 wire 返 prep_tableau 嘅 project_code。

Run:  python scripts/diag_project_code.py
出 results/diag_project_code.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
# 細碼 pattern：SP00022 / GPGV007 / OPCE005 / A003 / B129 / J084 / IN015 / 項目16 / A1.2
CODE = re.compile(r"^\s*((?:SP|GPGV|OPCE|CAPEX|IN|FA|PP|CE|LI|DICJ)\d{2,}|[A-Z]{1,2}\d{3,}|項目\d+|[A-Z]\d[\.\-]\d+)", re.I)
NAMEISH = ("name", "項目", "project", "subproject", "memo", "desc", "名", "initiative", "contents")


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": ""})


def main():
    L = ["# diag_project_code — project_code 品質 + 揾每家真細碼來源欄"]
    for ent, n in ENT.items():
        p = ROOT / "data" / ent / "output" / f"company_{n}_kpi_report.parquet"
        L.append(f"\n{'='*70}\n── {ent} (company_{n}) ──")
        if not p.exists():
            L.append(f"   !! {p.relative_to(ROOT)} 揾唔到 (未跑?)"); continue
        df = pd.read_parquet(p)
        nrow = len(df)
        dc = _s(df["dicj_code"]) if "dicj_code" in df.columns else pd.Series("", index=df.index)
        L.append(f"   rows={nrow:,}   dicj_code distinct={dc[dc!=''].nunique():,} blank={ (dc=='').mean()*100:.1f}%")
        # 現有 project_code 狀態
        if "project_code" in df.columns:
            pc = _s(df["project_code"])
            L.append(f"   [現有 project_code] blank={(pc=='').mean()*100:4.1f}%  ==dicj={((pc!='')&(pc==dc)).mean()*100:4.1f}%  distinct={pc[pc!=''].nunique():,}  e.g. {list(pc[pc!=''].drop_duplicates().head(5))}")
        else:
            L.append("   [現有 project_code] 冇呢欄 → prep_tableau 會用 project名/dicj fallback")
        # 掃全部欄揾細碼候選
        L.append("   ── 細碼候選欄 (code-match≥40% & distinct≥20 & 唔淨係 dicj) ──")
        cand = []
        for c in df.columns:
            s = _s(df[c].astype(str))
            nb = s[s != ""]
            if len(nb) < nrow * 0.2:
                continue
            cm = nb.map(lambda x: bool(CODE.match(x))).mean()
            nd = nb.nunique()
            eqd = (nb.reset_index(drop=True) == dc[s != ''].reset_index(drop=True)).mean() if (s != '').any() else 0
            if cm >= 0.40 and nd >= 20:
                cand.append((c, cm, nd, eqd, list(nb.drop_duplicates().head(4))))
        cand.sort(key=lambda t: (-t[1], -t[2]))
        for c, cm, nd, eqd, ex in cand[:8]:
            L.append(f"      {str(c)[:34]:<34} code{cm*100:4.0f}%  distinct={nd:<5} ==dicj{eqd*100:3.0f}%  e.g.{ex}")
        if not cand:
            L.append("      (冇明顯細碼欄 → 細碼可能埋喺名)")
        # 名開頭可抽 leading code？
        L.append("   ── 名開頭有 leading code (可 regex 抽碼) ──")
        for c in df.columns:
            if not any(k in str(c).lower() for k in NAMEISH):
                continue
            s = _s(df[c].astype(str)); nb = s[s != ""]
            if len(nb) < nrow * 0.2:
                continue
            hit = nb.map(lambda x: bool(CODE.match(x)))
            if hit.mean() >= 0.30:
                ex = [CODE.match(x).group(1) for x in nb[hit].drop_duplicates().head(5)]
                L.append(f"      {str(c)[:34]:<34} 開頭碼 {hit.mean()*100:4.0f}%   抽到 e.g. {ex}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_project_code.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
