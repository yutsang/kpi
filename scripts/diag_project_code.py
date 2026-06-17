r"""Inspect the ORIGINAL raw data books (NOT parquet) — full column inventory per entity,
so we can pick the right columns for the 4-layer project hierarchy:
   dicj_code(碼) / 項目名稱(golden名) / project_code(細碼) / project名(細名) / subproject.

parquet 經 step0 normalize 改寫過細碼 → 唔可信。直接讀 conf yearly_sources 指向嘅原始 xlsx，
逐欄報：覆蓋率 / distinct / 似碼% / 4 個樣本值，flag 邊條似 dicj碼 / 細碼 / 名。

Run:  python scripts/diag_project_code.py            # 只 2025 source
      python scripts/diag_project_code.py --all-years
出 results/diag_project_code.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys, re, argparse
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
# 細碼 pattern：SP00022 / GPGV007 / OPCE005 / A003 / B129 / J084 / IN015 / 項目16 / A1.2 / 博彩項目3
CODE = re.compile(r"^\s*((?:SP|GPGV|OPCE|CAPEX|IN|FA|PP|CE|LI|DICJ|博彩)\w*\d|[A-Z]{1,2}\d{3,}|項目\d+|[A-Z]\d[\.\-]\d+)", re.I)


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaT": ""})


def _kind(name, codepct, distinct, nrow):
    n = str(name).lower()
    if "dicj" in n or codepct >= 0.7 and distinct <= max(80, nrow * 0.01):
        return "← 似 dicj碼(粗)"
    if codepct >= 0.5 and distinct >= 40:
        return "← 似 細碼(fine)?"
    if any(k in n for k in ("name", "項目名", "名稱", "memo", "desc", "subproject", "initiative", "contents")):
        return "← 似 名"
    return ""


def inspect_sheet(L, fpath, sheet, header):
    if isinstance(sheet, list):
        sheet = sheet[0]
    try:
        df = pd.read_excel(fpath, sheet_name=sheet, header=header, dtype=str, nrows=20000)
    except Exception as e:
        L.append(f"     !! 讀唔到 {fpath.name}[{sheet}]: {e}"); return
    nrow = len(df)
    L.append(f"   {fpath.name} [{sheet}]  rows(取樣≤20k)={nrow:,}  欄數={df.shape[1]}")
    for c in df.columns:
        s = _s(df[c].astype(str)); nb = s[s != ""]
        if len(nb) == 0:
            continue
        cov = len(nb) / nrow * 100
        nd = nb.nunique()
        cm = nb.map(lambda x: bool(CODE.match(x))).mean()
        ex = list(nb.drop_duplicates().head(4))
        ex = [str(e)[:22] for e in ex]
        flag = _kind(c, cm, nd, nrow)
        L.append(f"      {str(c)[:30]:<30} cov{cov:4.0f}% dist={nd:<5} 碼{cm*100:3.0f}%  {ex}  {flag}")


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--all-years", action="store_true")
    a = ap.parse_args()
    L = ["# diag_project_code — 原始 data book 欄 inventory (揾 4 層 project 碼/名來源)"]
    for ent, n in ENT.items():
        confp = ROOT / "conf" / f"company_{n}" / "parameters.yml"
        L.append(f"\n{'='*72}\n── {ent} (company_{n}) ──")
        if not confp.exists():
            L.append(f"   !! conf 揾唔到 {confp.relative_to(ROOT)}"); continue
        cfg = yaml.safe_load(confp.read_text(encoding="utf-8")) or {}
        rawdir = ROOT / "data" / ent / "raw"
        ys = cfg.get("yearly_sources") or [{"year": 2025,
                                            "raw_filename": cfg.get("raw_filename"),
                                            "raw_sheet": cfg.get("raw_sheet", 0)}]
        # conf columns map (話我知邊條係 project/account 等)
        cm = cfg.get("columns") or {}
        L.append(f"   conf columns: project={cm.get('project')!r} account={cm.get('account_code')!r} "
                 f"unique_id={cm.get('unique_id')!r}")
        for y in ys:
            if not a.all_years and int(y.get("year", 0)) != 2025:
                continue
            fp = rawdir / (y.get("raw_filename") or "")
            sheet = y.get("raw_sheet", 0)
            hdr = y.get("header_row", cfg.get("header_row", 0))
            L.append(f"   ─ year {y.get('year')} ─ header_row={hdr}")
            if not fp.exists():
                L.append(f"     !! 原始檔揾唔到 {fp.relative_to(ROOT)}"); continue
            inspect_sheet(L, fp, sheet, hdr)
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_project_code.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
