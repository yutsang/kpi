"""Dump 2023 raw rows for the baseline classification (sjm 2 tabs + melco 明細賬).

For each source: writes a PROJECTS tsv (for V baseline + 項目組 compare) and a SIGS tsv
(account_code|account_desc for H baseline; carries 項目組 性質/支出性质 to compare).

  python scripts/dump_raw_23.py            # all sources
  python scripts/dump_raw_23.py --src melco
Output → results/{src}_projects_23.tsv + results/{src}_sigs_23.tsv (paste / drop to results/).
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# col map per source (header_row=0 for all 3). h = 項目組 H label col (None = no project-team H).
SOURCES = {
    "sjm_capex": dict(file="sjm_2023.xlsx", sheet="項目明細賬capex", header=0,
                      amt="Val/COArea Crcy", proj="Project Code with Project Name",
                      ac="Cost Element", ad="Cost element descr.", v="項目性質", h=None, capex="Capex"),
    "sjm_opex":  dict(file="sjm_2023.xlsx", sheet="項目明細賬opex", header=0,
                      amt="Amount in local currency", proj="项目名称",
                      ac="G/L Account", ad="GL account description", v="項目性質", h=None, capex="Opex"),
    "melco":     dict(file="melco_2023.xlsx", sheet="明細賬", header=0,
                      amt="本位币金额", proj="項目名稱",
                      ac="ledger_account", ad="spend_category", v="項目性質", h="支出性质-mapping", capex=None),
    "wynn":      dict(file="wynn_2023.xlsx", sheet="Capex and Opex summary", header=0,
                      amt="Entry Voucher Amount/ Expense Amount ", proj="项目名称中文",
                      ac="Account", ad="Nature of Expenses", v="項目性質", h="comp费用大类", capex="Capex/Opex"),
}


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    hits = list((ROOT / "data").glob(f"*/raw/{name}")) + list((ROOT / "data").glob(f"*/raw/*{name}*"))
    return hits[0] if hits else None


def _col(df, name):
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def dump(src, cfg):
    fp = find_file(cfg["file"])
    if not fp:
        print(f"[{src}] X {cfg['file']} not found"); return
    df = pd.read_excel(fp, sheet_name=cfg["sheet"], header=cfg["header"], dtype=object)
    amt = _col(df, cfg["amt"]); proj = _col(df, cfg["proj"]); v = _col(df, cfg["v"])
    ac = _col(df, cfg["ac"]); ad = _col(df, cfg["ad"]); h = _col(df, cfg["h"]) if cfg["h"] else None
    if not amt:
        print(f"[{src}] X amount col {cfg['amt']!r} missing — cols: {list(df.columns)[:15]}"); return
    df["_amt"] = pd.to_numeric(df[amt], errors="coerce").fillna(0.0)
    res = ROOT / "results"; res.mkdir(exist_ok=True)

    # PROJECTS (V baseline + 項目組 性質 compare)
    if proj and v:
        g = (df.assign(_p=df[proj].astype(str).str.strip().str.replace(r"[\r\n]+", " ", regex=True),
                       _v=df[v].astype(str).str.strip())
               .groupby(["_p", "_v"])["_amt"].agg(amount="sum", n="size").reset_index()
               .rename(columns={"_p": "project", "_v": "項目組_性質"}).sort_values("amount", ascending=False, key=abs))
        g.to_csv(res / f"{src}_projects_23.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"[{src}] ✓ {len(g)} projects → results/{src}_projects_23.tsv  (Σ={df['_amt'].sum():,.0f})")

    # SIGS (H baseline; carry 項目組 性質 + 支出性质 to compare)
    if ac or ad:
        df["_sig"] = (df[ac].astype(str).str.strip() if ac else "") + "|" + (df[ad].astype(str).str.strip() if ad else "")
        gb = ["_sig"]
        agg = {"_amt": ["sum", "size"]}
        sub = df.groupby("_sig")
        out = sub["_amt"].agg(amount="sum", n="size").reset_index()
        # dominant 項目組 H (支出性质) + 性質 per sig
        if h:
            out["項目組_支出性质"] = sub[h].agg(lambda s: s.astype(str).str.strip().mode().iloc[0] if not s.dropna().empty else "").values
        if v:
            out["項目組_性質_mix"] = sub[v].agg(lambda s: " / ".join(s.astype(str).str.strip().value_counts().head(2).index)).values
        out = out.sort_values("amount", ascending=False, key=abs)
        out.to_csv(res / f"{src}_sigs_23.tsv", sep="\t", index=False, encoding="utf-8-sig")
        print(f"[{src}] ✓ {len(out)} sigs → results/{src}_sigs_23.tsv")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", choices=list(SOURCES) + ["all"], default="all")
    a = ap.parse_args()
    for s, c in SOURCES.items():
        if a.src in ("all", s):
            dump(s, c)
