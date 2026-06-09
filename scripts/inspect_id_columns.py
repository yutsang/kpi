"""Cross-entity identifying-column menu — what extra row-level detail each entity can expose so the
project team can identify 'this transaction' in the deliverable (24/25 delivery years).

The 大表 today carries account_code / account_desc / description / vendor. The project team wants
ALL 6 entities fully identifiable, and where a source has MULTIPLE layers (e.g. MGM Ledger Hierarchy
L4 + L5 + Spend Category + Invoice# + PO + Cost Center), to split them out too. Those columns already
live in tagged_rows. This lists, per entity × {24,25}, every IDENTIFYING-candidate column (by name
keyword) with non-blank% + distinct + sample — the menu to build a per-entity `audit_extra_cols`.

Run (Windows):  python scripts/inspect_id_columns.py
Output: prints + results/inspect_id_columns.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = [(1, "galaxy"), (2, "sjm"), (3, "wynn"), (4, "vml"), (5, "melco"), (6, "mgm")]
# column-name keywords that mark an "identifying" detail column
ID_KW = ["account", "acct", "code", "desc", "descr", "name", "memo", "narr", "摘要", "說明", "說 明",
         "invoice", "發票", "po ", "purchase", "order", "vendor", "supplier", "供應", "journal",
         "wbs", "cost cent", "cost cent", "centre", "hierarch", "level", "spend", "category", "worktag",
         "project", "項目", "item", "type", "性質", "reference", "ref", "document", "憑證", "唯一",
         "unique", "identifier", "序號", "ledger", "section", "nature", "element", "object", "ac1"]
EXCLUDE_KW = ["amount", "debit", "credit", "金額", "val/", "crcy", "currency", "date", "year",
              "period", "ratio", "split_", "_row", "_uid", "vertical", "horizontal", "ng_scope",
              "row_type", "report_", "是否", "抽憑", "net-off", "原始", "original_amount", "絕對值",
              "fiscal", "posting", "tran", "translation", "balance", "fullpath", "folder", "sheet name"]


def _idcand(c):
    s = str(c).lower()
    if any(k in s for k in EXCLUDE_KW):
        return False
    return any(k.strip() in s for k in ID_KW)


def main():
    L = ["# inspect_id_columns — identifying-column menu per entity × {24,25} (for audit_extra_cols)"]
    for code, alias in ENTITIES:
        trp = ROOT / "data" / alias / "interim" / f"company_{code}_tagged_rows.parquet"
        L.append(f"\n{'='*80}\n## {alias} (company_{code})")
        if not trp.exists():
            L.append(f"   tagged_rows missing: {trp}"); continue
        df = pd.read_parquet(trp)
        per = next((c for c in ("report_period", "report_year") if c in df.columns), None)
        delv = df[df[per].astype(str).isin(["24", "25"])] if per else df
        L.append(f"   rows={len(df):,}  24/25 rows={len(delv):,}  total cols={len(df.columns)}")
        cands = [c for c in df.columns if _idcand(c)]
        L.append(f"   identifying-candidate cols ({len(cands)}):")
        for c in cands:
            s = delv[c].astype("string").fillna("").str.strip()
            nb = s.ne("").mean() * 100 if len(s) else 0
            if nb < 1:
                continue
            nun = s[s.ne("")].nunique()
            samp = " | ".join(map(str, s[s.ne("")].value_counts().head(2).index))
            L.append(f"      {str(c)[:36]:36s} nb{nb:5.1f}% uniq{nun:>6}  {samp[:50]}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_id_columns.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
