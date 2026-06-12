"""投資項目清單 → 兩套 per-entity 輸出 (user 2026-06-12):
 (1) results/proj_master_<ent>.tsv   — 項目維度 golden 數字（清單 database tab,按 承批公司項目序號
     dedupe;identity + 全部「實際投資/期後調整/潛在調整後」金額欄）→ 日後逐項目對數
 (2) results/proj_identity_<ent>.tsv — raw JE 嘅 distinct (project_code, dicj_code, project,
     subproject) 組合 + rows/Σ —— 項目名/子項目/號碼/DICJ 號碼對照表
Run (Windows):  python scripts/dump_project_master.py
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
LDIR = ROOT / "data" / "投資項目清單"
IDENT = ["清單序號", "承批公司項目序號", "項目類型", "項目性質", "項目名稱", "for mapping"]
AMT_PREFIX = ("2023年實際", "2023年度實際", "2023年報告", "2024年實際投資金額", "2025年實際投資金額",
              "2024年實際期後", "2025年實際期後", "潛在調整後投資金額", "潜在调整后投资金额",
              "三年累計", "公司填報三年", "潛在調整後三年")
RAW_CANDS = {  # entity → (folder, project_code cands, dicj cands, project cands, subproject cands, amount cands)
    "galaxy": ("galaxy", ["project_code"], ["dicj_code"], ["project"], ["subproject"], ["amount_mop"]),
    "sjm": ("sjm", ["承批公司項目序號"], ["dicj_code"], ["Project Name"], ["Subproject"], ["Val/COArea Crcy"]),
    "wynn": ("wynn", ["project_code"], ["dicj_code"], ["project", "Name of Investment Project"],
             ["subproject", "Sub project"], ["amount_mop", "Entry Voucher Amount/ Expense Amount"]),
    "vml": ("vml", ["project_code", "項目編號（for 2025執行報告）"], ["dicj_code"],
            ["project", "項目名稱"], ["subproject", "SubProject_Name"], ["amount_mop", "MOP Amt"]),
    "melco": ("melco", ["project_code", "承批公司項目序號"], ["dicj_code"],
              ["project", "項目名稱"], ["subproject", "Project/Sub-Project Name"],
              ["amount_mop", "Amount - Amended"]),
    "mgm": ("mgm", ["project_code", "Project_code"], ["dicj_code"], ["project", "Project_name"],
            ["subproject"], ["amount_mop", "Debit minus Credit"]),
}


def _s(df, col):
    v = df.get(col)
    if v is None:
        return pd.Series("", index=df.index)
    if isinstance(v, pd.DataFrame):
        v = v.iloc[:, 0]
    return v.astype("string").fillna("").str.strip()


def _find(df, cands):
    for c in cands:
        if c in df.columns:
            return c
        hit = next((x for x in df.columns if str(c).lower() in str(x).lower()), None)
        if hit:
            return hit
    return None


def main():
    out = ROOT / "results"; out.mkdir(exist_ok=True)
    # ── (1) 清單 golden per project ───────────────────────────────────────────
    for p in sorted(LDIR.glob("*.xls*")) if LDIR.exists() else []:
        if p.name.startswith("~$"):
            continue
        ent = next((e for e in RAW_CANDS if e in p.name.lower()), None)
        if not ent:
            continue
        print(f"[清單] {ent}: {p.name[:40]}...", flush=True)
        try:
            xl = pd.ExcelFile(p)
            shn = next((s for s in xl.sheet_names if "database" in str(s).lower()
                        or "分項目數據簿" in str(s)), xl.sheet_names[0])
            probe = pd.read_excel(p, sheet_name=shn, header=None, dtype=str, nrows=15)
            hdr = next((i for i in range(len(probe))
                        if probe.iloc[i].astype("string").fillna("").str.strip()
                        .isin(["清單序號", "承批公司項目序號"]).any()), 0)
            t = pd.read_excel(p, sheet_name=shn, header=hdr, dtype=str)
            t.columns = [str(c).strip() for c in t.columns]
            t = t.loc[:, ~t.columns.duplicated()]
            keep = [c for c in IDENT if c in t.columns]
            keep += [c for c in t.columns if str(c).startswith(AMT_PREFIX) and c not in keep]
            sub = t[keep].copy()
            pcol = "承批公司項目序號" if "承批公司項目序號" in sub.columns else keep[0]
            sub[pcol] = _s(sub, pcol)
            sub = sub[sub[pcol].ne("")].drop_duplicates(subset=[pcol], keep="first")
            fp = out / f"proj_master_{ent}.tsv"
            sub.to_csv(fp, sep="\t", index=False)
            print(f"   → {fp.name}: {len(sub)} projects × {len(keep)} cols", flush=True)
        except Exception as e:
            print(f"   !! {p.name}: {e}", flush=True)

    # ── (2) raw identity (project_code × dicj × project × subproject) ────────
    for ent, (folder, pc_c, dc_c, pj_c, sp_c, am_c) in RAW_CANDS.items():
        rawdir = ROOT / "data" / folder / "raw"
        if not rawdir.exists():
            continue
        frames = []
        for f in sorted(rawdir.glob(f"{ent}_20*.xls*")):
            if f.name.startswith("~$") or "copy" in f.name.lower():
                continue
            print(f"[identity] {ent}: {f.name}...", flush=True)
            try:
                sheets = pd.ExcelFile(f).sheet_names
                low = {str(s).strip().lower(): s for s in sheets}
                sh = next((low[w] for w in ("data", "opex&capex匯總", "25je", "24je", "combine(clean)",
                                            "combine", "sheet1") if w in low), sheets[0])
                df = pd.read_excel(f, sheet_name=sh, dtype=str)
                df.columns = [str(c).strip() for c in df.columns]
                cols = {k: _find(df, cands) for k, cands in
                        (("project_code", pc_c), ("dicj_code", dc_c), ("project", pj_c), ("subproject", sp_c))}
                ac = _find(df, am_c)
                g = pd.DataFrame({k: _s(df, v) if v else "" for k, v in cols.items()})
                g["amount"] = pd.to_numeric(_s(df, ac).str.replace(",", "", regex=False),
                                            errors="coerce").fillna(0.0) if ac else 0.0
                g["file"] = f.name
                frames.append(g)
            except Exception as e:
                print(f"   !! {f.name}: {e}", flush=True)
        if frames:
            allg = pd.concat(frames, ignore_index=True)
            idy = (allg.groupby(["project_code", "dicj_code", "project", "subproject"], dropna=False)
                   .agg(rows=("amount", "size"), amount=("amount", "sum")).reset_index()
                   .sort_values("amount", key=lambda x: x.abs(), ascending=False))
            fp = out / f"proj_identity_{ent}.tsv"
            idy.to_csv(fp, sep="\t", index=False)
            print(f"   → {fp.name}: {len(idy)} distinct tuples", flush=True)
    print("\nDONE — results\\proj_master_*.tsv + proj_identity_*.tsv (paste back / 留 results)")


if __name__ == "__main__":
    main()
