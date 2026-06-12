"""Inspect the NEW tied raw files for galaxy / wynn / vml / melco (placed in each entity's
data/<ent>/raw/, overwriting old names). One run answers:

 (1) 6-bucket totals per entity — 23 / 24 / 24_23SY 用調整後; 25 / 25_24SY / 25_23SY 用調整前
     (per-entity amount/year rules supplied by the user 2026-06-12):
       galaxy: unified cols (amount_mop / adjustment_amount / adjusted_amount / year)
       wynn  : amount='Entry Voucher Amount/ Expense Amount'; 25 year col 是否2025年
               (2025年度/2024年度/2023年度); 24 year col 是否2024年 (2024年度/2023年計畫2024年完成情況);
               調整後 = 取數 + 調整金額
       vml   : amount='MOP Amt'; 25 year col 'Yr related' (Yr 2025/Yr 2024); 24 file 調整金額+調整後金額,
               全部 → bucket 24 (冇 23_24SY); 23 = new file (generic inspect)
       melco : 25 keep old conf (already ties); 24 amount='本位幣金額'+調整金額, year col
               'Report Years-重分類2023年期後調整後'; 23 = new file (generic inspect)
 (2) project / project_code / subproject fill% + distinct  (齊全否)
 (3) 調整事項 mode — every col whose name contains 調整: non-blank count, Y-tick count, numeric Σ
     → 分得出 tick-cols / 單一col / 金額col 三種模式
 (4) 投資項目清單 folder (golden 項目編號) — dump structure + sample codes + naive match rate
     vs each entity's project codes.

Run (Windows):  python scripts/inspect_new4.py
Output: prints + results/inspect_new4.txt (small — paste back)
"""
from __future__ import annotations
import sys
import time
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTS = ["galaxy", "wynn", "vml", "melco"]
PROJ_CANDS = ["project_code", "承批公司項目序號", "項目編號", "project", "Project Name", "項目名稱",
              "Name of Investment Project", "Project name - Amended", "项目名称&Unique ID", "Project"]
SUB_CANDS = ["subproject", "Subproject", "Sub project", "SubProject_Name", "子項目",
             "Project/Sub-Project Name", "WBS element"]

# (entity, year) → rules. amount=取數col, adjust=調整col, adjusted=調整後col(優先),
# ycol=年份col(substring匹配), map=值→bucket, bucket=全檔單一bucket, basis=前/後
RULES = {
    ("galaxy", "2025"): dict(amount="amount_mop", ycol="year",
                             map={"2025": "25", "2024": "25_24SY", "2023": "25_23SY"}, basis="前"),
    ("galaxy", "2024"): dict(amount="amount_mop", adjust="adjustment_amount", adjusted="adjusted_amount",
                             ycol="year", map={"2024": "24", "2023": "24_23SY"}, basis="後"),
    ("galaxy", "2023"): dict(amount="amount_mop", adjust="adjustment_amount", adjusted="adjusted_amount",
                             ycol="year", map={"2023": "23"}, basis="後"),
    ("wynn", "2025"): dict(amount="Entry Voucher Amount/ Expense Amount", ycol="是否2025",
                           map={"2025年度": "25", "2024年度": "25_24SY", "2023年度": "25_23SY"}, basis="前"),
    ("wynn", "2024"): dict(amount="Entry Voucher Amount/ Expense Amount", adjust="調整金額", ycol="是否2024",
                           map={"2024年度": "24", "2023年計": "24_23SY"}, basis="後", fuzzy_year=True),
    ("wynn", "2023"): dict(amount="Entry Voucher Amount/ Expense Amount", adjust="調整金額",
                           bucket="23", basis="後"),
    ("vml", "2025"): dict(amount="MOP Amt", ycol="Yr related",
                          map={"Yr 2025": "25", "Yr 2024": "25_24SY", "Yr 2023": "25_23SY"}, basis="前"),
    ("vml", "2024"): dict(amount="MOP Amt", adjust="調整金額", adjusted="調整後金額",
                          bucket="24", basis="後"),
    ("vml", "2023"): dict(amount="MOP Amt", adjust="調整金額", adjusted="調整後金額",
                          bucket="23", basis="後"),
    ("melco", "2025"): dict(amount="Amount - Amended", ycol="years",
                            map={"2025": "25", "2024年期後": "25_24SY", "2024期後": "25_24SY"}, basis="前"),
    ("melco", "2024"): dict(amount="本位幣金額", adjust="調整金額", ycol="Report Years-重分類",
                            map={"2024": "24", "2023": "24_23SY"}, basis="後", fuzzy_year=True),
    ("melco", "2023"): dict(amount="本位幣金額", adjust="調整金額", bucket="23", basis="後"),
}


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce")


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def find_col(df, needle):
    if needle in df.columns:
        return needle
    cands = [c for c in df.columns if str(needle).lower() in str(c).lower()]
    return cands[0] if cands else None


def pick_sheet(path):
    names = pd.ExcelFile(path).sheet_names
    low = {str(s).strip().lower(): s for s in names}
    for want in ("data", "opex&capex匯總", "25je", "24je", "23je", "2023je",
                 "combine(clean)", "明細賬", "combine", "sheet1"):
        if want in low:
            return low[want], names
    return names[0], names


def inspect_file(L, ent, path, year, codes_out):
    sheet, names = pick_sheet(path)
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    mt = time.strftime("%m-%d %H:%M", time.localtime(path.stat().st_mtime))
    L.append(f"\n### {path.name} (sheet={sheet!r}, mtime={mt})  rows={len(df):,}  sheets={names}")
    # unified-schema file (new build) → unified rules regardless of entity
    if "amount_mop" in df.columns:
        r = RULES.get(("galaxy", year)) or {}
        L.append("   [unified schema detected → unified rules]")
    else:
        r = RULES.get((ent, year))
    if not r:
        L.append("   (no rules — generic)"); r = {}
    ac = find_col(df, r.get("amount", "amount_mop"))
    if not ac:
        L.append(f"   !! amount col {r.get('amount')!r} MISSING — cols: {[str(c)[:22] for c in df.columns][:25]}")
        return
    amt = _num(df[ac]).fillna(0.0)
    final = amt
    note = f"取數={ac!r}"
    if r.get("basis") == "後":
        adjc = find_col(df, r.get("adjusted") or "__none__")
        if adjc:
            final = _num(df[adjc]).fillna(amt)   # 調整後欄優先
            note += f"  調整後欄={adjc!r}"
        else:
            adc = find_col(df, r.get("adjust") or "調整金額")
            if adc:
                final = amt + _num(df[adc]).fillna(0.0)
                note += f"  調整後=取數+{adc!r}(Σ調={_num(df[adc]).fillna(0).sum()/1e6:,.2f}M)"
            else:
                note += "  !! 調整欄 MISSING(用取數)"
    L.append(f"   {note}  basis=調整{r.get('basis','?')}  Σ={final.sum()/1e6:,.2f}M")
    # buckets
    if r.get("bucket"):
        L.append(f"   bucket {r['bucket']}: Σ={final.sum()/1e6:,.2f}M  ({len(df):,}r)")
    elif r.get("ycol"):
        yc = find_col(df, r["ycol"])
        if not yc:
            L.append(f"   !! year col {r['ycol']!r} MISSING")
        else:
            yv = _s(df, yc)
            L.append(f"   year col={yc!r}:")
            mapped = pd.Series("?", index=df.index)
            for raw, b in (r.get("map") or {}).items():
                m = yv.str.contains(raw, regex=False) if r.get("fuzzy_year") else yv.eq(raw)
                mapped = mapped.mask(m, b)
            for b in sorted(set(mapped)):
                m = mapped.eq(b)
                L.append(f"      {b:8s} {int(m.sum()):>8,}r  Σ={final[m].sum()/1e6:>12,.2f}M")
            if (mapped == "?").any():
                bad = yv[mapped.eq('?')].replace("", "(blank)").value_counts().head(5)
                L.append(f"      ? 未map值: " + "  ".join(f"{k}×{n}" for k, n in bad.items()))
    # project / subproject completeness
    pcol = next((c for c in PROJ_CANDS if find_col(df, c)), None)
    parts = []
    for cands, lab in ((PROJ_CANDS, "proj"), (SUB_CANDS, "sub")):
        found = [find_col(df, c) for c in cands]
        found = [c for c in found if c]
        seen = set()
        found = [c for c in found if not (c in seen or seen.add(c))][:2]
        for c in found:
            s = _s(df, c)
            parts.append(f"{lab}:{str(c)[:18]} fill={s.ne('').mean()*100:.0f}% d={s.nunique()}")
    L.append("   " + ("  |  ".join(parts) if parts in ([],) or parts else "!! NO project/subproject cols"))
    # collect codes for 清單 matching
    ccol = next((find_col(df, c) for c in ("project_code", "承批公司項目序號", "項目編號") if find_col(df, c)), None)
    if ccol:
        codes_out.setdefault(ent, set()).update(
            _s(df, ccol).str.replace(r"\.0$", "", regex=True).replace("", pd.NA).dropna().unique()[:800])
    # 調整 cols census
    adj_cols = [c for c in df.columns if "調整" in str(c) or "调整" in str(c)]
    if adj_cols:
        L.append("   調整-cols:")
        for c in adj_cols[:10]:
            s = _s(df, c)
            nb = int(s.ne("").sum())
            ys = int(s.str.upper().eq("Y").sum())
            v = _num(df[c])
            L.append(f"      {str(c)[:42]:42s} nonblank={nb:>6,}  Y={ys:>5,}  Σnum={float(v.fillna(0).sum())/1e6:>10.2f}M")


def main():
    L = ["# inspect_new4 — galaxy/wynn/vml/melco new tied raw: 6-bucket totals + proj/sub + 調整模式 + 清單 match"]
    codes = {}
    for ent in ENTS:
        rawdir = ROOT / "data" / ent / "raw"
        L.append(f"\n{'='*76}\n## {ent} — {rawdir}")
        if not rawdir.exists():
            L.append("   X raw dir missing"); continue
        files = sorted(p for p in rawdir.glob("*.xls*")
                       if not p.name.startswith("~$") and "copy" not in p.name.lower())
        L.append("   files: " + ", ".join(f"{p.name}({p.stat().st_size//1024}KB)" for p in files))
        for p in files:
            yr = next((y for y in ("2025", "2024", "2023") if y in p.name), None)
            if not yr:
                continue
            try:
                inspect_file(L, ent, p, yr, codes)
            except Exception as e:
                L.append(f"   !! {p.name}: {type(e).__name__}: {e}")

    # ── 投資項目清單 (golden project list) ────────────────────────────────────
    L.append(f"\n{'='*76}\n## 投資項目清單 (golden 項目編號)")
    for ent, eset in codes.items():
        L.append(f"   entity {ent} code samples: {sorted(list(eset))[:8]}  (n={len(eset)})")
    ldir = ROOT / "data" / "投資項目清單"
    if not ldir.exists():
        L.append("   X data\\投資項目清單 missing");
    else:
        for p in sorted(ldir.glob("*.xls*")):
            if p.name.startswith("~$"):
                continue
            ent = next((e for e in ("galaxy", "mgm", "melco", "sjm", "vml", "wynn")
                        if e in p.name.lower()), None)
            try:
                xl = pd.ExcelFile(p)
                shn = next((s for s in xl.sheet_names if "database" in str(s).lower()
                            or "分項目數據簿" in str(s)), xl.sheet_names[0])
                # header = the row that contains 清單序號
                probe = pd.read_excel(p, sheet_name=shn, header=None, dtype=str, nrows=15)
                hdr = next((i for i in range(len(probe))
                            if probe.iloc[i].astype("string").fillna("").str.strip().eq("清單序號").any()), 0)
                t = pd.read_excel(p, sheet_name=shn, header=hdr, dtype=str)
                t.columns = [str(c).strip() for c in t.columns]
                L.append(f"\n   FILE {p.name[:40]} [{shn}] header_row={hdr} rows={len(t):,}")
                L.append(f"      cols: {[str(c)[:18] for c in t.columns][:18]}")
                cc = find_col(t, "清單序號") or find_col(t, "序號")
                nc = find_col(t, "項目名稱") or find_col(t, "名稱")
                if cc:
                    cv = _s(t, cc).str.replace(r"\.0$", "", regex=True)
                    pairs = [(c_, str(n_)[:22]) for c_, n_ in
                             zip(cv, _s(t, nc) if nc else [""] * len(t)) if c_][:8]
                    L.append(f"      清單序號 sample: {pairs}")
                    lset = {v for v in cv.unique() if v}
                    if ent and codes.get(ent):
                        inter = codes[ent] & lset
                        L.append(f"      ★ match {ent}: {len(inter)}/{len(codes[ent])} entity-codes in 清單"
                                 f" (清單 n={len(lset)})  e.g. {sorted(list(inter))[:6]}")
                # 金額欄 sums (golden 對數用)
                sums = []
                for c in t.columns:
                    v = _num(t[c].astype("string"))
                    if v.notna().sum() > 3 and abs(float(v.sum() or 0)) > 1e5:
                        sums.append((str(c)[:24], float(v.sum())))
                sums.sort(key=lambda x: -abs(x[1]))
                if sums:
                    L.append("      Σ欄(top5): " + "  ".join(f"{c}:Σ{s/1e6:,.1f}M" for c, s in sums[:5]))
            except Exception as e:
                L.append(f"   FILE {p.name}: read error {e}")
    out = ROOT / "results" / "inspect_new4.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
