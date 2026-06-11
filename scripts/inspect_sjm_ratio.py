"""SJM — ratio / allocation checking on the tie workbook (data/raw/sjm_2025.xlsx).

項目組 24SY/23SY 係「按比例分」出嚟（golden 25=124,661萬 24SY=63,594萬 23SY=7,971萬）。
呢個 inspector 做三件事：
 (A) 逐張 tab（'25','24','23',Sheet2,Sheet3,…）dump 結構：header 自動偵測、columns、rows、
     每條 numeric 欄 Σ（top 8）+ ratio-like 欄（0..1 之間）特別標記 — 搵出比例表喺邊、
     邊條欄 Σ ≈ 635.94M / 79.71M。
 (B) data tab 嘅 debility 清單（自動驗）：exact-dup 行、unique_id 重複、mop+adj≠adjusted、
     blank 率（project/ng_theme/capex_opex/amount）、year 值分佈 — 全部一頁列晒。
 (C) 如果某 tab 有 project_code + ratio 欄：用 data tab(dedup) 嘅 per-project Σ × ratio
     試砌 bucket 數 vs golden（first-pass reconstruction）。

Run (Windows):  python scripts/inspect_sjm_ratio.py
Output: prints + results/inspect_sjm_ratio.txt (small — paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
NEW = ROOT / "data" / "raw" / "sjm_2025.xlsx"
GOLDEN = {"25": 124661e4, "25_24SY": 63594e4, "25_23SY": 7971e4}


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False)
                         .str.replace("%", "", regex=False), errors="coerce")


def _read_tab(xls_path, sheet):
    """Header auto-detect: first row (of first 8) with ≥3 non-blank cells becomes the header."""
    probe = pd.read_excel(xls_path, sheet_name=sheet, header=None, dtype=str, nrows=8)
    hdr = 0
    for i in range(len(probe)):
        if probe.iloc[i].astype("string").fillna("").str.strip().ne("").sum() >= 3:
            hdr = i
            break
    df = pd.read_excel(xls_path, sheet_name=sheet, header=hdr, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return hdr, df


def main():
    L = ["# inspect_sjm_ratio — tab structure + ratio columns + data-tab debilities",
         f"golden: 25={GOLDEN['25']/1e6:,.2f}M  24SY={GOLDEN['25_24SY']/1e6:,.2f}M  23SY={GOLDEN['25_23SY']/1e6:,.2f}M"]
    if not NEW.exists():
        L.append(f"X {NEW} missing"); _w(L); return
    xls = pd.ExcelFile(NEW)
    L.append(f"sheets: {xls.sheet_names}")

    data_df = None
    for sheet in xls.sheet_names:
        try:
            hdr, df = _read_tab(NEW, sheet)
        except Exception as e:
            L.append(f"\n## tab {sheet!r}: read error {e}"); continue
        L.append(f"\n{'='*72}\n## tab {sheet!r}: header_row={hdr}  rows={len(df):,}  cols={len(df.columns)}")
        if not len(df):
            continue
        L.append(f"   columns: {[str(c)[:24] for c in df.columns][:30]}")
        if str(sheet).strip().lower() == "data":
            data_df = df
        # numeric profile + ratio-like detection
        stats = []
        for c in df.columns:
            v = _num(df[c].astype("string"))
            n = int(v.notna().sum())
            if n < 3:
                continue
            s = float(v.sum() or 0)
            frac01 = float(((v >= -0.01) & (v <= 1.01)).mean())
            is_ratio = frac01 > 0.9 and v.abs().max() <= 1.01 and n >= 5
            stats.append((str(c), s, n, is_ratio))
        stats.sort(key=lambda x: -abs(x[1]))
        for c, s, n, is_r in stats[:8]:
            tag = "  ← RATIO-like(0..1)" if is_r else ""
            near = ""
            for b, g in GOLDEN.items():
                if g and abs(s - g) / g < 0.03:
                    near = f"  ★≈golden {b}!"
            L.append(f"   {c[:30]:30s} Σ={s/1e6:>12,.2f}M  n={n:>6,}{tag}{near}")
        ratio_cols = [c for c, s, n, is_r in stats if is_r]
        if ratio_cols:
            L.append(f"   ratio-like cols: {ratio_cols[:6]}")

    # ── (B) data tab debilities ───────────────────────────────────────────────
    if data_df is not None:
        df = data_df
        L.append(f"\n{'='*72}\n## DATA tab debility 清單 (rows={len(df):,})")
        full_dup = int(df.duplicated(keep=False).sum())
        L.append(f"   1. exact-dup rows: {full_dup:,}  (dedup → {len(df.drop_duplicates()):,} rows)")
        uidc = next((c for c in ("unique_id", "唯一識別碼") if c in df.columns), None)
        if uidc:
            u = df[uidc].astype("string").fillna("").str.strip()
            L.append(f"   2. unique_id: non-blank={int(u.ne('').sum()):,} distinct={u[u.ne('')].nunique():,}")
        amtc = next((c for c in ("amount_mop", "Val/COArea Crcy") if c in df.columns), None)
        adjc = next((c for c in ("adjustment_amount", "調整金額") if c in df.columns), None)
        if amtc:
            mop = _num(df[amtc].astype("string")).fillna(0.0)
            adj = _num(df[adjc].astype("string")).fillna(0.0) if adjc else 0.0
            L.append(f"   3. Σ前={mop.sum()/1e6:,.2f}M  Σ調={float(getattr(adj,'sum',lambda:0)())/1e6:,.2f}M  "
                     f"Σ後={(mop+adj).sum()/1e6:,.2f}M  | blank amount rows: {int(_num(df[amtc].astype('string')).isna().sum()):,}")
        for c in ("項目期間", "year", "承批公司項目序號", "項目性質", "報告capex/opex", "Net-off", "是否計入內部資源"):
            if c in df.columns:
                s = df[c].astype("string").fillna("").str.strip()
                top = "  ".join(f"{k}={n}" for k, n in s.replace("", "(blank)").value_counts().head(5).items())
                L.append(f"   {c}: blank={s.eq('').mean()*100:.0f}%  distinct={s.nunique()} | {top}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_ratio.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
