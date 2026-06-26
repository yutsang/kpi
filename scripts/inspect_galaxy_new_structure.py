r"""inspect_galaxy_new_structure.py — 驗 galaxy_2025_new.xlsx 結構（換 raw 前 check）
Run: python scripts\inspect_galaxy_new_structure.py
Out: results\inspect_galaxy_new_structure.txt
對齊 conf 期望(欄名/Period值/sheet) + 同舊 galaxy_2025.xlsx 比金額改動。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_galaxy_new_structure.txt"
RAW  = ROOT / "data" / "galaxy" / "raw"
NEW_CANDS = [RAW/"galaxy_2025_new.xlsx", ROOT/"galaxy_2025_new.xlsx",
             Path.home()/"Downloads"/"galaxy_2025_new.xlsx"]
OLD  = RAW / "galaxy_2025.xlsx"
SHEET = "Combine(clean)"

# conf/company_1 2025 期望嘅欄（columns_override + dicj + filter_col + adjust）
EXPECT_COLS = ["Period", "DICJ Code", "Reported Amount(MOP)", "Capex/Opex", "Project",
               "NG11 Category", "Account Code", "Account Description", "Description",
               "Vendor Name", "Submit No.", "調整金額", "調整後金額", "一級調整", "二級調整"]
# year_split filter_col=Period 接受嘅值
EXPECT_PERIOD = {"25年度報告","2025年度報告","24年度期後","24年度報告","2024年度報告",
                 "23年度期後","23年度報告","2023年度報告"}


def _num(s): return pd.to_numeric(s.astype(str).str.replace(",","",regex=False), errors="coerce")


def _blank(s): return s.fillna("").astype(str).str.strip().isin(["","nan","None","NaN","<NA>"])


def main():
    L = ["# galaxy_2025_new.xlsx 結構檢查（換 raw 前）", ""]
    new = next((c for c in NEW_CANDS if c.exists()), None)
    if new is None:
        L.append("!! 揾唔到 galaxy_2025_new.xlsx，試過：")
        L += [f"   {c}" for c in NEW_CANDS]
        L.append("把檔放去 data\\galaxy\\raw\\ 再跑")
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8"); print("\n".join(L)); return
    L.append(f"NEW 檔: {new}")

    xl = pd.ExcelFile(new)
    L.append(f"Sheets: {xl.sheet_names}")
    sheet = SHEET if SHEET in xl.sheet_names else xl.sheet_names[0]
    L.append(f"用 sheet: [{sheet}]" + ("" if sheet == SHEET else f"  ⚠ conf 期望 '{SHEET}' 唔喺度！"))

    df = pd.read_excel(new, sheet_name=sheet, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"rows={len(df):,}  cols={len(df.columns)}")

    # 1) 欄齊唔齊
    L += ["", "── 1) conf 期望欄 vs 實際 ──"]
    miss = [c for c in EXPECT_COLS if c not in df.columns]
    L.append(f"  缺欄: {miss if miss else '無 ✓'}")
    extra = [c for c in df.columns if c not in EXPECT_COLS]
    L.append(f"  多出/其他欄: {extra}")

    # 2) Period 值
    L += ["", "── 2) Period 值（filter_col；唔喺白名單會掉 bucket）──"]
    if "Period" in df.columns:
        vc = df["Period"].fillna("<空>").astype(str).str.strip().value_counts()
        for v, n in vc.items():
            flag = "" if v in EXPECT_PERIOD else "  ⚠ 唔喺白名單"
            L.append(f"   {v:<18} {n:>9,}{flag}")
    else:
        L.append("  !! 冇 Period 欄")

    # 3) DICJ Code 填充率
    L += ["", "── 3) DICJ Code 填充率 ──"]
    if "DICJ Code" in df.columns:
        nb = (~_blank(df["DICJ Code"])).sum()
        L.append(f"   非空 {nb:,}/{len(df):,} ({nb/len(df)*100:.1f}%)  樣本={df.loc[~_blank(df['DICJ Code']),'DICJ Code'].dropna().unique()[:10].tolist()}")
    else:
        L.append("  !! 冇 DICJ Code 欄")

    # 4) 金額：調整前/調整金額/調整後 + 一致性
    L += ["", "── 4) 金額（總額 + 調整後 = 調整前+調整金額 一致性）──"]
    if "Reported Amount(MOP)" in df.columns:
        pre = _num(df["Reported Amount(MOP)"])
        L.append(f"   Reported Amount(MOP)[調整前] Σ = {pre.sum()/1e6:,.2f}M")
    if "調整金額" in df.columns:
        adj = _num(df["調整金額"]); L.append(f"   調整金額 Σ = {adj.sum()/1e6:,.2f}M  (非零 {int((adj.fillna(0)!=0).sum()):,} 行)")
    if "調整後金額" in df.columns:
        post = _num(df["調整後金額"]); L.append(f"   調整後金額 Σ = {post.sum()/1e6:,.2f}M")
    if all(c in df.columns for c in ["Reported Amount(MOP)","調整金額","調整後金額"]):
        diff = (post.fillna(0) - (pre.fillna(0)+adj.fillna(0)))
        bad = int((diff.abs() > 1).sum())
        L.append(f"   一致性 調整後−(調整前+調整金額)：|差|>1 嘅行 = {bad:,}（理想=0）；總差={diff.sum():,.0f}")

    # 5) Period × 調整後 (per-bucket 報告數)
    if "Period" in df.columns and "調整後金額" in df.columns:
        L += ["", "── 5) Period × 調整後金額 Σ(萬) ──"]
        for p, g in df.groupby(df["Period"].fillna("<空>").astype(str).str.strip()):
            L.append(f"   {p:<18} {_num(g['調整後金額']).sum()/1e4:>12,.0f}萬  ({len(g):,}行)")

    # 6) 同舊檔比
    L += ["", "── 6) vs 舊 galaxy_2025.xlsx ──"]
    if OLD.exists():
        odf = pd.read_excel(OLD, sheet_name=SHEET if SHEET in pd.ExcelFile(OLD).sheet_names else 0, dtype=str)
        odf.columns = [str(c).strip() for c in odf.columns]
        L.append(f"   舊 rows={len(odf):,}  新 rows={len(df):,}  Δ={len(df)-len(odf):+,}")
        L.append(f"   舊欄缺新檔: {[c for c in odf.columns if c not in df.columns]}")
        L.append(f"   新欄缺舊檔: {[c for c in df.columns if c not in odf.columns]}")
        for col in ["Reported Amount(MOP)", "調整後金額"]:
            if col in df.columns and col in odf.columns:
                L.append(f"   {col}: 舊 Σ={_num(odf[col]).sum()/1e6:,.2f}M  新 Σ={_num(df[col]).sum()/1e6:,.2f}M  Δ={(_num(df[col]).sum()-_num(odf[col]).sum())/1e6:+,.2f}M")
    else:
        L.append(f"   (舊檔唔喺度 {OLD})")

    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
