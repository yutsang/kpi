r"""診斷 wynn subproject —— 揾返 wynn 原本有嘅 subproject 欄叫咩名、填幾多。

背景：wynn conf 冇 set subproject_name_col → prep_tableau 預設揾「Sub project」欄。
若 raw 嘅 header 唔係 exactly「Sub project」（多咗空格 / 大細階 / 改名 / 被 excel 拉拽整爛），
prep 就揾唔到 → subproject 欄被 4 層 fill 用 project 名頂上（睇落似冇 subproject）。

呢個 script 掃三層，列出所有含 sub/project/name 嘅欄 + 非空比例 + 樣本值：
  ① 最終 report parquet   data/wynn/output/company_3_kpi_report.parquet
  ② interim raw parquet    data/wynn/interim/company_3_raw.parquet
  ③ raw xlsx（data/wynn 下所有 .xlsx 逐 sheet）

Run（Windows）:  python scripts\diag_wynn_subproject.py
出 results\diag_wynn_subproject.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
L: list[str] = ["# diag_wynn_subproject —— wynn subproject 欄診斷"]


def _hit(c: str) -> bool:
    s = str(c).lower()
    return any(k in s for k in ("sub", "project", "name", "投資", "項目", "名稱", "细", "細"))


def _fill(series) -> tuple[int, int, list]:
    s = series.astype(str).str.strip()
    blank = s.isin(["", "nan", "None", "NaN", "<NA>"])
    nb = int((~blank).sum())
    samp = [x for x in s[~blank].unique().tolist()[:6]]
    return nb, len(series), samp


def _scan_df(tag: str, df: pd.DataFrame):
    L.append(f"\n{'='*64}\n## {tag}  ({len(df):,} rows)")
    cols = [c for c in df.columns if _hit(c)]
    if not cols:
        L.append("   (冇 sub/project/name 相關欄)")
    for c in cols:
        nb, tot, samp = _fill(df[c])
        L.append(f"   「{c}」 非空 {nb:,}/{tot:,}  例: " + " | ".join(str(x)[:30] for x in samp))


def main():
    p1 = ROOT / "data/wynn/output/company_3_kpi_report.parquet"
    if p1.exists():
        try: _scan_df(f"① report parquet  {p1.relative_to(ROOT)}", pd.read_parquet(p1))
        except Exception as e: L.append(f"\n① 讀 {p1.name} 失敗: {e}")
    else:
        L.append(f"\n① {p1.relative_to(ROOT)} 唔存在")

    p2 = ROOT / "data/wynn/interim/company_3_raw.parquet"
    if p2.exists():
        try: _scan_df(f"② interim raw parquet  {p2.relative_to(ROOT)}", pd.read_parquet(p2))
        except Exception as e: L.append(f"\n② 讀 {p2.name} 失敗: {e}")
    else:
        L.append(f"\n② {p2.relative_to(ROOT)} 唔存在（未跑過 / 已 del）")

    wdir = ROOT / "data/wynn"
    xlsxs = sorted([x for x in wdir.rglob("*.xlsx") if "~$" not in x.name]) if wdir.exists() else []
    L.append(f"\n{'='*64}\n## ③ raw xlsx（{len(xlsxs)} 個）")
    for x in xlsxs:
        try:
            xl = pd.ExcelFile(x)
        except Exception as e:
            L.append(f"\n  {x.relative_to(ROOT)} 開唔到: {e}"); continue
        for sh in xl.sheet_names:
            try:
                df = xl.parse(sh, nrows=2000, dtype=str)
            except Exception as e:
                L.append(f"\n  {x.name}[{sh}] 讀唔到: {e}"); continue
            _scan_df(f"{x.name} [sheet: {sh}]", df)

    out = ROOT / "results" / "diag_wynn_subproject.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
