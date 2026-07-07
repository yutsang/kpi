#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
adj_pivot.py — by-NG 調整 bridge 表

入 entity + year_bucket，行 = ng_label，欄：
  amount_mop(base) | <每個 調整一級 一欄> | 調整合計 | 調整後(=base+調整合計) | 調整後/前%

單位全部用「萬」(amount_mop ÷1e4 對齊 調整_萬)。

用法：
  python scripts/adj_pivot.py mgm 25
  python scripts/adj_pivot.py mgm 25,25_24SY,25_23SY        # 多 bucket 合計
  python scripts/adj_pivot.py mgm 25 --base pre             # base 用 調整前_萬（mgm/sjm 建議）
  python scripts/adj_pivot.py --selftest                    # synthetic 自測，唔使 CSV

備註 base 選擇：
  amount_mop（預設）：galaxy/vml/melco/wynn 25 = 調整前，所以 base+調整 = 調整後 成立。
  但 mgm / sjm 的 amount_mop = 調整後，base+調整 就唔等於 調整後。
  → mgm/sjm 想睇「前→後」reconciliation 用 --base pre。
  無論揀邊個，output 都會印三個 anchor(Σamount_mop / Σ調整前 / Σ調整後) 俾你核對。
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

# 大 CSV 候選路徑（Windows prep_tableau 輸出）
_CANDS = [
    Path("tableau_combined_25.csv"),
    Path("data/tableau_combined_25.csv"),
    Path("data/08_reporting/tableau_combined_25.csv"),
]
_COLS = ["entity", "year_bucket", "ng_code", "ng_label",
         "調整一級", "調整_萬", "調整前_萬", "調整後_萬", "amount_mop"]
_BLANK = {"", "0", "0.0", "nan", "None", "未分類調整"}  # 未分類調整 仍當一個 lv1 名（見下）


def _find_csv(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            sys.exit(f"✗ 揾唔到 CSV: {p}")
        return p
    for c in _CANDS:
        if c.exists():
            return c
    sys.exit("✗ 揾唔到 tableau_combined_25.csv（用 --csv 指定路徑）")


def _load(csv: Path, entity: str, buckets: list[str]) -> pd.DataFrame:
    frames = []
    for ch in pd.read_csv(csv, usecols=lambda c: c.strip() in _COLS,
                          chunksize=300_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        m = ch["entity"].astype(str).str.strip().eq(entity) & \
            ch["year_bucket"].astype(str).str.strip().isin(buckets)
        if m.any():
            frames.append(ch.loc[m])
    if not frames:
        return pd.DataFrame(columns=_COLS)
    return pd.concat(frames, ignore_index=True)


def _ng_sortkey(lbl: str) -> tuple:
    """NG0 博彩... → 排序用 (0, lbl)；非 NG 排最後。"""
    s = str(lbl).strip().upper()
    if s.startswith("NG"):
        num = "".join(ch for ch in s[2:].split()[0] if ch.isdigit()) if len(s) > 2 else ""
        return (0, int(num) if num.isdigit() else 999, s)
    return (1, 999, s)


def build_adj_pivot(df: pd.DataFrame, base: str = "amount_mop") -> tuple[pd.DataFrame, dict]:
    """df 已 filter 好 entity+bucket。回 (pivot 表, anchors dict)。全部單位 = 萬。"""
    if df.empty:
        return pd.DataFrame(), {}

    for c in _COLS:
        if c not in df.columns:
            df[c] = ""

    # 數值欄（萬）
    adj = pd.to_numeric(df["調整_萬"], errors="coerce").fillna(0.0)
    pre = pd.to_numeric(df["調整前_萬"], errors="coerce").fillna(0.0)
    post = pd.to_numeric(df["調整後_萬"], errors="coerce").fillna(0.0)
    amt_mop_wan = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0) / 1e4

    if base in ("pre", "調整前", "調整前_萬"):
        base_series, base_name = pre, "調整前"
    else:
        base_series, base_name = amt_mop_wan, "amount_mop"

    work = pd.DataFrame({
        "ng_label": df["ng_label"].astype(str).str.strip().replace({"": "(未分類)"}),
        "_base": base_series.values,
        "_adj": adj.values,
        "_lv1": df["調整一級"].astype(str).str.strip().replace({"": "（無調整一級）"}),
    })

    # base per NG（全部行）
    base_by_ng = work.groupby("ng_label")["_base"].sum()

    # 調整 拆 lv1（pivot），只保留有調整嘅（drop 全零欄，包括 無調整一級）
    piv = work.pivot_table(index="ng_label", columns="_lv1", values="_adj",
                           aggfunc="sum", fill_value=0.0)
    if not piv.empty:
        piv = piv.loc[:, (piv != 0).any(axis=0)]
        order = piv.abs().sum().sort_values(ascending=False).index
        piv = piv[order]

    # 調整合計（直接由 df 加總，保證正確）
    adj_total = work.groupby("ng_label")["_adj"].sum()

    out = pd.DataFrame(index=base_by_ng.index)
    out[base_name] = base_by_ng
    for c in piv.columns:
        out[c] = piv[c]
    out["調整合計"] = adj_total
    out["調整後"] = out[base_name] + out["調整合計"]
    denom = out[base_name].replace(0, pd.NA)
    out["調整後/前%"] = (out["調整後"] / denom * 100)

    out = out.fillna(0.0)
    # NG 排序
    out = out.loc[sorted(out.index, key=_ng_sortkey)]

    # 合計 row
    total = out.drop(columns=["調整後/前%"]).sum(numeric_only=True)
    total["調整後/前%"] = (total["調整後"] / total[base_name] * 100) if total[base_name] else 0.0
    out.loc["合計"] = total

    anchors = {
        "Σamount_mop(萬)": float(amt_mop_wan.sum()),
        "Σ調整前(萬)": float(pre.sum()),
        "Σ調整(萬)": float(adj.sum()),
        "Σ調整後(萬)": float(post.sum()),
        "base_used": base_name,
    }
    return out, anchors


def _print_table(out: pd.DataFrame) -> None:
    with pd.option_context("display.max_rows", None, "display.max_columns", None,
                           "display.width", 240, "display.float_format", lambda x: f"{x:,.1f}"):
        print(out)


def _run(entity: str, year_bucket: str, base: str, csv: str | None) -> None:
    buckets = [b.strip() for b in str(year_bucket).replace("，", ",").split(",") if b.strip()]
    csv_p = _find_csv(csv)
    print(f"讀 {csv_p} … entity={entity} buckets={buckets} base={base}")
    df = _load(csv_p, entity, buckets)
    if df.empty:
        sys.exit(f"✗ 冇 row: entity={entity} bucket={buckets}")
    out, anchors = build_adj_pivot(df, base=base)

    print(f"\n=== {entity} {'/'.join(buckets)} — by-NG 調整 bridge（單位：萬）===")
    _print_table(out)
    print("\nanchors（核對 base 揀啱冇）:")
    for k, v in anchors.items():
        print(f"  {k}: {v:,.1f}" if isinstance(v, float) else f"  {k}: {v}")
    if anchors.get("base_used") == "amount_mop":
        d_pre = abs(anchors["Σamount_mop(萬)"] - anchors["Σ調整前(萬)"])
        d_post = abs(anchors["Σamount_mop(萬)"] - anchors["Σ調整後(萬)"])
        if d_post < d_pre:
            print("  ⚠ 此 entity amount_mop ≈ 調整後（mgm/sjm 類）；base+調整 ≠ 調整後。"
                  "想睇『前→後』請加 --base pre")

    # 寫檔
    outdir = Path("results")
    outdir.mkdir(exist_ok=True)
    stem = f"adj_pivot_{entity}_{'-'.join(buckets)}"
    tsv = outdir / f"{stem}.tsv"
    out.to_csv(tsv, sep="\t", encoding="utf-8-sig")
    print(f"\n✓ 寫 {tsv}")
    try:
        xlsx = outdir / f"{stem}.xlsx"
        out.to_excel(xlsx)
        print(f"✓ 寫 {xlsx}")
    except Exception as e:  # openpyxl 冇裝就跳過
        print(f"(xlsx 跳過: {e})")


def _selftest() -> None:
    """synthetic：唔使真 CSV，證邏輯。"""
    df = pd.DataFrame({
        "entity": ["mgm"] * 6,
        "year_bucket": ["25"] * 6,
        "ng_code": ["NG0", "NG0", "NG5", "NG5", "NG8", "NG8"],
        "ng_label": ["NG0 博彩項目", "NG0 博彩項目", "NG5 文化藝術",
                     "NG5 文化藝術", "NG8 美食", "NG8 美食"],
        "調整一級": ["超支調整", "重複入賬", "超支調整", "", "折讓", "超支調整"],
        "調整_萬": ["-100", "-50", "-30", "0", "-20", "-5"],
        "調整前_萬": ["1000", "500", "300", "200", "150", "80"],
        "調整後_萬": ["900", "450", "270", "200", "130", "75"],
        "amount_mop": ["9000000", "4500000", "2700000", "2000000", "1300000", "750000"],
    })
    print("── selftest: base=amount_mop（mgm→amount_mop=調整後, base+調整≠後）──")
    out, anc = build_adj_pivot(df.copy(), base="amount_mop")
    _print_table(out); print(anc)
    print("\n── selftest: base=pre（正確前→後）──")
    out2, anc2 = build_adj_pivot(df.copy(), base="pre")
    _print_table(out2); print(anc2)
    # 驗：base=pre 時 調整後 應 == Σ調整後_萬 per NG
    exp = {"NG0 博彩項目": 900 + 450, "NG5 文化藝術": 270 + 200, "NG8 美食": 130 + 75}
    ok = all(abs(out2.loc[k, "調整後"] - v) < 1e-6 for k, v in exp.items())
    print(f"\n驗算 調整後==Σ調整後_萬: {'PASS ✓' if ok else 'FAIL ✗'}")


def main() -> None:
    ap = argparse.ArgumentParser(description="by-NG 調整 bridge 表")
    ap.add_argument("entity", nargs="?", help="galaxy/sjm/wynn/vml/melco/mgm")
    ap.add_argument("year_bucket", nargs="?", help="25 或 25,25_24SY,25_23SY")
    ap.add_argument("--base", default="amount_mop",
                    help="amount_mop(預設) 或 pre(=調整前_萬)")
    ap.add_argument("--csv", default=None, help="CSV 路徑（預設自動揾）")
    ap.add_argument("--selftest", action="store_true", help="synthetic 自測")
    a = ap.parse_args()
    if a.selftest:
        _selftest()
        return
    if not a.entity or not a.year_bucket:
        ap.error("要俾 entity 同 year_bucket，或 --selftest")
    _run(a.entity, a.year_bucket, a.base, a.csv)


if __name__ == "__main__":
    main()
