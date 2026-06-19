r"""拆 vml 人工成本(opex) 嘅來源 —— 邊條 prep 規則 / 分類1 mark 谷大佢（HQ vml staff 細：790/1457/1253）。

每 bucket：opex 人工成本 Σ調整後，再拆：
  - account_desc 係 CONTRACT LABOR / SALARIES / SOCIAL SECURITY（prep vml H修正→人工）幾多
  - is_labor flag 真 幾多（prep is_labor→人工）
  - 項目組H(分類1) = 人工成本/管理成本分攤/部門間成本分攤（step4 column_map，項目組自己 mark）幾多
睇邊樣係 over 嘅元兇 → 決定 remove 邊條 prep override，俾新 mark 話事。

Run（Windows）:  python scripts\inspect_vml_labor.py
出 results\inspect_vml_labor.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
L: list[str] = ["# inspect_vml_labor —— vml opex 人工成本 來源拆解（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df = df[df["entity"].astype(str) == "vml"].copy()
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("horizontal_id", "final_capex_opex", "account_desc", "is_labor", "項目組H"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    opex = ~df["final_capex_opex"].eq("Capex")
    lab = df[(df.horizontal_id == "H_LABOR") & opex]
    acc = lab["account_desc"].astype(str).str.upper()
    il = lab["is_labor"].astype(str).str.strip().str.lower()
    il_t = ~il.isin(["", "0", "0.0", "n", "no", "false", "f", "nan", "none", "<na>"])
    pth = lab["項目組H"].astype(str).str.strip()

    is_contract = acc.str.contains("CONTRACT LABOR", na=False)
    is_salary = acc.str.contains("SALAR", na=False)
    is_social = acc.str.contains("SOCIAL SECURITY", na=False)
    is_ptlabor = pth.isin(["人工成本", "管理成本分攤", "部門間成本分攤"])

    L.append(f"\n  vml opex 人工成本 共 {lab['post'].sum():,.0f}萬（{len(lab):,}行）")
    L.append(f"\n  ── by bucket（調整後萬）──")
    for yb, v in lab.groupby("yb")["post"].sum().items():
        L.append(f"     {yb:<9} {v:>9,.0f}")

    L.append(f"\n  ── 拆來源（重疊可能）每 bucket 調整後萬 ──")
    L.append(f"     {'bucket':<9}{'CONTRACT':>10}{'SALARIES':>10}{'SOCIAL':>9}{'is_labor':>10}{'項目組H人工':>11}")
    for yb in ["23", "24", "24_23SY", "25", "25_24SY"]:
        m = lab["yb"] == yb
        if not m.any():
            continue
        L.append(f"     {yb:<9}"
                 f"{lab.loc[m & is_contract,'post'].sum():>10,.0f}"
                 f"{lab.loc[m & is_salary,'post'].sum():>10,.0f}"
                 f"{lab.loc[m & is_social,'post'].sum():>9,.0f}"
                 f"{lab.loc[m & il_t.values,'post'].sum():>10,.0f}"
                 f"{lab.loc[m & is_ptlabor.values,'post'].sum():>11,.0f}")

    L.append(f"\n  ── CONTRACT LABOR 行而家 項目組H(分類1) 係咩？（睇新 mark 有冇 cover）──")
    cl = lab[is_contract]
    for k, v in cl.groupby(pth[cl.index].replace("", "(空白)"))["post"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(8).items():
        L.append(f"     項目組H={str(k)[:16]:<16} {v:>9,.0f}萬")
    _w()


def _w():
    out = ROOT / "results" / "inspect_vml_labor.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
