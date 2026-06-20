r"""揾 galaxy 房/票/comp 喺邊個 raw column 可以 filter 出 golden 值（user：肉眼睇到/filter得出）。
逐個分類欄（項目組H / comp_type / 科目層級 / 科目明細 / account_desc / 項目組V）show value 分佈，
特別 highlight 含 room/ticket keyword 嘅值 + 佢哋而家落咗咩 H。

Run（Windows）:  python scripts\inspect_galaxy_comp_col.py
出 results\inspect_galaxy_comp_col.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
COLS = ["項目組H", "comp_type", "科目層級", "科目明細", "account_desc", "項目組V", "remark"]
ROOM = r"房|客房|Room|Lodging|Hotel|Suite|住宿|Accommod"
TICK = r"票|門票|贈票|Ticket|Admission|入場|Club"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["m"] = post
    df["yb"] = df["year_bucket"].astype(str)
    df["ent"] = df["entity"].astype(str)
    for c in COLS + ["horizontal_label"]:
        df[c] = _s(df[c]) if c in df.columns else ""
    g = df[df.ent == "galaxy"]
    L = ["# inspect_galaxy_comp_col — galaxy 房/票 comp 喺邊個 column filter 到（golden 客房23=30,188/24=32,356；票23=3,990/24=4,024）"]
    for yv in ("23", "24"):
        gy = g[g.yb == yv]
        L.append(f"\n{'='*66}\n## galaxy {yv}")
        for col in COLS:
            if col not in gy.columns or not gy[col].astype(bool).any():
                continue
            # 含 room keyword 嘅 value
            mroom = gy[col].str.contains(ROOM, regex=True, na=False)
            if mroom.any() and gy.loc[mroom, "m"].abs().sum() >= 200:
                L.append(f"\n  ── [{col}] 含房 keyword Σ={gy.loc[mroom,'m'].sum() if False else gy.loc[mroom,'m'].sum():,.0f}萬 ──".replace("mroom", "mroom"))
                gg = gy[mroom].groupby([col, "horizontal_label"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
                for _, r in gg.head(8).iterrows():
                    L.append(f"       {r[col][:24]:<24} 現H={r['horizontal_label'][:9]:<9} {r['m']:>9,.0f}")
        # 票 keyword（單獨掃）
        for col in COLS:
            if col not in gy.columns:
                continue
            mt = gy[col].str.contains(TICK, regex=True, na=False)
            if mt.any() and gy.loc[mt, "m"].abs().sum() >= 100:
                L.append(f"\n  ── [{col}] 含票 keyword Σ={gy.loc[mt,'m'].sum():,.0f}萬 ──")
                gg = gy[mt].groupby([col, "horizontal_label"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
                for _, r in gg.head(6).iterrows():
                    L.append(f"       {r[col][:24]:<24} 現H={r['horizontal_label'][:9]:<9} {r['m']:>9,.0f}")
        # 項目組H 全 value 分佈（睇佢有冇 comp 分類）
        if "項目組H" in gy.columns and gy["項目組H"].astype(bool).any():
            L.append(f"\n  ── [項目組H] 全 value 分佈 top（睇有冇 客房/會場/食飲/贈票 分類）──")
            gh = gy.groupby("項目組H")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
            for _, r in gh.head(18).iterrows():
                if r["項目組H"] and abs(r["m"]) >= 100:
                    L.append(f"       {r['項目組H'][:30]:<30} {r['m']:>9,.0f}")
    out = ROOT / "results" / "inspect_galaxy_comp_col.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
