r"""睇每家嘅「項目組H」(pt_class_H) + comp_type 有冇真金額可以做 comp/staff 規律 mark。

user 要 bottom-up 搵規律 mark comp/staff（接受小殘差）。最有機會 = 我哋已 carry 嘅
「項目組H」(項目組自己 H 分類)。呢個 dump 每家：項目組H 各值 × 而家 horizontal_label × 萬，
+ comp_type 各值 × 萬。睇邊個 mark 欄有 room/venue/ticket/staff 真金額 → 寫 map rule。

Run（Windows）:  python scripts\inspect_ptclass.py            # 全 6 家
                 python scripts\inspect_ptclass.py galaxy wynn
出 results\inspect_ptclass.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
L: list[str] = ["# inspect_ptclass —— 項目組H / comp_type 有冇真金額做 comp/staff mark（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    for c in ("entity", "horizontal_label", "項目組H", "項目組V", "comp_type", "科目層級", "科目明細"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])

    for ent in ENTS:
        e = df[df.entity == ent]
        if not len(e):
            continue
        L.append(f"\n{'='*78}\n# {ent}  (項目組H 非空 {int(e['項目組H'].ne('').sum()):,}/{len(e):,}；comp_type 非空 {int(e['comp_type'].ne('').sum()):,})")
        # 項目組H 各值 → 萬 + 而家 H
        for col in ("項目組H", "comp_type"):
            sub = e[e[col].ne("")]
            if not len(sub):
                L.append(f"\n  「{col}」全空"); continue
            L.append(f"\n  ── 「{col}」各值（萬 + 而家 horizontal_label）──")
            g = sub.groupby(col)["amt"].agg(["sum", "size"]).reset_index().sort_values("sum", key=lambda s: s.abs(), ascending=False)
            for _, r in g.head(22).iterrows():
                cur = sub[sub[col] == r[col]].groupby("horizontal_label")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
                curs = " | ".join(f"{k or '(空)'}={v:,.0f}" for k, v in cur.head(3).items())
                L.append(f"     {str(r[col])[:26]:<26} Σ={r['sum']:>9,.0f}萬 ({int(r['size'])}行)  現H: {curs}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_ptclass.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
