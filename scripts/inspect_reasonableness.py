r"""H/V classification reasonableness audit。
 H：逐個 horizontal_label 列 top account_desc + NG mix + entity mix（睇 H tag 啱唔啱）。
 V：逐個 vertical_label 列 top project/subproject 名 + NG + top account_desc（睇 V tag 啱唔啱）。
金額 = amount_mop（合併對數：23/24調整後 + 25調整前）。

Run（Windows）:
  python scripts\inspect_reasonableness.py          # H + V（出兩個檔）
  python scripts\inspect_reasonableness.py H        # 淨 H
  python scripts\inspect_reasonableness.py V galaxy # 淨 V，淨 galaxy
出 results\inspect_reason_H.txt / inspect_reason_V.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
args = [a for a in sys.argv[1:]]
WHICH = [a for a in args if a.upper() in ("H", "V")]
ENTF = [a for a in args if a.lower() in ("galaxy", "wynn", "vml", "melco", "mgm", "sjm")]
DO_H = (not WHICH) or ("H" in [w.upper() for w in WHICH])
DO_V = (not WHICH) or ("V" in [w.upper() for w in WHICH])


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def load():
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["m"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4  # 萬
    for c in ("horizontal_label", "vertical_label", "account_desc", "subproject", "project",
              "ng_code", "entity", "final_capex_opex"):
        df[c] = _s(df[c]) if c in df.columns else ""
    if ENTF:
        df = df[df["entity"].isin(ENTF)]
    return df


def ng_mix(sub, n=4):
    g = sub.groupby("ng_code")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    tot = sub["m"].sum() or 1
    return " ".join(f"{r['ng_code'] or '∅'}={r['m']/tot*100:.0f}%" for _, r in g.head(n).iterrows())


def ent_mix(sub, n=3):
    g = sub.groupby("entity")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    return " ".join(f"{r['entity']}={r['m']:,.0f}" for _, r in g.head(n).iterrows())


def run_H(df):
    L = ["# inspect_reason_H — 逐 H label 嘅 account_desc / NG / entity（萬，amount_mop）" + (f" [{','.join(ENTF)}]" if ENTF else "")]
    grand = df["m"].sum() or 1
    order = df.groupby("horizontal_label")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, lr in order.iterrows():
        lab = lr["horizontal_label"]
        sub = df[df.horizontal_label == lab]
        L.append(f"\n{'='*70}\n## [{lab or '∅'}]  Σ={sub['m'].sum():,.0f}萬 ({sub['m'].sum()/grand*100:.1f}%)  {len(sub):,}行")
        L.append(f"   NG: {ng_mix(sub)}")
        L.append(f"   entity: {ent_mix(sub)}")
        L.append(f"   ── top account_desc ──")
        g = sub.groupby("account_desc")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(12).iterrows():
            if abs(r["m"]) >= 50:
                L.append(f"      {r['account_desc'][:42]:<42} {r['m']:>10,.0f}")
    _w("H", L)


def run_V(df):
    L = ["# inspect_reason_V — 逐 V label 嘅 project / NG / account_desc（萬，amount_mop）" + (f" [{','.join(ENTF)}]" if ENTF else "")]
    order = df.groupby("vertical_label")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, lr in order.iterrows():
        lab = lr["vertical_label"]
        sub = df[df.vertical_label == lab]
        L.append(f"\n{'='*70}\n## [{lab or '∅'}]  Σ={sub['m'].sum():,.0f}萬  {len(sub):,}行")
        L.append(f"   NG: {ng_mix(sub, 5)}")
        L.append(f"   entity: {ent_mix(sub)}")
        L.append(f"   ── top project/subproject（名 + NG）──")
        g = sub.groupby(["subproject", "ng_code"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(12).iterrows():
            if abs(r["m"]) >= 50:
                L.append(f"      {r['ng_code'][:4]:<4} {r['subproject'][:46]:<46} {r['m']:>9,.0f}")
        L.append(f"   ── top account_desc ──")
        ga = sub.groupby("account_desc")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in ga.head(6).iterrows():
            if abs(r["m"]) >= 50:
                L.append(f"      {r['account_desc'][:42]:<42} {r['m']:>9,.0f}")
    _w("V", L)


def _w(tag, L):
    out = ROOT / "results" / f"inspect_reason_{tag}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = load()
    if DO_H:
        run_H(df)
    if DO_V:
        run_V(df)


if __name__ == "__main__":
    main()
