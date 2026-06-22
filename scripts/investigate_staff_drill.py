r"""investigate_staff_drill —— 跟進 investigate_staff，攞落 fix 所需嘅 year-landing。

A. galaxy 23 漏咗嘅 staff candidate（per-year allowlist 冇 23）
   → 睇 8000/8020/8221 糧系 account 喺 23 而家擺咗去邊（量度 -3,214 嘅缺口）
   → 若 23 account 全空白 = 23-build 問題（唔係淨係 allowlist 缺）
B. galaxy 爭議 account（8107500 / 7926000 / allowlist code）by year × 係咪 H_LABOR
   → 量化 8107500 推高 galaxy 24 幾多
C. vml staff candidate by year（80180 CONTRACT LABOR 等）→ 24 over / 25 under 該點郁

Run（Windows）:  python scripts\investigate_staff_drill.py
出 results\investigate_staff_drill.txt  ← paste 返嚟
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "investigate_staff_drill.txt"
KW = (r"salary|wage|payroll|overtime|provident|gratuity|severance|pension|staff\s*cost|"
      r"casual\s*worker|contract\s*labou?r|employee\s*(?:benefit|relation|transport)|"
      r"recruitment\s*levy|bonus|薪|工資|花紅|公積金|員工|人工")
L: list[str] = ["# investigate_staff_drill —— galaxy23 / galaxy24-8107500 / vml（萬，recon basis）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": "", "0.0": "0"})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    y2 = df["year_bucket"].astype(str).str[:2]
    df["m"]  = post.where(y2.isin(["23", "24"]), pre)
    df["ry"] = y2
    df["e"]  = df["entity"].astype(str)
    df["hid"]   = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    df["co"]    = _s(df.get("final_capex_opex", pd.Series("", index=df.index)))
    df["desc"]  = _s(df.get("account_desc", pd.Series("", index=df.index)))
    df["acode"] = _s(df.get("account_code", pd.Series("", index=df.index))).str.replace(r"\.0$", "", regex=True)
    df["pth"]   = _s(df.get("項目組H", pd.Series("", index=df.index)))
    df["opex"]  = ~df["co"].eq("Capex")
    df["islab"] = df["hid"].eq("H_LABOR")
    df["kw"]    = df["desc"].str.contains(KW, case=False, regex=True, na=False)
    df["galpay"] = df["acode"].str.match(r"^(?:8000|8020|8221)")

    def yr3(sub):
        return [sub[sub.ry.eq(y)]["m"].sum() for y in ("23", "24", "25")]

    # ── A. galaxy 23 漏咗嘅 staff ───────────────────────────────────────────────
    L.append("=" * 80)
    L.append("## A. galaxy 23 — 糧系 account(8000/8020/8221) 或 keyword，opex，而家唔係 H_LABOR")
    L.append("   （galaxy 23 缺口 = HQ17,082 − 我13,868 = -3,214 → 呢度應補返幾多？）")
    A = df[df.e.eq("galaxy") & df.opex & df.ry.eq("23") & ~df.islab & (df.galpay | df.kw)]
    if A.empty:
        L.append("   ⚠ 0 行 → galaxy 23 account 可能全空白 = 23-build 問題（唔係 allowlist 缺）")
    else:
        L.append(f"   命中 {len(A):,} 行，Σ={A['m'].sum():,.0f}萬（vs 缺口 3,214）。現擺喺：")
        for h, v in A.groupby("hid")["m"].sum().sort_values(key=lambda s: s.abs(), ascending=False).items():
            L.append(f"     {h or '(空)':<16}{v:>11,.0f}萬")
        L.append("\n   ── by account（top 25）──")
        g = A.groupby(["acode", "desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False).head(25)
        for _, r in g.iterrows():
            L.append(f"     {r['acode']:<11}{(r['desc'] or '(空)')[:34]:<34}{r['m']:>11,.0f}萬")
    # 對照：galaxy 23 H_LABOR 而家由咩 account 砌成
    cur = df[df.e.eq("galaxy") & df.opex & df.ry.eq("23") & df.islab]
    L.append(f"\n   對照 galaxy 23 現有 H_LABOR Σ={cur['m'].sum():,.0f}萬，top account：")
    for _, r in cur.groupby(["acode", "desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False).head(10).iterrows():
        L.append(f"     {r['acode']:<11}{(r['desc'] or '(空)')[:34]:<34}{r['m']:>11,.0f}萬")

    # ── B. galaxy 爭議 account by year ─────────────────────────────────────────
    L.append("\n" + "=" * 80)
    L.append("## B. galaxy 爭議/allowlist account — Σ by year ｜ 當中幾多已係 H_LABOR")
    L.append("   （8107500 Event Service Fee 喺 24 allowlist 但 項目組H=others → 推高 24？）")
    L.append("   " + f"{'account':<10}{'desc':<26}" + "".join(f"{y:>9}{'(lab)':>9}" for y in ("23", "24", "25")))
    codes = ["8107500", "7926000", "8000000", "8000960", "8000150", "8000140"]
    gx = df[df.e.eq("galaxy") & df.opex & df.acode.isin(codes)]
    for ac in codes:
        sub = gx[gx.acode.eq(ac)]
        if sub.empty:
            continue
        dsc = sub["desc"].mode().iloc[0] if not sub["desc"].mode().empty else ""
        cells = ""
        for y in ("23", "24", "25"):
            sy = sub[sub.ry.eq(y)]
            cells += f"{sy['m'].sum():>9,.0f}{sy[sy.islab]['m'].sum():>9,.0f}"
        L.append("   " + f"{ac:<10}{dsc[:26]:<26}{cells}")

    # ── C. vml staff candidate by year ────────────────────────────────────────
    L.append("\n" + "=" * 80)
    L.append("## C. vml — staff candidate（is_labor 或 keyword 或 項目組=人工）by year ｜ (lab)")
    L.append("   （vml 24 +435 over / 25 -315 under；80180 CONTRACT LABOR 而家 H_PROFESSIONAL）")
    vp = df["pth"].str.contains(r"人工|staff|薪", case=False, regex=True, na=False)
    V = df[df.e.eq("vml") & df.opex & (df.islab | df.kw | vp)]
    L.append("   " + f"{'account':<11}{'desc':<24}{'cur_hid':<16}" + "".join(f"{y:>9}{'(lab)':>8}" for y in ("23", "24", "25")))
    g = V.groupby(["acode", "desc", "hid"])["m"].sum().reset_index()
    order = g.groupby("acode")["m"].transform(lambda s: s.abs().sum())
    g = g.assign(_o=order).sort_values("_o", ascending=False)
    seen = 0
    for _, r in g.iterrows():
        if seen >= 18:
            break
        sub = V[V.acode.eq(r["acode"]) & V.desc.eq(r["desc"]) & V.hid.eq(r["hid"])]
        cells = ""
        for y in ("23", "24", "25"):
            sy = sub[sub.ry.eq(y)]
            cells += f"{sy['m'].sum():>9,.0f}{sy[sy.islab]['m'].sum():>8,.0f}"
        L.append("   " + f"{r['acode']:<11}{(r['desc'] or '(空)')[:24]:<24}{(r['hid'] or '(空)')[:16]:<16}{cells}")
        seen += 1
    _w()


def _w():
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
