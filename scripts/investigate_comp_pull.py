r"""investigate_comp_pull —— 點解 Tableau 拉唔到 comp？

背景：user 喺 Tableau 拉「comp」拉出嚟嘅 pivot，數字其實係 STAFF(H_LABOR)。
呢個 script：
  1. 重出 STAFF(H_LABOR) pivot   → 證明你 paste 嗰個 = 人工，唔係 comp
  2. 重出 COMP(5 桶) pivot        → comp 真正應該出嘅數（對得返 recon COMP）
  3. comp 要 filter 咩            → horizontal_id / label / capex_opex / 空白 label
  4. staff vs comp filter 撞唔撞  → label 跨 id？comp row is_labor=1？
  5. amount_mop vs 調整前/後      → Tableau 應該 drag 邊條 measure

Run（Windows）:  python scripts\investigate_comp_pull.py
出 results\investigate_comp_pull.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "investigate_comp_pull.txt"
ENTS = ["galaxy", "wynn", "vml", "melco", "mgm", "sjm"]
COMP_H = ["H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"]
YBS = ["23", "24", "24_23SY", "25", "25_24SY", "25_23SY"]
L: list[str] = ["# investigate_comp_pull —— 點解 Tableau 拉唔到 comp（萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": "", "0.0": "0"})


def _pivot(sub, mcol, title, note=""):
    L.append("\n" + "=" * 78)
    L.append(f"## {title}")
    if note:
        L.append(f"   {note}")
    hdr = "   " + f"{'entity':<8}{'co':<7}" + "".join(f"{b:>11}" for b in YBS) + f"{'总和':>12}"
    L.append(hdr)
    tot = {b: 0.0 for b in YBS}
    for ent in ENTS:
        es = sub[sub["entity"].astype(str).eq(ent)]
        for cov in sorted(es["co"].unique()):
            r = es[es["co"].eq(cov)]
            if r.empty:
                continue
            cells, rtot = [], 0.0
            for b in YBS:
                v = r[r["year_bucket"].astype(str).eq(b)][mcol].sum()
                tot[b] += v
                rtot += v
                cells.append(f"{v:>11,.0f}" if abs(v) >= 0.5 else f"{'':>11}")
            L.append(f"   {ent:<8}{(cov or '(blank)'):<7}" + "".join(cells) + f"{rtot:>12,.0f}")
    L.append(f"   {'总和':<7}{'':<7}" + "".join(f"{tot[b]:>11,.0f}" for b in YBS) + f"{sum(tot.values()):>12,.0f}")


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    amt  = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    y2 = yb.str[:2]
    # recon basis：staff = 後 if y2∈{23,24} else 前；comp = 後 only if y2==24 else 前
    df["staff_m"] = post.where(y2.isin(["23", "24"]), pre)
    df["comp_m"]  = post.where(y2.eq("24"), pre)
    hid  = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    hlab = _s(df.get("horizontal_label", pd.Series("", index=df.index)))
    co   = _s(df.get("final_capex_opex", pd.Series("", index=df.index)))
    df["co"] = co
    df["hid"] = hid
    df["hlab"] = hlab

    staff = df[hid.eq("H_LABOR")]
    comp  = df[hid.isin(COMP_H)]

    # ── 1. 證明 paste 嗰個 = STAFF ──────────────────────────────────────────────
    _pivot(staff, "staff_m",
           "1. STAFF (H_LABOR) pivot —— 應同你 paste 嗰個一樣（證明你拉嘅係人工）",
           "basis: 23/24/24_23SY=調整後, 25/25_xx=調整前")

    # ── 2. COMP 真正應出嘅數 ────────────────────────────────────────────────────
    _pivot(comp, "comp_m",
           "2. COMP (5 桶) pivot —— 呢個先係 comp 應出嘅數（對得返 recon COMP）",
           "basis: 24/24_23SY=調整後, 23/25/25_xx=調整前 ; galaxy 25 應 ≈ 91,289")

    # ── 3. comp 要 filter 咩 ────────────────────────────────────────────────────
    L.append("\n" + "=" * 78)
    L.append("## 3. Tableau 要揀咩先拉到 comp（comp 係 5 個 label，要全部揀）")
    L.append("\n  ── comp horizontal_id × horizontal_label × 行數 ──")
    g = comp.groupby(["hid", "hlab"]).size().reset_index(name="n").sort_values("n", ascending=False)
    for _, r in g.iterrows():
        L.append(f"   {r['hid']:<16}{(r['hlab'] or '(空白)'):<16}{r['n']:>10,} 行")
    L.append("\n  ── comp final_capex_opex split（若 Tableau 淨 filter Opex 會漏 Capex comp）──")
    for cov, sub in comp.groupby("co"):
        L.append(f"   {(cov or '(空白)'):<10}{len(sub):>9,} 行   Σ comp_m={sub['comp_m'].sum():>12,.0f}萬")
    nb = int(comp["hlab"].eq("").sum())
    L.append(f"\n  ── comp horizontal_label 空白 = {nb:,} 行（空 label Tableau 揀唔到 → 會漏）")

    # ── 4. staff vs comp filter 撞唔撞 ─────────────────────────────────────────
    L.append("\n" + "=" * 78)
    L.append("## 4. staff vs comp filter 撞唔撞")
    labcomp = int((hid.eq("H_LABOR") & hlab.str.contains(r"Comp|comp|贈|房間|餐飲|會場|場地", na=False, regex=True)).sum())
    L.append(f"   H_LABOR 但 label 含 comp 字眼 = {labcomp:,} 行")
    if "is_labor" in df.columns:
        il = _s(df["is_labor"])
        L.append(f"   comp(5桶) 但 is_labor=1 = {int((hid.isin(COMP_H) & il.isin(['1','True','true','Y','y'])).sum()):,} 行")
    L.append("\n  ── 每個 comp label 喺邊啲 horizontal_id 出現（跨 id = label 唔乾淨，filter 會撈錯）──")
    for lab in [x for x in comp["hlab"].unique() if x]:
        ids = sorted(df[hlab.eq(lab)]["hid"].unique())
        flag = "  ⚠ 跨 id" if len(ids) > 1 else ""
        L.append(f"   {lab:<16}→ {', '.join(ids)}{flag}")

    # ── 5. amount_mop vs 調整前/後（Tableau drag 邊條 measure？）─────────────────
    L.append("\n" + "=" * 78)
    L.append("## 5. amount_mop vs 調整前_萬 / 調整後_萬（per year_bucket）")
    L.append(f"   全表：amount_mop Σ={amt.sum():,.0f}（raw）={amt.sum()/1e4:,.0f}萬 ; 前 Σ={pre.sum():,.0f}萬 ; 後 Σ={post.sum():,.0f}萬")
    L.append("\n  ── per year_bucket: amt_萬 ｜ 調整前 ｜ 調整後（睇 amount_mop 跟邊個 → 決定 Tableau measure）──")
    L.append("   " + f"{'year_bucket':<12}{'amt_萬':>14}{'調整前':>14}{'調整後':>14}")
    for ybk in YBS:
        m = yb.eq(ybk)
        L.append("   " + f"{ybk:<12}{amt[m].sum()/1e4:>14,.0f}{pre[m].sum():>14,.0f}{post[m].sum():>14,.0f}")

    _w()


def _w():
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
