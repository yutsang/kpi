r"""快（CSV-based，唔讀 raw）：comp 房 by NG —— 我哋 Comp房間 vs golden 客房，
gap NG 列 top account_desc/description（睇贈房 tag 咗去邊個 H）。user 方向：desc by NG。

Run（Windows）:  python scripts\inspect_ng_desc.py galaxy
                 python scripts\inspect_ng_desc.py galaxy NG3   # 淨 NG3 詳列
出 results\inspect_ng_desc_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
GOLD = ROOT / "internal_resource_golden.xlsx"
args = sys.argv[1:]
ENT = next((a for a in args if a.lower() in ("galaxy", "wynn", "vml", "melco", "mgm", "sjm")), "galaxy")
NGF = next((a for a in args if a.upper().startswith("NG")), None)


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str) == ENT].copy()
    e["m"] = pd.to_numeric(e.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    e["pre"] = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    for c in ("horizontal_label", "account_desc", "description", "ng_code", "項目組H", "year_bucket"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["y2"] = e["year_bucket"].str[:2]
    # golden NG 用中文名 → 對 ng_code
    NG2CODE = {"博彩娛樂": "NG0", "吸引外國客源": "NG1", "會議展覽": "NG2", "娛樂表演": "NG3",
               "體育盛事": "NG4", "文化藝術": "NG5", "健康養生": "NG6", "主題遊樂": "NG7",
               "美食之都": "NG8", "社區旅遊": "NG9", "海上旅遊": "NG10", "其他": "NG11"}
    # golden 房 by NG（23+24 internal_resource）
    g_room = {}
    if GOLD.exists():
        g = pd.read_excel(GOLD); g.columns = [str(c).strip() for c in g.columns]
        g = g[g["entity"].astype(str).str.strip() == ENT]
        g["客房支出"] = pd.to_numeric(g.get("客房支出", 0), errors="coerce").fillna(0.0)
        for _, r in g.iterrows():
            ng = NG2CODE.get(str(r.get("NG", "")).strip(), str(r.get("NG", "")).strip())
            g_room[ng] = g_room.get(ng, 0) + r["客房支出"]
    L = [f"# inspect_ng_desc — {ent if (ent:=ENT) else ''} comp房 by NG（amount_mop萬）"]
    L.append(f"\n## 房 comp（Comp房間）by NG：我哋(24prefix調後) vs golden客房(23+24)")
    room = e[e.horizontal_label.eq("Comp房間")]
    our_by_ng = room[room.y2.isin(["23", "24"])].groupby("ng_code")["m"].sum()
    ngs = sorted(set(list(our_by_ng.index) + list(g_room.keys())))
    gaps = []
    for ng in ngs:
        ours = our_by_ng.get(ng, 0)
        gold = g_room.get(ng, 0)
        flag = "" if abs(ours - gold) < 200 else " ⚠"
        L.append(f"   {ng:<5} 我={ours:>8,.0f}  golden={gold:>8,.0f}  Δ={ours-gold:>8,.0f}{flag}")
        if abs(ours - gold) >= 500 and gold > ours:
            gaps.append(ng)

    # gap NG（golden 房 > 我哋）→ 列嗰 NG 全部 row 嘅 top account_desc + description（搵贈房）
    focus = [NGF] if NGF else gaps
    for ng in focus:
        sub = e[e.ng_code.eq(ng) & e.y2.isin(["23", "24"])]
        L.append(f"\n{'='*60}\n## NG {ng}（golden房={g_room.get(ng,0):,.0f} 我房={our_by_ng.get(ng,0):,.0f}）—— 全 NG top account_desc / description（搵贈房收埋邊）")
        L.append("  ── top account_desc（by amount, 全 H）──")
        ga = sub.groupby(["account_desc", "horizontal_label"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in ga.head(15).iterrows():
            if abs(r["m"]) >= 50:
                L.append(f"     {r['account_desc'][:34]:<34} H={r['horizontal_label'][:8]:<8} {r['m']:>8,.0f}")
        L.append("  ── top description（搵 room/房/lodging/guest 字）──")
        gd = sub.groupby(["description", "horizontal_label"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in gd.head(15).iterrows():
            if abs(r["m"]) >= 50:
                hit = " ★房?" if pd.Series([r["description"]]).str.contains(r"room|房|lodg|guest|hotel|stay|住", case=False, regex=True, na=False).iloc[0] else ""
                L.append(f"     {r['description'][:34]:<34} H={r['horizontal_label'][:8]:<8} {r['m']:>8,.0f}{hit}")
    out = ROOT / "results" / f"inspect_ng_desc_{ENT}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
