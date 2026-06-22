r"""inspect_staff_acctdesc —— 深拆 account_desc + 調整 + NG拆法 simulation。

承接 inspect_staff_dimensions，再落 account_desc 維度徹底睇：
  A. H_LABOR 全 account_desc（opex/capex 分開）—— payroll 完整組成
  B. CAPEX 非-H_LABOR 但「資本化人工」候選（項目組H=staff 或 desc 似資本化勞務：
     AUC/Asset Under Construction/Design/Development/PM/Management/Consult/Labour/Allocation）
     → galaxy/vml/sjm 漏咗嘅 Tab1 資本化人工，睇撈唔撈得返
  C. CAPEX account_desc top inventory（全 capex，唔限 keyword）—— 大舊 capex 係咩、有幾多似人工
  D. 調整(pre→post) by 調整二級 on H_LABOR —— 解釋 galaxy pre31,977→post13,955 剷咗咩
  E. MGM opex→capex 拆 simulation —— 按 NG capex% scale 到啱啱 tie HQ-opex 6,829，列出搬幾多

basis：pre=調整前_萬、post=調整後_萬；24-prefix（y2='24'）。
Run（Windows）:  python scripts\inspect_staff_acctdesc.py [ent ...]  (default galaxy vml sjm mgm)
出 results\inspect_staff_acctdesc.txt  ← paste 返嚟
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
OUT = ROOT / "results" / "inspect_staff_acctdesc.txt"
ENTS = sys.argv[1:] or ["galaxy", "vml", "sjm", "mgm"]

# 「資本化人工」候選 desc（比純 payroll keyword 闊：capex 入面嘅勞務/設計/PM/在建工程）
CAPLAB = (r"AUC|Asset Under Constr|Design\s*&|Design and|Development|Project\s*Manage|"
          r"\bPM\b|Management Fee|Consult|Profession|Labou?r|Allocation|Headcount|"
          r"salary|wage|payroll|staff|設計|管理費|勞務|人工|在建")
HQ_OPEX_24 = {"galaxy": 14574, "wynn": 8765, "vml": 1457, "melco": 13912, "mgm": 6829, "sjm": 0}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _top(L, sub, by, val, n=12, pre="       "):
    g = sub.groupby(by)[val].sum().reset_index().sort_values(val, key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(n).iterrows():
        label = "  ".join(str(r[c])[:34] for c in (by if isinstance(by, list) else [by]))
        L.append(f"{pre}{label[:52]:<52}{r[val]:>11,.0f}")
    return g


def main():
    L = ["# inspect_staff_acctdesc —— 深拆 account_desc/調整/NG拆法（萬，24-prefix）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"]  = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    df["e"] = df["entity"].astype(str)
    for c, n in [("horizontal_id", "hid"), ("final_capex_opex", "co"), ("account_code", "code"),
                 ("account_desc", "ad"), ("項目組H", "pth"), ("ng_code", "ng"), ("ng_label", "nglab"),
                 ("調整二級", "adj2"), ("horizontal_label", "hlab")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["code"] = df["code"].str.replace(r"\.0$", "", regex=True)
    df["cap"] = df["co"].eq("Capex")
    df["islab"] = df["hid"].eq("H_LABOR")
    df["pth_staff"] = df["pth"].str.contains(r"staff\s*cost|人工|薪|payroll", case=False, regex=True, na=False)
    df["caplab"] = df["ad"].str.contains(CAPLAB, case=False, regex=True, na=False)

    for ent in ENTS:
        e = df[df.e.eq(ent) & df.y2.eq("24")]
        if not len(e):
            continue
        L.append("\n" + "█" * 84)
        L.append(f"█ {ent}  24-prefix   HQ-opex={HQ_OPEX_24.get(ent,0):,}")
        L.append("█" * 84)
        lab = e[e.islab]

        # ── A. H_LABOR 全 account_desc（opex / capex）──
        L.append(f"\n  A. H_LABOR account_desc（opex post={lab[~lab.cap]['post'].sum():,.0f} / pre={lab[~lab.cap]['pre'].sum():,.0f}）")
        L.append("     ── opex（pre）top ──")
        _top(L, lab[~lab.cap], ["code", "ad"], "pre")
        if lab[lab.cap]["pre"].sum():
            L.append(f"     ── capex（pre）top ──")
            _top(L, lab[lab.cap], ["code", "ad"], "pre")

        # ── B. CAPEX 非-H_LABOR 資本化人工候選 ──
        capc = e[e.cap & ~e.islab & (e.pth_staff | e.caplab)]
        L.append(f"\n  B. CAPEX 非人工 但資本化人工候選（項目組H=staff 或 desc似勞務）：{len(capc):,}行 Σ(pre)={capc['pre'].sum():,.0f}")
        L.append("     （pth=項目組H標staff嗰啲最硬；desc-only要人眼睇係咪真勞務）")
        g = capc.groupby(["pth_staff", "code", "ad"])["pre"].sum().reset_index().sort_values("pre", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(14).iterrows():
            tag = "PTH=staff" if r["pth_staff"] else "desc-only"
            L.append(f"       [{tag}] {(r['code'] or '(空)'):<10}{(r['ad'] or '(空)')[:32]:<32}{r['pre']:>10,.0f}")

        # ── C. CAPEX account_desc top inventory（全capex）──
        L.append(f"\n  C. CAPEX 全 account_desc top（睇大舊capex係咩，capex Σ(pre)={e[e.cap]['pre'].sum():,.0f}）")
        _top(L, e[e.cap], ["code", "ad"], "pre", n=14)

        # ── D. 調整 pre→post by 調整二級 on H_LABOR ──
        lab2 = lab.copy(); lab2["adj"] = lab2["pre"] - lab2["post"]
        adj_tot = lab2["adj"].sum()
        if abs(adj_tot) > 1:
            L.append(f"\n  D. H_LABOR 調整(pre−post)={adj_tot:,.0f}（pre {lab['pre'].sum():,.0f}→post {lab['post'].sum():,.0f}）by 調整二級")
            _top(L, lab2[lab2["adj"].abs() > 0], ["adj2"], "adj", n=8)
        else:
            L.append(f"\n  D. H_LABOR 冇調整（pre=post={lab['pre'].sum():,.0f}）")

    # ── E. MGM opex→capex 拆 simulation ──
    if "mgm" in ENTS:
        m = df[df.e.eq("mgm") & df.y2.eq("24")]
        lab = m[m.islab & ~m.cap]
        cur_ox = lab["post"].sum()
        need_move = cur_ox - HQ_OPEX_24["mgm"]   # 要搬走幾多 opex 先 tie
        L.append("\n" + "═" * 84)
        L.append(f"## E. MGM opex→capex 拆 simulation（opex 而家 {cur_ox:,.0f}，HQ {HQ_OPEX_24['mgm']:,}，要搬 {need_move:,.0f} 去 capex）")
        # 按 NG capex share 排，由高 cap% NG 開始搬其 payroll-opex，累積到 need_move
        rows = []
        for ng, g in m.groupby("ng"):
            capx = g[g.cap]["pre"].sum(); opx = g[~g.cap]["pre"].sum()
            payox = g[g.islab & ~g.cap]["post"].sum()
            share = capx / (capx + opx) if (capx + opx) else 0
            rows.append((ng, share, payox, g["nglab"].iloc[0]))
        rows.sort(key=lambda r: -r[1])
        cum = 0.0
        L.append(f"     {'NG':<6}{'cap%':>6}{'payroll_ox':>11}{'累積搬':>9}  nglabel")
        for ng, share, payox, lab_ in rows:
            take = payox if cum < need_move else 0
            if take and cum + take > need_move:
                take = need_move - cum
            cum += take
            mark = f"搬{take:,.0f}" if take else "—"
            L.append(f"     {ng:<6}{share*100:>5,.0f}%{payox:>11,.0f}{cum:>9,.0f}  {lab_[:18]}  {mark}")
        L.append(f"     → 由高cap%NG搬payroll到累積={need_move:,.0f} → opex tie 6,829；capex+{need_move:,.0f}（總額不變）")
        L.append("     ⚠ 呢個係 proxy；若你 project payroll 表有 capex/opex 維度，用嗰個拆更準")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
