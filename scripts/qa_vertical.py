r"""qa_vertical —— V(vertical)分類 QA，自動 spot 可疑分類（user 2026-06-23 想要 retest）。

兩類紅旗：
  A) 同一 project 嘅 V 分裂（一個 project 散落 ≥2 個 V，每個都有顯著金額）→ 似 項目37 嗰種 bug。
  B) account_desc/subproject keyword 同 V 矛盾，例如：
     - V=路演 但內容係 推廣費/廣告（冇 roadshow/展銷/tradeshow）
     - V=演出表演 但內容冇 演唱會/concert/show/tour/表演
     - 內容有 推廣/廣告/marketing 但 V∈{路演,體育賽事,文藝展覽表演}（疑似應宣傳）
     - V=會展活動 但內容冇 會展/MICE/exhibition/展
出 results\qa_vertical.txt — 逐家列 top 紅旗，畀 user 一眼掃。

Run（Windows）:  python scripts\qa_vertical.py
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "qa_vertical.txt"
ENTS = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
SPLIT_MIN = 50.0   # 萬：一個 V 要 ≥ 呢個金額先算「顯著」
TOPN = 15


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# qa_vertical —— V 分類 QA 紅旗（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    df["ent"] = df["entity"].astype(str)
    for c, n in [("vertical_label", "vl"), ("project", "pj"), ("subproject", "sp"),
                 ("account_desc", "ad"), ("description", "ds"), ("ng_label", "ng")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["blob"] = (df["sp"] + " " + df["ad"] + " " + df["ds"])

    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        L.append(f"\n{'='*78}\n## {ent}")

        # ── A) project V 分裂 ──
        by = e.groupby(["pj", "vl"])["amt"].sum().reset_index()
        flagged = []
        for pj, sub in by.groupby("pj"):
            if pj == "":
                continue
            sig = sub[sub["amt"].abs() >= SPLIT_MIN]
            if sig["vl"].nunique() >= 2:
                flagged.append((pj, sub["amt"].sum(), sub.sort_values("amt", key=lambda s: s.abs(), ascending=False)))
        flagged.sort(key=lambda t: -abs(t[1]))
        L.append(f"\n  ── A) project V 分裂（≥2 個 V 各 ≥{SPLIT_MIN:.0f}萬）top {TOPN} ──")
        if not flagged:
            L.append("     （無）")
        for pj, tot, sub in flagged[:TOPN]:
            vs = " | ".join(f"{r.vl}={r.amt:,.0f}" for r in sub.itertuples() if abs(r.amt) >= SPLIT_MIN)
            L.append(f"     [{pj[:34]:<34}] Σ={tot:,.0f}萬 → {vs}")

        # ── B) keyword 同 V 矛盾 ──
        def _has(s, pat):
            return s.str.contains(pat, case=False, regex=True, na=False)
        promo_kw = r"推廣|廣告|宣傳|promotion|marketing|advertis|媒體|media|網站|website|品牌|brand"
        road_kw = r"road\s*show|roadshow|trade\s*show|tradeshow|路演|展銷|展销"
        concert_kw = r"演唱會|演唱会|concert|\btour\b|fan\s*meeting|live\s*in|表演|演出|symphony|festival|gala|residency|show\b"
        mice_kw = r"會展|會議|展覽|展览|exhibition|\bMICE\b|conference|tradeshow|展會|論壇|forum|expo"

        checks = [
            ("V=路演但內容似推廣(非roadshow)", e["vl"].eq("路演") & ~_has(e["blob"], road_kw) & _has(e["blob"], promo_kw)),
            ("V=演出表演但冇演出keyword", e["vl"].eq("演出表演") & ~_has(e["blob"], concert_kw)),
            ("V=會展活動但冇會展keyword", e["vl"].eq("會展活動") & ~_has(e["blob"], mice_kw)),
            ("內容似推廣但V=路演/體育/文藝", _has(e["blob"], promo_kw) & e["vl"].isin(["路演", "體育賽事", "文藝展覽表演"]) & ~_has(e["blob"], road_kw)),
        ]
        L.append(f"  ── B) keyword 同 V 矛盾（每類 top 6 subproject）──")
        for name, m in checks:
            s = e[m]
            if len(s) == 0:
                continue
            tot = s["amt"].sum()
            L.append(f"     ▸ {name}：{len(s):,}行 Σ={tot:,.0f}萬")
            for spv, v in s.groupby("sp")["amt"].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(6).items():
                if abs(v) < 10:
                    continue
                L.append(f"          {(spv or '(空)')[:46]:<48}{v:>9,.0f}萬")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
