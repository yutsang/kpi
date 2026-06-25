r"""inspect_capex_staff_hints.py — capex 行而家落咗非人工 H、但 desc/科目有人工線索
Run: python scripts\inspect_capex_staff_hints.py
In : tableau_combined_25.csv
Out: results\inspect_capex_staff_hints.txt
目的：capex 人工成本偏少 → 揾 capex 落咗 建設/設施/其他 H、但 account_desc/description 帶
人工關鍵詞嘅行，畀人 review 係咪可以撈返做人工。
"""
from __future__ import annotations
import sys, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "inspect_capex_staff_hints.txt"
CSV  = next((c for c in [ROOT/"tableau_combined_25.csv", Path("tableau_combined_25.csv")] if c.exists()), None)
COLS = ["entity", "amount_mop", "final_capex_opex", "horizontal_label",
        "account_code", "account_desc", "description", "vendor"]

# 人工線索（中英）；construction/設備 入面嘅 labour 都計（capex 資本化人工）
HINT = re.compile(r"payroll|salar|wage|staff\s*cost|staffing|head\s*count|secondment|secondee|"
                  r"manpower|labou?r|outsourc|sub-?contract\s*labou?r|人工|工資|薪|員工成本|"
                  r"人力|派遣|外判人手|加班|overtime|人員|工時|勞務", re.I)
LABOR_H = "人工成本"


def main():
    if CSV is None:
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! csv 揾唔到", encoding="utf-8"); print("no csv"); return
    cur_labor = {}   # entity → capex 人工 現額
    hit_eh = {}; hit_ad = {}; hit_desc = {}
    n_hit = 0; amt_hit = 0.0
    for ch in pd.read_csv(CSV, usecols=lambda c: c in COLS, chunksize=300_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        ch["amt"] = pd.to_numeric(ch["amount_mop"], errors="coerce").fillna(0.0) / 1e4
        ch["cap"] = ch["final_capex_opex"].fillna("").astype(str).str.strip().str.lower()
        cap = ch[ch["cap"] == "capex"].copy()
        if not len(cap): continue
        # 現有 capex 人工 baseline
        lab = cap[cap["horizontal_label"] == LABOR_H].groupby("entity")["amt"].sum()
        for e, v in lab.items(): cur_labor[e] = cur_labor.get(e, 0.0) + v
        # 候選：capex 非人工 H 但帶人工線索
        cand = cap[cap["horizontal_label"] != LABOR_H].copy()
        txt = (cand["account_desc"].fillna("") + " ｜ " + cand["description"].fillna("")).astype(str)
        cand = cand[txt.str.contains(HINT, na=False)]
        if not len(cand): continue
        cand["rows"] = 1
        n_hit += len(cand); amt_hit += cand["amt"].sum()
        for key, by in [("eh",["entity","horizontal_label"]),
                        ("ad",["entity","account_code","account_desc"]),
                        ("ds",["entity","description"])]:
            tgt = {"eh":hit_eh,"ad":hit_ad,"ds":hit_desc}[key]
            g = cand.groupby(by, as_index=False).agg(amt=("amt","sum"), rows=("rows","sum"))
            tgt.setdefault("_", []).append(g)
        print(f"  hits: {n_hit:,} 行 / {amt_hit:,.0f}萬", flush=True)

    L = ["# capex 人工線索（capex 落非人工H 但 desc/科目帶人工字眼）",
         f"命中：{n_hit:,} 行 / {amt_hit:,.0f}萬", ""]
    L += ["── 現有 capex 人工成本 baseline (每家) ──"]
    for e in sorted(cur_labor): L.append(f"   {e:<8} {cur_labor[e]:>10,.0f}萬")
    def dump(store, title, keys, topn=40):
        parts = store.get("_", [])
        if not parts: return
        df = pd.concat(parts, ignore_index=True)
        g = df.groupby(keys, as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")).sort_values("amt", ascending=False)
        L.append(""); L.append(f"── {title} ──")
        for _, r in g.head(topn).iterrows():
            k = " | ".join(str(r[c])[:40] for c in keys)
            L.append(f"   {k:<52} {r['amt']:>10,.0f}萬 {int(r['rows']):>8,}行")
    dump(hit_eh, "命中 by entity × 現有H（撈得返嘅來源）", ["entity","horizontal_label"])
    dump(hit_ad, "命中 top account_code+desc", ["entity","account_code","account_desc"], 50)
    dump(hit_desc, "命中 top description", ["entity","description"], 40)
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L[:50])); print(f"\n... wrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
