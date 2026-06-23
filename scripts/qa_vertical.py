r"""qa_vertical —— V 分類 QA 紅旗（user 2026-06-23）。v3：C 按 capex/opex 拆 + 加 NG↔V eligibility + 跨年一致性。

紅旗：
  A) 項目組V 非空但我哋=其他          —— 項目組有 ground truth，我哋反而其他。
  B) capex 行但 V=純活動類            —— capex 應 內部/外部設施（regression check，應近乎空）。
  C) 同一 (subproject, capex/opex) 跨 ≥2 V —— 已按 capex/opex 拆，免報「設施(capex)+活動(opex)」嗰種正確分流。
  D) keyword 同 V 矛盾（收窄）。
  E) NG↔V eligibility               —— NG-locked V（博彩→NG0、海上→NG10、康養→NG6、美食→NG8）出錯 NG。
  F) 跨年一致性                      —— 同 (subproject, capex/opex) 喺唔同報告年 dominant V 唔同（relabel 跨年唔穩）。

出 results\qa_vertical.txt。Run（Windows）:  python scripts\qa_vertical.py
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
TOPN = 12

ACT_V = {"演出表演", "路演", "體育賽事", "會展活動", "宣傳推廣", "康養活動", "節日慶典",
         "特別菜單或宴會", "文藝展覽表演", "美食-其他", "政府、公益及社區活動", "海上活動"}
# NG-locked V：呢啲 V 只應喺對應 NG（其餘 V 如演出/文藝/體育/會展可 cross NG，唔 flag）
NG_LOCK = {"博彩娛樂場優化": "NG0", "博彩設施設備優化": "NG0", "海上活動": "NG10",
           "康養活動": "NG6", "特別菜單或宴會": "NG8", "美食-其他": "NG8"}

CONCERT = (r"演唱會|演唱会|concert|\btour\b|fan\s*meeting|live\s*in|表演|演出|symphony|symphonic|festival|gala|"
           r"residency|show\b|匯演|滙演|音樂會|演藝|劇院|劇場|water\s*dance|水舞間|MGM\s*2049|2049|iqiyi|愛奇藝|"
           r"tmea|騰訊音樂|尖叫|超星|idol|偶像|music\s*live|night\b|盛典|頒奬|頒獎|awards|巡演|駐場|spectacle|演唱")
MICE = (r"會展|會議|展覽|展览|exhibition|\bMICE\b|conference|tradeshow|展會|論壇|forum|expo|convention|summit|"
        r"峰會|年會|annual\s*meeting|distributor\s*meeting|leadership|名匯|博覽|\bMITE\b|大會|招待會|圓桌|MDRT|meeting|品牌大會")


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _has(s, pat):
    return s.str.contains(pat, case=False, regex=True, na=False)


def main():
    L = ["# qa_vertical v3 —— V 分類 QA 紅旗（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    df["ent"] = df["entity"].astype(str)
    df["co"] = _s(df.get("final_capex_opex", pd.Series("", index=df.index))).map(lambda x: "Cap" if x == "Capex" else "Op")
    df["ry"] = df["year_bucket"].astype(str).str[:2]
    for c, n in [("vertical_label", "vl"), ("project", "pj"), ("subproject", "sp"),
                 ("account_desc", "ad"), ("description", "ds"), ("項目組V", "ptv"), ("ng_code", "ngc")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["blob"] = (df["sp"] + " " + df["ad"] + " " + df["ds"])

    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        L.append(f"\n{'='*78}\n## {ent}")

        # A) 項目組V 非空但我哋=其他
        a = e[e["ptv"].ne("") & e["vl"].eq("其他")]
        L.append(f"\n  ── A) 項目組V 有 label 但我哋=其他 top {TOPN} ──")
        if len(a) == 0:
            L.append("     （無）")
        for ptv, v in a.groupby("ptv")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(TOPN).items():
            L.append(f"     項組V={ptv[:30]:<30}{v:>10,.0f}萬")

        # B) capex 行但 V=純活動
        b = e[e["co"].eq("Cap") & e["vl"].isin(ACT_V)]
        L.append(f"  ── B) capex 行但 V=純活動類（應近乎空）top {TOPN} ──")
        if len(b) == 0:
            L.append("     （無）")
        for (pj, vl), v in b.groupby(["pj", "vl"])["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(TOPN).items():
            L.append(f"     [{pj[:30]:<30}] V={vl[:8]:<8}{v:>10,.0f}萬")

        # C) 同 (subproject, capex/opex) 跨 ≥2 V
        byc = e[e["sp"].ne("")].groupby(["sp", "co", "vl"])["amt"].sum().reset_index()
        flagged = []
        for (spv, co), sub in byc.groupby(["sp", "co"]):
            sig = sub[sub["amt"].abs() >= 50]
            if sig["vl"].nunique() >= 2:
                flagged.append((spv, co, sub["amt"].sum(), sig))
        flagged.sort(key=lambda t: -abs(t[2]))
        L.append(f"  ── C) 同 (subproject, {'Cap/Op'}) 跨 ≥2 V（各 ≥50萬）top {TOPN} ──")
        if not flagged:
            L.append("     （無）")
        for spv, co, tot, sub in flagged[:TOPN]:
            vs = " | ".join(f"{r.vl}={r.amt:,.0f}" for r in sub.itertuples())
            L.append(f"     [{spv[:28]:<28}|{co}] → {vs}")

        # D) keyword vs V
        d1 = e[e["vl"].eq("演出表演") & ~_has(e["blob"], CONCERT)]
        d2 = e[e["vl"].eq("會展活動") & ~_has(e["blob"], MICE)]
        L.append(f"  ── D) keyword vs V 矛盾（已剔走真演出/會展 keyword）──")
        for nm, s in [("V=演出但全冇演出keyword", d1), ("V=會展但全冇會展keyword", d2)]:
            if len(s) == 0:
                continue
            L.append(f"     ▸ {nm}：Σ={s['amt'].sum():,.0f}萬")
            for spv, v in s.groupby("sp")["amt"].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(4).items():
                if abs(v) < 30:
                    continue
                L.append(f"          {(spv or '(空)')[:44]:<46}{v:>9,.0f}萬")

        # E) NG↔V eligibility（NG-locked V 出錯 NG）
        e2 = e.copy()
        e2["lockng"] = e2["vl"].map(NG_LOCK)
        bad = e2[e2["lockng"].notna() & e2["lockng"].ne(e2["ngc"])]
        L.append(f"  ── E) NG↔V eligibility（NG-locked V 出錯 NG）──")
        if len(bad) == 0:
            L.append("     （無）")
        for (vl, ngc), v in bad.groupby(["vl", "ngc"])["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(8).items():
            if abs(v) < 10:
                continue
            L.append(f"     V={vl[:12]:<12} 喺 {ngc:<5}（應 {NG_LOCK.get(vl)}）{v:>10,.0f}萬")

        # F) 跨年一致性（同 (subproject, co) 唔同報告年 dominant V 唔同）
        dom = (e[e["sp"].ne("")].groupby(["sp", "co", "ry", "vl"])["amt"].sum().abs()
               .reset_index().sort_values("amt", ascending=False)
               .drop_duplicates(["sp", "co", "ry"]))   # 每 (sp,co,ry) 取最大金額 V
        fyr = []
        for (spv, co), sub in dom.groupby(["sp", "co"]):
            if sub["ry"].nunique() >= 2 and sub["vl"].nunique() >= 2:
                fyr.append((spv, co, sub))
        L.append(f"  ── F) 跨年：同 (subproject, Cap/Op) 唔同年 dominant V 唔同 top 8 ──")
        if not fyr:
            L.append("     （無）")
        for spv, co, sub in fyr[:8]:
            ys = " | ".join(f"{r.ry}:{r.vl}" for r in sub.sort_values("ry").itertuples())
            L.append(f"     [{spv[:30]:<30}|{co}] {ys}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
