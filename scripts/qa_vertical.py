r"""qa_vertical —— V 分類 QA 紅旗（user 2026-06-23 想要 retest）。v2：減假警報 + 加更強驗證法。

紅旗（精準度由高至低）：
  A) 項目組V 非空但我哋 = 其他   —— 項目組有 ground truth label，我哋反而冚做其他（最該救）。
  B) capex 行但 V = 純活動類       —— capex（建設/設備）唔應該係 演出/路演/體育/會展/宣傳 等活動 V，
                                     通常係「建設場館」被標咗場館入面嘅活動（應 內部設施/建設）。
  C) 同一 subproject 跨 ≥2 個 V    —— 同一細項應該同一 V（項目37 嗰種分裂）。
  D) keyword 同 V 矛盾（已收窄，剔走真演出/真會展/真體育藝術 keyword，淨低先可疑）。

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

# 純活動 V（capex 唔應該係呢啲；應係 內部/外部設施 或 建設）
ACT_V = {"演出表演", "路演", "體育賽事", "會展活動", "宣傳推廣", "康養活動", "節日慶典",
         "特別菜單或宴會", "文藝展覽表演", "美食-其他", "政府、公益及社區活動", "海上活動"}

# 已擴充 keyword（reduce false positive）
CONCERT = (r"演唱會|演唱会|concert|\btour\b|fan\s*meeting|live\s*in|表演|演出|symphony|symphonic|festival|gala|"
           r"residency|show\b|匯演|滙演|音樂會|演藝|劇院|劇場|water\s*dance|水舞間|MGM\s*2049|2049|iqiyi|愛奇藝|"
           r"tmea|騰訊音樂|尖叫|超星|idol|偶像|music\s*live|night\b|盛典|頒奬|頒獎|awards|巡演|駐場|spectacle|演唱")
MICE = (r"會展|會議|展覽|展览|exhibition|\bMICE\b|conference|tradeshow|展會|論壇|forum|expo|convention|summit|"
        r"峰會|年會|annual\s*meeting|distributor\s*meeting|leadership|名匯|博覽|\bMITE\b|大會|招待會|圓桌|MDRT|"
        r"meeting|partnership|品牌大會|世界城市")
SPORT_ART = (r"golf|grand\s*prix|大賽車|tennis|網球|\bnba\b|\bfiba\b|volleyball|排球|全運|world\s*cup|olympic|"
             r"賽事|錦標|公開賽|雙年展|art\s*cura|藝術.*展|大展|cultural|體育|賽車|馬拉松|marathon|電競|esport")


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _has(s, pat):
    return s.str.contains(pat, case=False, regex=True, na=False)


def main():
    L = ["# qa_vertical v2 —— V 分類 QA 紅旗（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    df["ent"] = df["entity"].astype(str)
    df["cap"] = _s(df.get("final_capex_opex", pd.Series("", index=df.index))).eq("Capex")
    for c, n in [("vertical_label", "vl"), ("project", "pj"), ("subproject", "sp"),
                 ("account_desc", "ad"), ("description", "ds"), ("項目組V", "ptv")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["blob"] = (df["sp"] + " " + df["ad"] + " " + df["ds"])

    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        L.append(f"\n{'='*78}\n## {ent}")

        # ── A) 項目組V 非空但我哋=其他 ──
        a = e[e["ptv"].ne("") & e["vl"].eq("其他")]
        L.append(f"\n  ── A) 項目組V 有 label 但我哋=其他（最該救）top {TOPN} ──")
        if len(a) == 0:
            L.append("     （無）")
        else:
            for ptv, v in a.groupby("ptv")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(TOPN).items():
                L.append(f"     項組V={ptv[:30]:<30}{v:>10,.0f}萬")

        # ── B) capex 行但 V=純活動 ──
        b = e[e["cap"] & e["vl"].isin(ACT_V)]
        L.append(f"  ── B) capex 行但 V=純活動類（疑：建設場館誤標活動）top {TOPN} ──")
        if len(b) == 0:
            L.append("     （無）")
        else:
            for (pj, vl), v in b.groupby(["pj", "vl"])["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(TOPN).items():
                L.append(f"     [{pj[:30]:<30}] V={vl[:8]:<8}{v:>10,.0f}萬")

        # ── C) 同一 subproject 跨 ≥2 個 V ──
        bysp = e[e["sp"].ne("")].groupby(["sp", "vl"])["amt"].sum().reset_index()
        flagged = []
        for spv, sub in bysp.groupby("sp"):
            sig = sub[sub["amt"].abs() >= 50]
            if sig["vl"].nunique() >= 2:
                flagged.append((spv, sub["amt"].sum(), sig))
        flagged.sort(key=lambda t: -abs(t[1]))
        L.append(f"  ── C) 同一 subproject 跨 ≥2 個 V（各 ≥50萬）top {TOPN} ──")
        if not flagged:
            L.append("     （無）")
        for spv, tot, sub in flagged[:TOPN]:
            vs = " | ".join(f"{r.vl}={r.amt:,.0f}" for r in sub.itertuples())
            L.append(f"     [{spv[:30]:<30}] → {vs}")

        # ── D) keyword 同 V 矛盾（收窄）──
        d1 = e[e["vl"].eq("演出表演") & ~_has(e["blob"], CONCERT)]
        d2 = e[e["vl"].eq("會展活動") & ~_has(e["blob"], MICE)]
        d3 = e[_has(e["blob"], r"marketing|maketing|媒體推廣|數字營銷") & e["vl"].isin(["演出表演", "路演", "體育賽事"])]
        L.append(f"  ── D) keyword vs V 矛盾（已剔走真演出/會展 keyword）──")
        for nm, s in [("V=演出但全冇演出keyword", d1), ("V=會展但全冇會展keyword", d2), ("似marketing但V=演出/路演/體育", d3)]:
            if len(s) == 0:
                continue
            L.append(f"     ▸ {nm}：Σ={s['amt'].sum():,.0f}萬")
            for spv, v in s.groupby("sp")["amt"].sum().sort_values(key=lambda x: x.abs(), ascending=False).head(5).items():
                if abs(v) < 30:
                    continue
                L.append(f"          {(spv or '(空)')[:44]:<46}{v:>9,.0f}萬")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
