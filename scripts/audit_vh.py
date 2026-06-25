r"""audit_vh.py — 徹底審視最終 V(vertical)/H(horizontal) 分類合理性
Run: python scripts\audit_vh.py
In : tableau_combined_25.csv (CWD 根目錄；prep_tableau 出嘅最終 relabel 後檔)
Out: results\audit_vh.txt  +  results\audit_vh_<entity>_VxH.tsv

審查維度（per entity + 全體）：
  A. V 分布（行/金額萬/%）           ← V_OTHER 佔比 = 低信心警號
  B. H 分布（行/金額萬/%）           ← H_OTHER 佔比 = 低信心警號
  C. V→H breakdown（每個 V 嘅 H 成分%）← 核心：呢個 theme 嘅成本性質合唔合理
  D. H→top account_desc              ← 核實 H：呢個成本性質啲科目啱唔啱
  E. 其他 deep-dive（V_OTHER/H_OTHER top account_desc）← 睇咩跌入 catch-all
  F. capex/opex × H                  ← capex 應只入 建設/設施/人工/藝術品
  G. 紅旗 combos（硬規則可疑組合）
  H. 跨家一致性（同一 account_desc → 唔同 H）
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "audit_vh.txt"
CSV_CANDIDATES = [ROOT / "tableau_combined_25.csv", Path("tableau_combined_25.csv"),
                  ROOT / "data" / "tableau_combined_25.csv"]
ENTITIES = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
USECOLS = ["entity", "year_bucket", "amount_mop", "final_capex_opex", "ng_scope",
           "ng_label", "vertical_label", "horizontal_label", "account_code", "account_desc"]

# 紅旗：(H, 條件) — 喺呢個 H 下，唔屬以下 V 集合就可疑
PERF_V = {"娛樂表演", "文藝展覽表演", "演唱會", "體育賽事", "美食活動", "主題遊樂場地", "文化藝術盛事"}
CONSTRUCTION_V = {"博彩設施設備優化", "博彩娛樂場優化"}  # + 任何含「設施」字眼


def _acc(store, key, frame):
    store.setdefault(key, []).append(frame)


def _final(store, key, by):
    parts = store.get(key, [])
    if not parts:
        return pd.DataFrame(columns=by + ["amt", "rows"])
    df = pd.concat(parts, ignore_index=True)
    return df.groupby(by, as_index=False).agg(amt=("amt", "sum"), rows=("rows", "sum"))


def main():
    csv = next((c for c in CSV_CANDIDATES if c.exists()), None)
    if csv is None:
        OUT.parent.mkdir(exist_ok=True)
        OUT.write_text("!! tableau_combined_25.csv 揾唔到（請先 python scripts\\prep_tableau.py）", encoding="utf-8")
        print("!! csv 揾唔到"); return

    S: dict = {}
    n = 0
    for chunk in pd.read_csv(csv, usecols=lambda c: c in USECOLS, chunksize=300_000,
                             dtype=str, encoding="utf-8-sig"):
        chunk = chunk.rename(columns=lambda c: c.strip())
        chunk["amt"] = pd.to_numeric(chunk.get("amount_mop"), errors="coerce").fillna(0.0) / 1e4
        chunk["rows"] = 1
        for c in ("entity", "vertical_label", "horizontal_label", "final_capex_opex",
                  "account_desc", "ng_label"):
            if c not in chunk.columns:
                chunk[c] = ""
            chunk[c] = chunk[c].fillna("").astype(str).str.strip()
        n += len(chunk)
        _acc(S, "v",  chunk.groupby(["entity", "vertical_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        _acc(S, "h",  chunk.groupby(["entity", "horizontal_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        _acc(S, "vh", chunk.groupby(["entity", "vertical_label", "horizontal_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        _acc(S, "hd", chunk.groupby(["entity", "horizontal_label", "account_desc"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        _acc(S, "ch", chunk.groupby(["entity", "final_capex_opex", "horizontal_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        _acc(S, "gd", chunk.groupby(["account_desc", "horizontal_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        ov = chunk[chunk["vertical_label"] == "其他"]
        oh = chunk[chunk["horizontal_label"] == "其他"]
        if len(ov): _acc(S, "ov", ov.groupby(["entity", "account_desc", "ng_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        if len(oh): _acc(S, "oh", oh.groupby(["entity", "account_desc", "vertical_label"], as_index=False).agg(amt=("amt","sum"), rows=("rows","sum")))
        print(f"  ...{n:,} rows", flush=True)

    V  = _final(S, "v",  ["entity", "vertical_label"])
    H  = _final(S, "h",  ["entity", "horizontal_label"])
    VH = _final(S, "vh", ["entity", "vertical_label", "horizontal_label"])
    HD = _final(S, "hd", ["entity", "horizontal_label", "account_desc"])
    CH = _final(S, "ch", ["entity", "final_capex_opex", "horizontal_label"])
    GD = _final(S, "gd", ["account_desc", "horizontal_label"])
    OV = _final(S, "ov", ["entity", "account_desc", "ng_label"])
    OH = _final(S, "oh", ["entity", "account_desc", "vertical_label"])

    L = [f"# audit_vh — V/H 合理性審查  (rows={n:,}, 金額單位=萬MOP)", ""]

    for ent in ENTITIES:
        ve = V[V.entity == ent]; he = H[H.entity == ent]
        tot = ve.amt.sum()
        if tot == 0:
            continue
        L += ["", "█" * 72, f"█ {ent.upper()}   總額={tot:,.0f}萬"]

        v_other = ve[ve.vertical_label == "其他"].amt.sum()
        h_other = he[he.horizontal_label == "其他"].amt.sum()
        L.append(f"  ⚑ V_其他={v_other:,.0f}萬 ({v_other/tot*100:.1f}%)   H_其他={h_other:,.0f}萬 ({h_other/tot*100:.1f}%)")

        # A. V 分布
        L += ["", "── A. V 分布 (金額排) ──", f"   {'vertical':<16}{'金額萬':>12}{'%':>7}{'行數':>9}"]
        for _, r in ve.sort_values("amt", ascending=False).iterrows():
            L.append(f"   {r.vertical_label:<16}{r.amt:>12,.0f}{r.amt/tot*100:>6.1f}%{int(r.rows):>9,}")

        # B. H 分布
        L += ["", "── B. H 分布 (金額排) ──", f"   {'horizontal':<16}{'金額萬':>12}{'%':>7}{'行數':>9}"]
        for _, r in he.sort_values("amt", ascending=False).iterrows():
            L.append(f"   {r.horizontal_label:<16}{r.amt:>12,.0f}{r.amt/tot*100:>6.1f}%{int(r.rows):>9,}")

        # C. V→H breakdown
        L += ["", "── C. V→H 成分 (每個 V top4 H，%=V內佔比) ──"]
        vhe = VH[VH.entity == ent]
        for _, vr in ve.sort_values("amt", ascending=False).head(16).iterrows():
            sub = vhe[vhe.vertical_label == vr.vertical_label].sort_values("amt", ascending=False)
            vt = sub.amt.sum()
            if vt <= 0: continue
            parts = " | ".join(f"{r.horizontal_label}:{r.amt/vt*100:.0f}%" for _, r in sub.head(4).iterrows())
            L.append(f"   {vr.vertical_label:<16}({vr.amt:>10,.0f}萬)  {parts}")

        # D. H→top account_desc
        L += ["", "── D. H→top account_desc (核實 H；每 H top3) ──"]
        hde = HD[HD.entity == ent]
        for _, hr in he.sort_values("amt", ascending=False).head(16).iterrows():
            sub = hde[hde.horizontal_label == hr.horizontal_label].sort_values("amt", ascending=False).head(3)
            parts = " | ".join(f"{(r.account_desc or '<空>')[:22]}={r.amt:,.0f}" for _, r in sub.iterrows())
            L.append(f"   {hr.horizontal_label:<16} {parts}")

        # F. capex/opex × H
        che = CH[CH.entity == ent]
        L += ["", "── F. capex/opex × H (capex 應只入 建設/設施/人工/藝術品) ──"]
        for cap in sorted(che.final_capex_opex.unique()):
            sub = che[che.final_capex_opex == cap].sort_values("amt", ascending=False).head(8)
            ct = che[che.final_capex_opex == cap].amt.sum()
            parts = " | ".join(f"{r.horizontal_label}:{r.amt:,.0f}" for _, r in sub.iterrows())
            L.append(f"   {cap or '<空>':<8}({ct:,.0f}萬)  {parts}")

        # G. 紅旗 combos
        flags = []
        f1 = vhe[(vhe.horizontal_label == "藝人演出費") & (~vhe.vertical_label.isin(PERF_V))]
        if f1.amt.sum() != 0: flags.append(f"藝人演出費 in 非表演V: {f1.amt.sum():,.0f}萬 ({len(f1)}組) → " + ",".join(f1.sort_values('amt',ascending=False).head(3).vertical_label))
        cap_build = che[(che.final_capex_opex.str.lower() == "opex") & (che.horizontal_label == "建設與設施支出")]
        if cap_build.amt.sum() != 0: flags.append(f"建設與設施 in OPEX: {cap_build.amt.sum():,.0f}萬")
        f3 = vhe[(vhe.vertical_label.isin(CONSTRUCTION_V)) & (vhe.horizontal_label.str.startswith("Comp"))]
        if f3.amt.sum() != 0: flags.append(f"Comp* in 博彩設施/娛樂場V: {f3.amt.sum():,.0f}萬 ({len(f3)}組)")
        f4 = che[(che.final_capex_opex.str.lower()=="capex") & (~che.horizontal_label.isin(["建設與設施支出","設施及器具採購","人工成本","藝術品","維護費"]))]
        if f4.amt.sum() != 0: flags.append(f"capex 落非建設/設施/人工/藝術品/維護: {f4.amt.sum():,.0f}萬 top:" + ",".join(f4.sort_values('amt',ascending=False).head(3).horizontal_label))
        L += ["", "── G. 🚩 紅旗 ──"] + ([f"   🚩 {x}" for x in flags] or ["   （無觸發）"])

        # E. 其他 deep-dive
        ove = OV[OV.entity == ent].sort_values("amt", ascending=False).head(8)
        ohe = OH[OH.entity == ent].sort_values("amt", ascending=False).head(8)
        if len(ove):
            L += ["", "── E1. V_其他 top account_desc (咩跌入 V catch-all) ──"]
            for _, r in ove.iterrows():
                L.append(f"   {(r.account_desc or '<空>')[:30]:<32}{r.amt:>10,.0f}萬  NG={r.ng_label}")
        if len(ohe):
            L += ["", "── E2. H_其他 top account_desc (咩跌入 H catch-all) ──"]
            for _, r in ohe.iterrows():
                L.append(f"   {(r.account_desc or '<空>')[:30]:<32}{r.amt:>10,.0f}萬  V={r.vertical_label}")

        # per-entity V×H tsv
        piv = vhe.pivot_table(index="vertical_label", columns="horizontal_label", values="amt", aggfunc="sum", fill_value=0.0)
        (ROOT / "results").mkdir(exist_ok=True)
        piv.round(0).to_csv(ROOT / "results" / f"audit_vh_{ent}_VxH.tsv", sep="\t", encoding="utf-8-sig")

    # H. 跨家一致性：同一 account_desc 落 >1 個 H，按金額排
    L += ["", "█" * 72, "█ 跨家一致性：同一 account_desc → 多個 H (金額排 top30)"]
    g = GD[GD.account_desc.str.len() > 0].copy()
    nH = g.groupby("account_desc").horizontal_label.nunique().rename("nH")
    g = g.merge(nH, on="account_desc")
    multi = g[g.nH > 1].copy()
    tot_desc = multi.groupby("account_desc").amt.sum().rename("desc_amt")
    multi = multi.merge(tot_desc, on="account_desc")
    for desc in multi.sort_values("desc_amt", ascending=False).account_desc.drop_duplicates().head(30):
        sub = multi[multi.account_desc == desc].sort_values("amt", ascending=False)
        parts = " | ".join(f"{r.horizontal_label}:{r.amt:,.0f}" for _, r in sub.iterrows())
        L.append(f"   {(desc or '<空>')[:28]:<30} {parts}")

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L[:60]))
    print(f"\n... (全文 {len(L)} 行) wrote {OUT.relative_to(ROOT)} + results\\audit_vh_<ent>_VxH.tsv")


if __name__ == "__main__":
    main()
