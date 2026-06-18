r"""一次過 inspect 用戶今輪提出嘅 NG/V 問題（讀 tableau_combined_25.csv，萬澳門元=調整後_萬）：

  §1 NG×V 掛錯枝（item 7）：V 唔喺該 NG eligible 集 → 列邊家/邊 NG/邊 V + 來源
      （project / ng_label(databook 範疇) / 項目組V / account_desc），等你 raw 改。
      重點 cell：NG5×會展活動、NG9×會展活動、NG9×美食-其他、NG9×文藝展覽表演。
  §2 sjm「政府、公益及社區活動」拆解（item 5c）：入面係咩 project/account。
  §3「內部設施-場館」by entity×NG（item 5 驗）：應該淨係 sjm NG9 有，其餘 NG9 應屬外部。
  §4 NG6 各家 project（item 6）：標出含 水療/SPA/中心/養生/wellness 嘅，睇邊啲應入 內部設施-康養。
  §5 vml「其他」拆解（item 10）：NG11其他 / H_OTHER / V其他 三個角度 top account。

Run（Windows）:  python scripts\inspect_ng_issues.py
出 results\inspect_ng_issues.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
CROSS = {"V_PROPERTY_UPGRADE", "V_PUBLIC_FACILITY", "V_COMMUNITY", "V_OTHER"}
L: list[str] = ["# inspect_ng_issues —— NG/V 問題逐項（萬澳門元 = 調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    for c in ("entity", "ng_code", "ng_label", "vertical_id", "vertical_label",
              "horizontal_label", "project", "subproject", "account_desc", "項目組V"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])

    # eligible_verticals
    try:
        from kpi.lib.conf import load_categories
        cats = load_categories()
    except Exception:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        cats = load_categories()
    elig = {ng: set(d.get("eligible_verticals") or []) | CROSS
            for ng, d in (cats.get("ng_categories") or {}).items()}

    # ── §1 NG×V 掛錯枝 ──
    L.append(f"\n{'='*78}\n## §1 NG×V 掛錯枝（V 唔喺該 NG eligible；按 |萬| 排）")
    stray = df[[(n in elig) and (v != "") and (v not in elig[n])
                for v, n in zip(df["vertical_id"], df["ng_code"])]].copy()
    if not len(stray):
        L.append("   （冇 stray 🎉）")
    else:
        g = (stray.groupby(["entity", "ng_code", "ng_label", "vertical_id", "vertical_label"])["amt"]
             .agg(["sum", "size"]).reset_index().sort_values("sum", key=lambda s: s.abs(), ascending=False))
        L.append(f"   {'entity':<7}{'NG':<5}{'ng_label':<10}{'V_id':<22}{'V_label':<14}{'萬':>9}{'行':>6}")
        for _, r in g.head(40).iterrows():
            L.append(f"   {r.entity:<7}{r.ng_code:<5}{r.ng_label[:9]:<10}{r.vertical_id[:21]:<22}{r.vertical_label[:13]:<14}{r['sum']:>9,.0f}{int(r['size']):>6}")
        L.append(f"   （共 {len(g)} 個 stray cell；總 {stray['amt'].sum():,.0f}萬）")
        # top cell 來源明細
        L.append(f"\n   ── top stray cell 來源（project | ng_label=databook範疇 | 項目組V | account）──")
        for _, r in g.head(8).iterrows():
            sub = stray[(stray.entity == r.entity) & (stray.ng_code == r.ng_code) & (stray.vertical_id == r.vertical_id)]
            L.append(f"   ▸ {r.entity} {r.ng_code} {r.vertical_label}（{r['sum']:,.0f}萬）：")
            d = (sub.groupby(["project", "項目組V", "account_desc"])["amt"].sum()
                 .reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False))
            for _, x in d.head(6).iterrows():
                L.append(f"       {x['amt']:>8,.0f}萬 | {x['project'][:30]:<30} | 組V={x['項目組V'][:12]:<12} | {x['account_desc'][:26]}")

    # ── §2 sjm 政府公益社區活動 ──
    L.append(f"\n{'='*78}\n## §2 sjm「政府、公益及社區活動」拆解")
    s2 = df[(df.entity == "sjm") & df.vertical_label.str.contains("政府", na=False)]
    L.append(f"   合計 {s2['amt'].sum():,.0f}萬，{len(s2)}行；by NG:")
    for ng, gg in s2.groupby("ng_label"):
        L.append(f"     {ng[:10]:<10} {gg['amt'].sum():,.0f}萬")
    L.append("   top project | account：")
    d = s2.groupby(["project", "account_desc"])["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
    for _, x in d.head(15).iterrows():
        L.append(f"     {x['amt']:>8,.0f}萬 | {x['project'][:34]:<34} | {x['account_desc'][:26]}")

    # ── §3 內部設施-場館 by entity×NG ──
    L.append(f"\n{'='*78}\n## §3「內部設施-場館」by entity×NG（應淨 sjm NG9 有；其餘 NG9 應外部）")
    s3 = df[df.vertical_label.str.contains("內部設施-場館", na=False)]
    g3 = s3.groupby(["entity", "ng_code", "ng_label"])["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
    for _, r in g3.iterrows():
        flag = "" if (r.entity == "sjm" and r.ng_code == "NG9") else ("  ← NG9 非sjm?" if r.ng_code == "NG9" else "")
        L.append(f"     {r.entity:<7}{r.ng_code:<5}{r.ng_label[:10]:<10}{r['amt']:>9,.0f}萬{flag}")

    # ── §4 NG6 各家 ──
    L.append(f"\n{'='*78}\n## §4 NG6（健康養生）各家 project — 標 水療/SPA/中心/養生/wellness")
    s4 = df[df.ng_code == "NG6"]
    KW = ("水療", "SPA", "Spa", "spa", "中心", "養生", "康養", "wellness", "Wellness", "WELLNESS", "溫泉", "理療")
    for ent, gg in s4.groupby("entity"):
        L.append(f"   ▸ {ent}（NG6 共 {gg['amt'].sum():,.0f}萬）：")
        d = gg.groupby(["project", "subproject", "vertical_label"])["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
        for _, x in d.head(10).iterrows():
            hit = "★" if any(k in (x['project'] + x['subproject']) for k in KW) else " "
            L.append(f"     {hit}{x['amt']:>8,.0f}萬 | {x['project'][:28]:<28} / {x['subproject'][:20]:<20} | V={x['vertical_label'][:12]}")

    # ── §5 vml 其他 ──
    L.append(f"\n{'='*78}\n## §5 vml「其他」拆解（3 角度）")
    v = df[df.entity == "vml"]
    for tag, mask in [("NG11(其他 bucket)", v.ng_label == "其他"),
                      ("H_OTHER(橫向其他)", v.horizontal_label == "其他"),
                      ("V=其他", v.vertical_label == "其他")]:
        sub = v[mask]
        L.append(f"\n   ── {tag}：{sub['amt'].sum():,.0f}萬，{len(sub)}行 — top account ──")
        d = sub.groupby("account_desc")["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
        for _, x in d.head(18).iterrows():
            L.append(f"     {x['amt']:>9,.0f}萬 | {x['account_desc'][:40]}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_ng_issues.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
