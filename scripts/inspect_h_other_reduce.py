r"""inspect_h_other_reduce —— 把每家 H_OTHER 逐類拆「可移真H vs 結構性非投資」，計移完之後 %。

user 2026-06-23：唔郁 comp/staff 下想 H_OTHER < 2%。
結構性（唔郁得，移咗就搞 staff/comp 或屈數）：超額staff、博彩contra、分攤/allocation、運輸travel、稅、food cost、pre-open。
可移（光明正大 → 真 H，非 comp 非 staff）：special event/promo→廣告、service fee→專業、supplies→設施。
→ 逐家 dump 每類 Σ + 樣本 account_desc，最後計「結構性 floor %」同「移完可移之後 %」。

Run（Windows）:  python scripts\inspect_h_other_reduce.py
出 results\inspect_h_other_reduce.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_h_other_reduce.txt"
ENTS = ["galaxy", "mgm", "melco", "vml", "wynn", "sjm"]

# (類別, 是否結構性, regex on account_desc + 項組H + account_code)
CATS = [
    ("結構:超額staff",      True,  r"salar|payroll|casual\s*labor|\bbonus\b|housing|medical|allowance|incentive|staff\s*cost|薪|工資|員工福利|人工|wages|8000|8020"),
    ("結構:博彩contra/gaming", True, r"contra|table\s*games|slot\s*revenue|gaming\s*credit|博彩|410\d{3}|412\d{3}|521010"),
    ("結構:分攤/allocation/recovery", True, r"allocation|interdept|interco|分攤|cost\s*recovery|\b89000|\b89001|841000|841005|611199|650550"),
    ("結構:運輸/travel",    True,  r"transport|travel|limo|舟車|運輸|airfare|lodging|ferry|交通|624110|80570|80580|590010|7153|8103"),
    ("結構:稅tax",          True,  r"\btax\b|稅|government\s*tax|tourism\s*tax|563715|7151020"),
    ("結構:food/COS(近comp)", True, r"food\s*cost|cost\s*of\s*sales|foodcost|食品成本|512010|7031|51000|511110"),
    ("結構:pre-open",       True,  r"pre[\s-]*open|開業前|8800400"),
    ("可移→廣告(events/promo)", False, r"special\s*event|promotion|\bmedia\b|廣告|宣傳|event\s*service|advertis|7710100|8107500|special\s*events"),
    ("可移→專業(service/prof)", False, r"service\s*fee|service\s*exp|professional|顧問|專業|consult|90010|90011"),
    ("可移→設施(supplies/equip)", False, r"supplies|guest\s*supplies|\bO/E\b|消耗|器具|equipment|7931|7932"),
]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_h_other_reduce —— H_OTHER 可移 vs 結構性（萬；amount_mop/1e4）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    df["ent"] = df["entity"].astype(str)
    hid = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    for c, n in [("account_desc", "ad"), ("項目組H", "pth"), ("account_code", "code")]:
        df[n] = _s(df.get(c, pd.Series("", index=df.index)))
    df["blob"] = (df["ad"] + " " + df["pth"] + " " + df["code"])

    for ent in ENTS:
        e = df[df["ent"].eq(ent)]
        tot = e["amt"].sum()
        o = e[hid.loc[e.index].eq("H_OTHER")].copy()
        so = o["amt"].sum()
        L.append(f"\n{'='*74}\n## {ent}  H_OTHER={so:,.0f} / 總={tot:,.0f} = {so/tot*100:.2f}%")
        if len(o) == 0:
            continue
        _used = pd.Series(False, index=o.index)
        struct_sum = 0.0; move_sum = 0.0
        for name, is_struct, pat in CATS:
            m = o["blob"].str.contains(pat, case=False, regex=True, na=False) & ~_used
            _used = _used | m
            v = o.loc[m, "amt"].sum()
            if abs(v) < 1 and int(m.sum()) == 0:
                continue
            if is_struct:
                struct_sum += v
            else:
                move_sum += v
            tag = "🔒" if is_struct else "✅可移"
            L.append(f"   {tag} {name:<26}{v:>10,.0f}萬  {int(m.sum()):>6}行")
            # 樣本
            for d, dv in o.loc[m].groupby("ad")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(2).items():
                L.append(f"          {(d or '(空)')[:38]:<40}{dv:>9,.0f}")
        rest = o.loc[~_used, "amt"].sum()
        L.append(f"   ❓ 未分類(review)              {rest:>10,.0f}萬  {int((~_used).sum()):>6}行")
        for d, dv in o.loc[~_used].groupby("ad")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False).head(4).items():
            L.append(f"          {(d or '(空)')[:38]:<40}{dv:>9,.0f}")
        # 計移完之後 %
        after = so - move_sum
        L.append(f"   ── 小結:結構性floor={struct_sum:,.0f} | 可移={move_sum:,.0f} | review={rest:,.0f}")
        L.append(f"      移完可移之後 H_OTHER ≈ {after:,.0f}萬 = {after/tot*100:.2f}%（結構性floor = {struct_sum/tot*100:.2f}%）")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
