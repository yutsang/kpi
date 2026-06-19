r"""驗證 data\vml\raw\vml_2025_new.xlsx (tab 25JE) 能唔能順利過 pipeline，再決定要唔要重跑 kedro vml。

VML pipeline 2025 讀 raw_filename=vml_2025.xlsx, raw_sheet=25JE。檢查新檔：
  ① pipeline 硬性要嘅欄齊唔齊（amount/account/desc/project/ng11/capex_opex/year… 缺 = 爆 or amount=0）
  ② MOP Amt 係咪 numeric + 總額非 0（amount=0 bug）
  ③ manual mark 欄（分類1/分類1.1/分類2/pt_class_V/投資領域/會計科目分類/Nature/Payroll）齊唔齊 + 值分佈
  ④ 分類1(H) / 分類2+pt_class_V(V) 嘅值有幾多 map 得到（conf column_map keys）vs map 唔到（→H_OTHER/blank V）
  ⑤ 對比舊檔 vml_2025.xlsx：欄 added/removed、row 數、分類1/2 值分佈變化

Run（Windows）:  python scripts\validate_vml_new.py
                 python scripts\validate_vml_new.py <new.xlsx> <old.xlsx>
出 results\validate_vml_new.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
NEW = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data/vml/raw/vml_2025_new.xlsx"
OLD = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "data/vml/raw/vml_2025.xlsx"
SHEET = "25JE"
L: list[str] = ["# validate_vml_new —— 新 vml_2025_new.xlsx (25JE) 能唔能過 pipeline"]

# pipeline 硬性要嘅欄（缺 = 爆 / amount=0）。alts = 容許嘅替代欄名。
REQUIRED = {
    "MOP Amt": ("amount(numeric!)", []), "Account": ("account_code", ["SUB"]),
    "A/C Name": ("account_desc", []), "Transaction Description": ("description", ["Journal Description"]),
    "SubProject_Name": ("project", ["Subproject", "項目名稱"]), "項目性質": ("ng11_category", []),
    "Ref number": ("unique_id", []), "CAPEX_OPEX": ("capex_opex", []),
    "Year": ("posting_year", []), "Month": ("period", []), "Posting Date": ("posting_date", []),
}
MANUAL = ["分類1", "分類1.1", "分類2", "pt_class_V", "pt_class_H", "投資領域",
          "會計科目分類", "Nature", "Payroll", "Job code"]
NICE = ["Vendor Short Name", "Commodity code", "Invoice No", "Inv Ref", "PO number", "ARPO#",
        "Journal No", "Subproject", "項目名稱", "調整金額"]


def _mapkeys(ov, col):
    keys = set()
    for r in (ov or []):
        cm = r.get("column_map") or {}
        if cm.get("col") == col:
            keys |= set((cm.get("map") or {}).keys())
    return keys


def _read(p):
    try:
        return pd.read_excel(p, sheet_name=SHEET, dtype=str)
    except Exception as e:
        return e


def _present(df, name, alts):
    cols = set(df.columns)
    if name in cols:
        return name
    for a in alts:
        if a in cols:
            return a
    # fuzzy strip
    for c in df.columns:
        if str(c).strip() == name.strip():
            return c
    return None


def _nb(df, c):
    s = df[c].astype(str).str.strip()
    return int((~s.isin(["", "nan", "None", "NaN", "<NA>"])).sum())


def main():
    cfg = yaml.safe_load((ROOT / "conf/company_4/parameters.yml").read_text(encoding="utf-8"))
    H_KEYS = _mapkeys(cfg.get("row_horizontal_overrides"), "分類1") | _mapkeys(cfg.get("row_horizontal_overrides"), "分類1.1")
    V_KEYS = _mapkeys(cfg.get("row_vertical_overrides"), "pt_class_V") | _mapkeys(cfg.get("row_vertical_overrides"), "分類2")
    L.append(f"  conf H-map(分類1) keys={len(H_KEYS)}、V-map(分類2/pt_class_V) keys={len(V_KEYS)}")

    if not NEW.exists():
        L.append(f"\n❌ 新檔唔存在：{NEW}"); _w(); return
    new = _read(NEW)
    if isinstance(new, Exception):
        L.append(f"\n❌ 開唔到新檔 {NEW.name} [{SHEET}]：{new}"); _w(); return
    L.append(f"\n## 新檔 {NEW.name} [{SHEET}]：{len(new):,} 行, {len(new.columns)} 欄")

    fail = []
    L.append(f"\n── ① 硬性欄 ──")
    for name, (role, alts) in REQUIRED.items():
        hit = _present(new, name, alts)
        if hit:
            L.append(f"   ✓ {name:<22}({role}) {'' if hit==name else '← 用替代 '+hit}  非空 {_nb(new,hit):,}")
        else:
            L.append(f"   ❌ 缺 {name:<22}({role})  alts={alts}")
            fail.append(name)

    L.append(f"\n── ② MOP Amt numeric ──")
    amtc = _present(new, "MOP Amt", [])
    if amtc:
        v = pd.to_numeric(new[amtc].astype(str).str.replace(",", "", regex=False), errors="coerce")
        L.append(f"   numeric 行 {int(v.notna().sum()):,}/{len(new):,}；Σ={v.sum()/1e4:,.0f}萬；blank/非數 {int(v.isna().sum()):,}")
        if v.sum() == 0 or v.notna().sum() == 0:
            L.append("   ❌ 金額全 0 / 非數 → amount=0 bug"); fail.append("MOP Amt 非數")

    L.append(f"\n── ③ manual mark 欄 ──")
    for c in MANUAL:
        hit = _present(new, c, [])
        if hit:
            vals = new[hit].astype(str).str.strip().replace({"nan": "", "None": ""})
            top = vals[vals.ne("")].value_counts().head(6)
            L.append(f"   ✓ {c:<14} 非空 {_nb(new,hit):,}  top: " + " | ".join(f"{k[:14]}={n}" for k, n in top.items()))
        else:
            L.append(f"   ⚠ 冇 {c}（若舊檔有 → 影響 V/H 分類）")

    L.append(f"\n── ④ 分類值 map coverage（map 唔到 → H_OTHER/blank V）──")
    for col, keys, lab in [("分類1", H_KEYS, "H"), ("分類1.1", H_KEYS, "H"), ("分類2", V_KEYS, "V"), ("pt_class_V", V_KEYS, "V")]:
        hit = _present(new, col, [])
        if not hit:
            continue
        vals = new[hit].astype(str).str.strip().replace({"nan": "", "None": ""})
        nonblank = vals[vals.ne("")]
        unmapped = sorted(set(nonblank) - set(keys))
        ok = int(nonblank.isin(keys).sum())
        L.append(f"   {col}({lab}) 非空 {len(nonblank):,}：map 到 {ok:,}；map 唔到 {len(nonblank)-ok:,}")
        if unmapped:
            L.append(f"       ⚠ 未知值（conf 冇 map，會落 {'H_OTHER' if lab=='H' else 'blank V'}）：" + " | ".join(str(x)[:16] for x in unmapped[:12]))

    L.append(f"\n── ⑤ 對比舊檔 {OLD.name} ──")
    if OLD.exists():
        old = _read(OLD)
        if isinstance(old, Exception):
            L.append(f"   開唔到舊檔：{old}")
        else:
            nc, oc = set(new.columns), set(old.columns)
            L.append(f"   舊 {len(old):,} 行 / 新 {len(new):,} 行（Δ {len(new)-len(old):+,}）")
            add = sorted(nc - oc); rem = sorted(oc - nc)
            L.append(f"   新增欄 {len(add)}：" + (" | ".join(str(x)[:18] for x in add[:15]) or "(冇)"))
            L.append(f"   消失欄 {len(rem)}：" + (" | ".join(str(x)[:18] for x in rem[:15]) or "(冇)"))
            if rem:
                L.append("       ⚠ 消失欄若係 pipeline 要嘅 → 會出事")
    else:
        L.append(f"   （舊檔唔存在 {OLD}，跳過對比）")

    L.append(f"\n{'='*60}")
    if fail:
        L.append(f"❌ 結論：唔過得 —— 問題：{fail}")
    else:
        L.append("✅ 結論：硬性欄齊 + 金額 OK → 可以過 pipeline。")
        L.append("   用法：備份舊 vml_2025.xlsx → 將 vml_2025_new.xlsx rename 做 vml_2025.xlsx →")
        L.append("        del data\\vml\\interim\\company_4_raw.parquet → set KPI_SKIP_STEP3=1 →")
        L.append("        kedro run --pipeline=vml → set KPI_SKIP_STEP3= → python scripts\\prep_tableau.py")
        L.append("   （若 ③/④ 有未知 mark 值 → 嗰啲會落 H_OTHER/blank V，睇下使唔使補 conf map）")
    _w()


def _w():
    out = ROOT / "results" / "validate_vml_new.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
