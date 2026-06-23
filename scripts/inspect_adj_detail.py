r"""inspect_adj_detail —— galaxy/wynn/melco：kedro 實際需要嘅 source 欄，喺 adj 標準 tab 揾（對名+alias），
睇「缺欄」係咪只係改咗名。+ dump adj 全部欄 + 關鍵欄 sample 值。

user 2026-06-23：缺欄可能只係 column 名唔同，要仔細 inspect。
kedro 要嘅欄由 conf columns:/columns_override/row_horizontal_overrides 攞。

Run（Windows）:  python scripts\inspect_adj_detail.py
出 results\inspect_adj_detail.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import warnings
warnings.filterwarnings("ignore")
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "inspect_adj_detail.txt"

# (adj 檔, 標準 tab)
ADJ = {
    "galaxy": ("galaxy_2025_adj.xlsx", "Combine(clean)"),
    "wynn":   ("wynn_2025_adj.xlsx",   "報告投資支出明細賬"),
    "melco":  ("melco_2025_adj.xlsx",  "Data"),
}
# kedro 要嘅 logical 欄 + alias（lowercase substring 比對）
REQ = {
    "galaxy": [
        ("amount金額",        ["amount_mop", "reported amount", "金額", "expense amount"]),
        ("調整",              ["adjustment", "調整", "一級調整", "二級調整"]),
        ("capex_opex",        ["capex_opex", "capex/opex", "capex / opex", "資本"]),
        ("project",           ["project", "項目", "name of investment"]),
        ("ng11性質",          ["ng_theme", "ng11 category", "項目性質", "ng category"]),
        ("account_code",      ["account_code", "account code", "a/c code", "科目代"]),
        ("account_desc",      ["account_desc", "account description", "nature of expense", "科目名"]),
        ("description",       ["description", "摘要", "memo"]),
        ("year年份filter",    ["posting year", "是否2025", "\\byear\\b"]),
        ("unique_id",         ["unique_id", "唯一識別", "submit no", "index"]),
        ("vendor",            ["vendor", "supplier", "供應商"]),
        ("⚑H分類:二級標簽",   ["二級標簽", "基礎|二級", "二級標籤"]),
        ("⚑H分類:一級標簽",   ["一級標簽", "基礎|一級", "一級標籤"]),
        ("⚑H分類:Source",     ["source", "來源"]),
    ],
    "wynn": [
        ("amount金額",        ["entry voucher", "expense amount", "金額"]),
        ("調整",              ["調整金額", "adjustment"]),
        ("capex_opex",        ["capex / opex", "capex/opex", "capex_opex"]),
        ("project",           ["name of investment project", "項目名"]),
        ("project中文/sub",   ["项目名称中文", "sub project", "subproject"]),
        ("ng11性質",          ["項目性質", "項目類型"]),
        ("account_code",      ["\\baccount\\b", "a/c", "科目", "entry account"]),
        ("account_desc",      ["nature of expense", "account desc", "expense desc"]),
        ("vendor",            ["name of supplier", "supplier"]),
        ("unique_id",         ["index", "唯一識別"]),
        ("year年份filter",    ["是否2025", "年度"]),
        ("⚑comp分類:comp费用大类", ["comp费用大类", "comp大类", "comp費用", "comp大類"]),
    ],
    "melco": [
        ("amount金額",        ["amount - amended", "amount amended", "調整金額"]),
        ("調整",              ["調整數", "調整金額", "初步識別調整"]),
        ("capex_opex",        ["ledger type", "\\bcapex\\b"]),
        ("project",           ["project name - amended", "項目名稱", "项目名称"]),
        ("ng11性質",          ["項目性質"]),
        ("account_code",      ["ledger_account_id", "ledger account", "wd账户", "wd ledger"]),
        ("account_desc",      ["ledger_account", "支出性質"]),
        ("description",       ["line_memo", "spend_category"]),
        ("unique_id",         ["kpmg唯一識別", "唯一識別"]),
        ("comp_nature",       ["comp性質-cn", "comp性質"]),
        ("comp分類",          ["comp性質分類"]),
        ("payroll人工",       ["kp識別人工", "識別人工", "人工成本分類"]),
        ("⚑H分類:建築設施設備劃分", ["建築設施設備劃分", "設施設備劃分", "區分設施", "設施建設", "建設與設施"]),
        ("⚑H分類:KP识别表演演出", ["kp识别表演演出", "識別表演", "表演演出", "kp识别表演"]),
    ],
}


def _find(cols, aliases):
    import re
    low = [str(c).lower() for c in cols]
    for a in aliases:
        a = a.lower()
        for i, c in enumerate(low):
            if a.startswith("\\b") or "\\b" in a:
                if re.search(a, c):
                    return cols[i]
            elif a in c:
                return cols[i]
    return None


def main():
    L = ["# inspect_adj_detail —— kedro 要嘅 source 欄喺 adj 揾（對名+alias）"]
    for ent, (fn, tab) in ADJ.items():
        adj = ROOT / "data" / ent / "raw" / fn
        L.append(f"\n{'='*74}\n## {ent}  adj={fn}  tab={tab}")
        if not adj.exists():
            L.append(f"   !! adj 揾唔到"); continue
        try:
            df = pd.read_excel(adj, sheet_name=tab, dtype=object, nrows=8)
        except Exception as e:
            L.append(f"   !! 讀 tab 失敗：{e}"); continue
        cols = [str(c).strip() for c in df.columns]
        L.append(f"   adj 欄數={len(cols)}")
        L.append(f"   ── kedro 要嘅欄 → adj 揾到? ──")
        miss = []
        for name, aliases in REQ[ent]:
            hit = _find(cols, aliases)
            if hit is None:
                miss.append(name); L.append(f"      ✗缺  {name:<22}（alias:{aliases[:3]}）")
            else:
                # sample 2 非空值
                try:
                    sv = df[hit].dropna().astype(str).head(2).tolist()
                except Exception:
                    sv = []
                L.append(f"      ✓    {name:<22}→「{hit}」  e.g. {sv}")
        crit = [m for m in miss if m.startswith("⚑")]
        L.append(f"   ── 判定：缺 {len(miss)} 欄" + (f"，當中關鍵分類欄缺：{crit}" if crit else "（基本欄齊）"))
        L.append(f"   ── adj 全部欄（揾改名用）──\n      {cols}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
