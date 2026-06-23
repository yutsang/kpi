r"""check_2025_adj —— 對比新 {entity}_2025_adj.xlsx vs 現有 2025 raw 嘅欄位，判斷「直接換 / 要 vlookup」。

user 2026-06-23：會喺每個 data\{entity}\raw\ 放一個 {entity}_2025_adj.xlsx（含最新調整數）。
本工具：
  1) 讀現有 2025 raw（conf yearly_sources 嘅 tab）攞欄位（authoritative）；
  2) 讀 {entity}_2025_adj.xlsx（列所有 sheet + 欄）；
  3) 對比欄位：共有 / raw有但adj缺(→直接換唔到，要 vlookup) / adj多出；
  4) 檢查 kedro 關鍵欄（amount/調整/account/project/ng/capex-opex/年份filter）adj 有冇；
  5) 判 ✓直接換 / ⚠要vlookup。

Run（Windows）:  python scripts\check_2025_adj.py
出 results\check_2025_adj.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import warnings
warnings.filterwarnings("ignore")   # 靜音 openpyxl Print area 等 noise
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "check_2025_adj.txt"

# (檔名, 候選 tab：標準 adj tab 行先，揾唔到 fallback 現有 raw tab) — user 2026-06-23 標準化 tab 名
SRC = {
    "galaxy": ("galaxy_2025.xlsx", ["Combine(clean)"]),
    "sjm":    ("sjm_2025.xlsx",    ["表格1", "25", "24", "23"]),       # 標準=表格1(單tab);現有3-bucket → vlookup
    "wynn":   ("wynn_2025.xlsx",   ["報告投資支出明細賬", "Opex&Capex匯總"]),  # 標準改名
    "vml":    ("vml_2025.xlsx",    ["投資支出明細賬", "25JE"]),          # 標準改名
    "melco":  ("melco_2025.xlsx",  ["Data"]),
    "mgm":    ("mgm_2025.xlsx",    ["data"]),                          # mgm 手調，僅參考
}
# kedro 關鍵欄 hint（adj 一定要有，先可以直接換）
KEY = {
    "amount金額":   ["amount", "金額", "debit", "credit", "adjusted_amount", "amount - amended", "entry voucher", "expense amount", "25_amt"],
    "調整adjust":   ["調整", "adjust"],
    "account_code": ["account_code", "ledger_account_id", "ledger account", "科目代", "account code",
                     "a/c code", "cost element", "account", "科目"],
    "account_desc": ["account_desc", "ledger_account", "支出性質", "account desc", "科目名", "a/c name",
                     "nature of expense", "expense description", "account description", "cost element descr"],
    "project項目":  ["project", "項目", "initiative", "subproject", "project name", "project code"],
    "ng性質":       ["ng_theme", "項目性質", "項目類型", "section", "ng11", "ng_"],
    "capex_opex":   ["capex", "opex", "ledger type", "資本", "capex_opex"],
    "年份/bucket":  ["是否2025", "yr related", "year", "years", "report_year", "期後", "所屬年份", "24sy", "static_period"],
}


def _cols(path, sheet):
    try:
        df = pd.read_excel(path, sheet_name=sheet, nrows=3, dtype=object)
        return [str(c).strip() for c in df.columns]
    except Exception as e:
        return [f"!!讀取失敗:{e}"]


def main():
    L = ["# check_2025_adj —— adj 檔 vs 現有 2025 raw 欄位對比"]
    for ent, (fn, sheets) in SRC.items():
        raw = ROOT / "data" / ent / "raw" / fn
        adj = ROOT / "data" / ent / "raw" / f"{ent}_2025_adj.xlsx"
        L.append(f"\n{'='*72}\n## {ent}")
        if not raw.exists():
            L.append(f"   !! 現有 raw 揾唔到：{raw.relative_to(ROOT)}");
        rawcols = []
        if raw.exists():
            try:
                _avail = set(pd.ExcelFile(raw).sheet_names)
            except Exception as e:
                _avail = set(); L.append(f"   !! 開現有 raw 失敗：{e}")
            _used = next((sh for sh in sheets if sh in _avail), None)
            if _used is None:
                L.append(f"   !! 候選 tab {sheets} 一個都唔喺 {fn}（現有 tab：{sorted(_avail)}）")
            else:
                rawcols = list(dict.fromkeys(_cols(raw, _used)))
                _std = "✓標準" if _used == sheets[0] else f"⚠fallback(標準={sheets[0]})"
                L.append(f"   現有 raw：{fn}  用 tab=「{_used}」{_std}  欄數={len(rawcols)}")
                L.append(f"      現有欄位（adj 要跟住放）：{rawcols}")
        if not adj.exists():
            L.append(f"   ⏳ adj 未放：data\\{ent}\\raw\\{ent}_2025_adj.xlsx（放咗再跑）")
            continue
        try:
            adjsheets = pd.ExcelFile(adj).sheet_names
        except Exception as e:
            L.append(f"   !! adj 讀取失敗：{e}"); continue
        adj_tab = sheets[0]   # 標準 tab — adj 應該有呢張（唔好 union 全部 sheet）
        if adj_tab not in adjsheets:
            L.append(f"   ⚠ adj 冇標準 tab「{adj_tab}」→ 要喺 adj 整一張同名 tab，或話我改 SRC。")
            L.append(f"      adj sheets(前40)：{adjsheets[:40]}")
            continue
        adjcols = list(dict.fromkeys(_cols(adj, adj_tab)))
        L.append(f"   adj 檔：{ent}_2025_adj.xlsx  用 tab=「{adj_tab}」  欄數={len(adjcols)}")

        rs, as_ = set(rawcols), set(adjcols)
        missing = [c for c in rawcols if c not in as_]   # raw 有 adj 冇
        extra = [c for c in adjcols if c not in rs]
        L.append(f"   ── 欄位對比 ──")
        L.append(f"      共有={len(rs & as_)} | raw有adj冇={len(missing)} | adj多出={len(extra)}")
        if missing:
            L.append(f"      ⚠ adj 缺欄（直接換會走樣）：{missing[:20]}")
        if extra:
            L.append(f"      （adj 多出，通常 OK）：{extra[:15]}")
        # 關鍵欄 check（adj 側）
        L.append(f"   ── kedro 關鍵欄（adj 有冇）──")
        _low = [c.lower() for c in adjcols]
        allok = True
        for k, hints in KEY.items():
            hit = next((adjcols[i] for i, c in enumerate(_low) if any(h in c for h in hints)), None)
            mark = "✓" if hit else "✗缺"
            if not hit:
                allok = False
            L.append(f"      {mark} {k:<14}{hit or ''}")
        # 判定
        if not missing and allok:
            L.append(f"   ✅ 判定：欄位齊 + 關鍵欄全 → 可直接換（update conf raw_filename→adj 或覆蓋檔，re-run kedro）")
        else:
            L.append(f"   ⚠️ 判定：要 vlookup / 調整（adj 缺 {len(missing)} 欄" + ("，關鍵欄有缺" if not allok else "") + "）")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
