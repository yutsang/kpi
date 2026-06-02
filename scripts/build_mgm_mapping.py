"""MGM 手動拼表助手 — 讀 MGM master Excel（Master 左 + Supporting 右，side-by-side），
join（項目序號 ↔ Item Code 入面個號），套 mapping（Project Session→NG、WD欄→H），
出一張 per-項目 mapping TSV 畀你喺 Windows 拼表用。

Master 欄（左）：項目序號 / 項目名稱 / Payroll / CAPEX / OPEX / Total
Supporting 欄（右）：Item Code / Project Name / Project Session / Investment Type /
                     WD1 / WD2-Allocation / WD3-Payroll / WD4-COGS / Patron Management / Total

Run (Windows):
  python scripts/build_mgm_mapping.py --file "mgm_master_25.xlsx" --sheet Master
  python scripts/build_mgm_mapping.py --file "mgm_master_25.xlsx" --inspect   # 先睇欄名/header row
Output: results/mgm_mapping.tsv
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent

# Project Session 名 → (NG code, NG label)。25 用「B1吸引外國客源」；23/24 用「吸引外國客源 - 活動」
NG_BY_LABEL = {"吸引外國客源": 1, "會議展覽": 2, "娛樂表演": 3, "體育盛事": 4, "文化藝術": 5,
               "健康養生": 6, "主題遊樂": 7, "美食之都": 8, "社區旅遊": 9, "海上旅遊": 10, "其他": 11}

# WD / 成本欄 → H bucket（WD1 最大宗，要 account 明細先細拆，暫標「待拆WD1」）
WD_TO_H = {"WD3": "人工成本", "WD4": "餐飲", "Patron": "Comp其他",
           "WD2": "人工成本", "WD1": "待拆WD1(廣告/合約/贊助/comp)"}


def ng_of(session):
    s = str(session)
    for lbl, ng in NG_BY_LABEL.items():
        if lbl in s:
            return f"NG{ng}", lbl
    return "", ""


def find_col(cols, *kw):
    for k in kw:
        for c in cols:
            if k.lower() in str(c).lower():
                return c
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--sheet", default=0)
    ap.add_argument("--inspect", action="store_true")
    args = ap.parse_args()
    fp = ROOT / args.file
    if not fp.exists(): fp = Path(args.file)
    if not fp.exists():
        print(f"X {args.file} not found"); return

    # find header row: 揾含「項目序號」或「Item Code」嗰行做 header
    raw = pd.read_excel(fp, sheet_name=args.sheet, header=None, dtype=object)
    hdr = None
    for i in range(min(8, len(raw))):
        row = " ".join(str(x) for x in raw.iloc[i].tolist())
        if ("項目序號" in row or "Item Code" in row) and ("Project Session" in row or "WD" in row or "Payroll" in row):
            hdr = i; break
    if hdr is None: hdr = 0
    df = pd.read_excel(fp, sheet_name=args.sheet, header=hdr, dtype=object)
    cols = list(df.columns)

    if args.inspect:
        print(f"header row = {hdr}; {len(cols)} 欄:")
        for i, c in enumerate(cols):
            sv = df[c].dropna().astype(str).head(3).tolist()
            print(f"  [{i}] {str(c)[:38]:38} e.g. {sv}")
        return

    C = {  # detect columns
        "序號": find_col(cols, "項目序號"),
        "名稱": find_col(cols, "項目名稱"),                 # Master 名（左）
        "名稱_s": find_col(cols, "Project Name"),           # Supporting 名（右）
        "Payroll_m": find_col(cols, "Payroll"),
        "CAPEX": find_col(cols, "CAPEX"),
        "OPEX": find_col(cols, "OPEX"),
        "Item": find_col(cols, "Item Code"),
        "Session": find_col(cols, "Project Session"),
        "InvType": find_col(cols, "Investment Type"),
        "WD1": find_col(cols, "WD1"),
        "WD2": find_col(cols, "WD2"),
        "WD3": find_col(cols, "WD3"),
        "WD4": find_col(cols, "WD4"),
        "Patron": find_col(cols, "Patron"),
    }
    print("偵測到嘅欄:", {k: v for k, v in C.items() if v is not None})

    def num(x):
        try: return float(str(x).replace(",", ""))
        except Exception: return 0.0
    def code(x):  # 「項目064-CAPEX」/「064」→ 064
        m = re.search(r"(\d{2,4})", str(x))
        return m.group(1).zfill(3) if m else ""

    # Master：序號 → (名稱, Payroll, CAPEX, OPEX)
    master = {}
    if C["序號"]:
        for _, r in df.iterrows():
            sn = code(r[C["序號"]])
            if sn and sn not in master:
                master[sn] = (str(r[C["名稱"]]) if C["名稱"] else "",
                              num(r[C["Payroll_m"]]) if C["Payroll_m"] else 0,
                              num(r[C["CAPEX"]]) if C["CAPEX"] else 0,
                              num(r[C["OPEX"]]) if C["OPEX"] else 0)

    # Supporting：每行 Item Code → NG + WD 拆 H
    out = []
    src = C["Item"] or C["Session"]
    for _, r in df.iterrows():
        if not src or pd.isna(r[src]) or str(r[src]).strip() == "":
            continue
        item = str(r[C["Item"]]) if C["Item"] else ""
        sn = code(item)
        ng, nglbl = ng_of(r[C["Session"]]) if C["Session"] else ("", "")
        wd1, wd2 = (num(r[C["WD1"]]) if C["WD1"] else 0), (num(r[C["WD2"]]) if C["WD2"] else 0)
        wd3, wd4 = (num(r[C["WD3"]]) if C["WD3"] else 0), (num(r[C["WD4"]]) if C["WD4"] else 0)
        pat = num(r[C["Patron"]]) if C["Patron"] else 0
        m = master.get(sn, ("", 0, 0, 0))
        out.append({
            "Item Code": item, "項目序號": sn,
            "Project Name": str(r[C["名稱_s"]]) if C["名稱_s"] else (str(r[C["名稱"]]) if C["名稱"] else m[0]),
            "NG": ng, "NG_label": nglbl,
            "Investment Type": str(r[C["InvType"]]) if C["InvType"] else "",
            "WD1→待拆": wd1, "WD2→人工": wd2, "WD3→人工": wd3, "WD4→餐飲": wd4, "Patron→Comp其他": pat,
            "Master_Payroll→人工": m[1], "Master_CAPEX→建設/設施": m[2], "Master_OPEX": m[3],
            "Master_Total(golden)": m[1] + m[2] + m[3],
        })
    res = pd.DataFrame(out)
    op_detail = ROOT / "results" / "mgm_mapping_detail.tsv"
    op_detail.parent.mkdir(parents=True, exist_ok=True)
    res.to_csv(op_detail, sep="\t", index=False)   # 逐 Item Code 明細（CAPEX/OPEX 分開）

    # ── 統一到 project（項目序號 + 項目名）：每個 project 一行，WD 加埋 ──
    proj = []
    for sn, sub in res.groupby("項目序號"):
        if not sn:
            continue
        names = [n for n in sub["Project Name"].astype(str).unique() if n and n != "nan"]
        mname = master.get(sn, ("",))[0]
        proj.append({
            "項目序號": sn,
            "項目名稱": mname or (names[0] if names else ""),
            "Item Codes": " / ".join(sub["Item Code"].astype(str).unique()),
            "NG": next((x for x in sub["NG"] if x), ""),
            "NG_label": next((x for x in sub["NG_label"] if x), ""),
            "Investment Type": " / ".join(x for x in sub["Investment Type"].astype(str).unique() if x and x != "nan"),
            "WD1→待拆": sub["WD1→待拆"].sum(), "WD2→人工": sub["WD2→人工"].sum(),
            "WD3→人工": sub["WD3→人工"].sum(), "WD4→餐飲": sub["WD4→餐飲"].sum(),
            "Patron→Comp其他": sub["Patron→Comp其他"].sum(),
            "Master_Payroll→人工": sub["Master_Payroll→人工"].iloc[0],
            "Master_CAPEX→建設/設施": sub["Master_CAPEX→建設/設施"].iloc[0],
            "Master_OPEX": sub["Master_OPEX"].iloc[0],
            "Master_Total(golden)": sub["Master_Total(golden)"].iloc[0],
            "名核對": "" if (not names or not mname or mname in names) else f"⚠ Master='{mname}' vs Supp='{names[0]}'",
        })
    seen = {p["項目序號"] for p in proj}
    for sn, (mname, pay, cap, opx) in master.items():   # 補返只有 Master、冇 Supporting 嘅項目（令 golden tie）
        if sn and sn not in seen:
            proj.append({"項目序號": sn, "項目名稱": mname, "Item Codes": "", "NG": "", "NG_label": "",
                         "Investment Type": "", "WD1→待拆": 0, "WD2→人工": 0, "WD3→人工": 0, "WD4→餐飲": 0,
                         "Patron→Comp其他": 0, "Master_Payroll→人工": pay, "Master_CAPEX→建設/設施": cap,
                         "Master_OPEX": opx, "Master_Total(golden)": pay + cap + opx,
                         "名核對": "(只有 Master，冇 Supporting WD)"})
    pj = pd.DataFrame(proj)
    op = ROOT / "results" / "mgm_mapping.tsv"
    pj.to_csv(op, sep="\t", index=False)
    n_warn = int((pj["名核對"] != "").sum()) if len(pj) else 0
    print(f"\n統一到 project: {len(pj)} 個項目 → {op}   (明細 {len(res)} 行 → {op_detail.name})")
    if n_warn:
        print(f"⚠ {n_warn} 個項目 Master/Supporting 名唔一致（睇『名核對』欄）")
    print("\nNG 分佈 (Master golden Σ 萬):")
    g = pj.groupby("NG")["Master_Total(golden)"].sum().sort_index()
    for ng, v in g.items():
        print(f"  {ng or '(未對應)':<8} {v:>14,.0f}")
    print(f"  {'總計':<8} {pj['Master_Total(golden)'].sum():>14,.0f}")


if __name__ == "__main__":
    main()
