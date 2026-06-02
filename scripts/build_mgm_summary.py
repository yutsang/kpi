"""MGM 投資方向 — SUMMARY-DRIVEN build (from the project-team supporting detail).

This REPLACES the simplified-golden approach for MGM. The supporting detail is row-level per 項目
and already carries everything we were guessing before:

  Item Code | Project Name | Project Session | Investment Type |
  WD1 | WD2-Allocation | WD3-Payroll | WD4-COGS | Patron Management | Total   (萬元 → ×10000)

  - Project Session  = the NG bucket DIRECTLY (B1..B11 / 名稱-設施活動)  → NG authoritative, no guess
  - bucket           = CAPEX (Item Code -CAPEX / Investment Type 設施) vs OPEX (-OPEX / 活動)
  - per-WD split     = a ready horizontal-ish breakdown:
       WD3-Payroll        → H_LABOR (人工)        [direct]
       WD4-COGS           → H_FNB   (餐飲)        [direct]
       Patron Management  → H_COMP_OTHER (comp)   [confirm]
       WD1 / WD2-Allocation → 待拆 (項目組 WD-split rules pending)
  Totals tie to the summary's Total by construction.

Input: a TAB-separated export of the table (save the pasted block, header row included), e.g.
  results/mgm_summary_25.tsv      (2025 plan)
  results/mgm_summary_2324.tsv    (2023-2024 plan)

Run on Windows:
  python scripts/build_mgm_summary.py --file results/mgm_summary_25.tsv --year 25
  python scripts/build_mgm_summary.py --file results/mgm_summary_2324.tsv --year 24
"""
from __future__ import annotations
import argparse, csv, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent

NG_LABELS = {"NG0": "博彩項目", "NG1": "吸引外國客源", "NG2": "會議展覽", "NG3": "娛樂表演",
             "NG4": "體育盛事", "NG5": "文化藝術", "NG6": "健康養生", "NG7": "主題遊樂",
             "NG8": "美食之都", "NG9": "社區旅遊", "NG10": "海上旅遊", "NG11": "其他"}
NG_NAME2CODE = {"吸引外國客源": "NG1", "會議展覽": "NG2", "娛樂表演": "NG3", "體育盛事": "NG4",
                "文化藝術": "NG5", "健康養生": "NG6", "主題遊樂": "NG7", "美食之都": "NG8",
                "社區旅遊": "NG9", "海上旅遊": "NG10", "其他": "NG11"}

# WD column → our horizontal id. WD1/WD2 stay 待拆 (項目組 split rules pending).
WD_TO_H = {
    "wd3": ("H_LABOR", "人工成本"),
    "wd4": ("H_FNB", "餐飲"),
    "patron": ("H_COMP_OTHER", "Comp其他"),
    "wd1": ("(待拆)WD1", "待拆 (WD1)"),
    "wd2": ("(待拆)WD2", "待拆 (WD2-分配)"),
}


def _num(x):
    s = str(x).strip().replace(",", "").replace("，", "")
    if s in ("", "-", "–", "—", "－", "nan", "None"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s) * 10000.0           # 萬元
    except Exception:
        return 0.0
    return -v if neg else v


def _ng_from_session(sess, inv_type):
    s = str(sess)
    m = re.search(r"B\s*(\d+)", s)               # 2025: "B1吸引外國客源"
    if m:
        return f"NG{int(m.group(1))}"
    head = re.split(r"[-－—–]", s)[0].strip()      # 2023-24: "文化藝術 - 設施"
    for nm, code in NG_NAME2CODE.items():
        if nm in head or nm in s:
            return code
    return "NG11"


def _bucket(item_code, inv_type):
    ic = str(item_code).upper()
    if "CAPEX" in ic:
        return "Capex"
    if "OPEX" in ic:
        return "Opex"
    it = str(inv_type)
    if it.startswith("設施"):
        return "Capex"
    if it.startswith("活動"):
        return "Opex"
    return "Opex"


def _find_col(header, *keys):
    for i, h in enumerate(header):
        hl = str(h).strip().lower()
        if any(k.lower() in hl for k in keys):
            return i
    return None


def _is_header(row):
    joined = " ".join(str(c).strip().lower() for c in row)
    has_id = ("project session" in joined) or ("item code" in joined)
    has_wd = ("wd1" in joined) or ("payroll" in joined) or ("patron" in joined) or ("wd2" in joined)
    return has_id and has_wd


def _rows_from_xlsx(path):
    """(rows, hidx, 'file::sheet') for the first sheet whose header matches, else (None,None,None).
    Peeks 25 rows per sheet to find the header, then reads that sheet fully."""
    try:
        peek = pd.read_excel(path, sheet_name=None, header=None, dtype=str, nrows=25)
    except Exception as e:
        print(f"  ! peek {path.name} failed: {e}"); return None, None, None
    for sname, sdf in peek.items():
        for i, r in enumerate(sdf.fillna("").astype(str).values.tolist()):
            if _is_header(r):
                full = pd.read_excel(path, sheet_name=sname, header=None, dtype=str)
                return full.fillna("").astype(str).values.tolist(), i, f"{path.name}::{sname}"
    return None, None, None


def _load_rows(args):
    """Return (rows, header_index, source_label). Prefer --file; else auto-scan conf raw files."""
    if args.file:
        fp = ROOT / args.file
        if fp.exists():
            if fp.suffix.lower() in (".xlsx", ".xls"):
                rows, hidx, src = _rows_from_xlsx(fp)
                if rows is not None:
                    return rows, hidx, src
                print(f"X {fp.name} 冇 sheet 啱 header")
                return None, None, None
            rows = list(csv.reader(fp.read_text(encoding="utf-8-sig").splitlines(), delimiter="\t"))
            hidx = next((i for i, r in enumerate(rows) if _is_header(r)), None)
            return rows, hidx, fp.name
        print(f"… {fp} 唔存在 → 改為自動掃 conf raw 檔")
    # auto-scan conf prebuild_sources (the files conf already names) + any other xlsx in raw dir
    raw_dir = ROOT / args.raw
    cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
    srcs = cfg.get("prebuild_sources") or {}
    cand = [raw_dir / v for v in srcs.values() if v]
    cand += [p for p in sorted(raw_dir.glob("*.xlsx")) if p not in cand]
    print(f"auto-scan {len([p for p in cand if p.exists()])} 個 raw 檔 (Project Session + WD/Patron)…")
    for p in cand:
        if not p.exists():
            continue
        rows, hidx, src = _rows_from_xlsx(p)
        if rows is not None:
            print(f"  ✓ 搵到 summary sheet: {src}")
            return rows, hidx, src
    print("X 自動掃唔到 summary sheet。以下係各檔 sheet + 第一行,畀我睇結構再 wire:")
    for p in cand:
        if not p.exists():
            continue
        try:
            for sname, sdf in pd.read_excel(p, sheet_name=None, header=None, dtype=str, nrows=2).items():
                first = sdf.fillna("").astype(str).values.tolist()
                head = " | ".join(first[0][:12]) if first else ""
                print(f"   {p.name}::{sname} → {head[:120]}")
        except Exception as e:
            print(f"   {p.name}: read fail {e}")
    return None, None, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default=None, help="override; else auto-scan conf raw files")
    ap.add_argument("--raw", default="data/mgm/raw")
    ap.add_argument("--golden", default="results/mgm_golden_25.tsv", help="capex source (f1 col L per 項目)")
    ap.add_argument("--year", default="25")
    args = ap.parse_args()

    rows_raw, hidx, source = _load_rows(args)
    if rows_raw is None or hidx is None:
        print("X 攞唔到 summary 表 (header 要有 'Item Code'/'Project Session' + WD/Patron)。")
        return
    print(f"source: {source}  (header @ row {hidx})")
    header = rows_raw[hidx]
    ci = {
        "item": _find_col(header, "item code", "item"),
        "name": _find_col(header, "project name", "name"),
        "sess": _find_col(header, "project session", "session"),
        "inv": _find_col(header, "investment type", "investment"),
        "wd1": _find_col(header, "wd1"),
        "wd2": _find_col(header, "wd2", "allocation"),
        "wd3": _find_col(header, "wd3", "payroll"),
        "wd4": _find_col(header, "wd4", "cogs"),
        "patron": _find_col(header, "patron"),
        "total": _find_col(header, "total"),
    }
    print(f"header @ line {hidx}: " + "  ".join(f"{k}={ci[k]}" for k in ci))

    cats = yaml.safe_load((ROOT / "conf/base/categories.yml").read_text(encoding="utf-8"))
    hlab = {h["id"]: h.get("label", h["id"]) for h in cats.get("horizontals", [])}
    h_order = [h["label"] for h in cats.get("horizontals", []) if h.get("id") != "H_COUNT" and h.get("label")]

    def _cell(r, key):
        i = ci[key]
        return r[i] if i is not None and i < len(r) else ""

    # The sheet stacks the 2025 plan then the 2023-2024 plan (each with a title row). Split by
    # that boundary so --year 25 takes ONLY the 2025 block (was mixing both → 不consol).
    detail, recon = [], []
    seen = 0
    section = "25"                                    # rows after the first header = 2025 block
    for r in rows_raw[hidx + 1:]:
        joined = " ".join(str(c) for c in r)
        if "2023" in joined and "2024" in joined:     # "MGM 2023-2024年度投資計劃" title → 2nd block
            section = "24"
            continue
        if _is_header(r):                             # the 2nd block's own header row
            continue
        item = str(_cell(r, "item")).strip()
        sess = str(_cell(r, "sess")).strip()
        if not item and not sess:
            continue
        if not re.search(r"項目\s*\d+", item + sess) and "項目" not in item:
            continue                                  # skip title / blank / total lines
        if section != args.year:                      # keep only the requested year's block
            continue
        name = str(_cell(r, "name")).strip()
        ng = _ng_from_session(sess, _cell(r, "inv"))
        bucket = _bucket(item, _cell(r, "inv"))
        wd = {k: _num(_cell(r, k)) for k in ("wd1", "wd2", "wd3", "wd4", "patron")}
        total = _num(_cell(r, "total"))
        seen += 1
        for k, amt in wd.items():
            if amt == 0:
                continue
            hid, hl = WD_TO_H[k]
            detail.append([item, name, ng, NG_LABELS.get(ng, "其他"), bucket,
                           hid, hlab.get(hid, hl), round(amt)])
        recon.append([item, name[:34], ng, bucket,
                      round(wd["wd1"]), round(wd["wd2"]), round(wd["wd3"]),
                      round(wd["wd4"]), round(wd["patron"]), round(total),
                      round(sum(wd.values())), round(sum(wd.values()) - total)])

    det = pd.DataFrame(detail, columns=["Item", "項目", "ng_code", "ng_label", "bucket",
                                        "horizontal_id", "horizontal_label", "amount_mop"])
    rec = pd.DataFrame(recon, columns=["Item", "項目", "ng_code", "bucket",
                                       "WD1", "WD2分配", "WD3人工", "WD4餐飲", "Patron",
                                       "Total", "ΣWD", "ΔΣWD-Total"])
    print(f"parsed {seen} 項目 rows ({args.year} block); detail rows={len(det):,}")

    # ── Add CAPEX from f1 (= golden capex column, which is f1 '2025 Raw Data' Debit-minus-Credit
    #    per 項目). The summary's WD columns are the OPERATING allocation (payroll/comp/COGS/…);
    #    the CAPITAL construction spend (~1,097M) is ADDITIVE on top, taken per 項目 序號, NG from
    #    the summary's Project Session, H = (待拆)建設capex (建設/設備/租賃/維護 split rules pending).
    #    ASSUMPTION (correct me): additive, and WD3 payroll already counted so golden payroll NOT
    #    re-added (the ~82M payroll tail is the 尾數 you said to ignore).
    cap_added, cap_sum = 0, 0.0
    base2ng = {}
    for _, rr in rec.iterrows():
        m = re.search(r"項目\s*0*(\d+)", str(rr["Item"]))
        if m:
            base2ng.setdefault(m.group(1).zfill(3), (rr["ng_code"], str(rr["項目"])))
    gp = ROOT / args.golden
    if args.year == "25" and gp.exists():
        for line in gp.read_text(encoding="utf-8-sig").splitlines():
            p = [c.strip() for c in line.split("\t")]
            if len(p) < 6 or not re.fullmatch(r"\d+", p[0]):
                continue
            cap = _num(p[3])                       # golden capex column (萬元 → MOP)
            if cap == 0:
                continue
            sn = p[0].zfill(3)
            ng, nm = base2ng.get(sn, ("NG11", p[1]))
            detail.append([f"項目{sn}-CAPEX(f1)", nm, ng, NG_LABELS.get(ng, "其他"),
                           "Capex", "(待拆)建設capex", "待拆 建設capex(f1)", round(cap)])
            cap_added += 1
            cap_sum += cap
        det = pd.DataFrame(detail, columns=det.columns)
        print(f"  + f1 CAPEX (建設, additive): {cap_added} 項目  Σ={cap_sum:,.0f}")
    elif args.year == "25":
        print(f"  ! {gp.name} 唔喺度 → CAPEX(~1,097M) 未加,output 只有 summary OPEX side。")

    # Consolidate by base 項目 number (roll up the -CAPEX / -OPEX lines of the SAME project into
    # one row → fixes the 'project 名字唔 consol' = same 項目 appearing as multiple rows).
    det["base"] = det["Item"].astype(str).str.extract(r"項目\s*0*(\d+)")[0]
    det["base"] = det["base"].fillna(det["Item"].astype(str))
    _name1 = det.groupby("base")["項目"].first()
    consol = det.pivot_table(index=["base", "ng_code", "ng_label"], columns="bucket",
                             values="amount_mop", aggfunc="sum", fill_value=0).reset_index()
    _bk = [c for c in consol.columns if c not in ("base", "ng_code", "ng_label")]
    consol["Total"] = consol[_bk].sum(axis=1)
    consol.insert(1, "項目", consol["base"].map(_name1))
    consol = consol.sort_values("Total", ascending=False)

    # NG × H pivot (待拆 columns appended after the taxonomy order)
    pv = det.pivot_table(index=["ng_code", "ng_label"], columns="horizontal_label",
                         values="amount_mop", aggfunc="sum", fill_value=0,
                         margins=True, margins_name="總計")
    extra = [c for c in pv.columns if c not in h_order and c != "總計"]
    pv = pv.reindex(columns=h_order + extra + (["總計"] if "總計" in pv.columns else []), fill_value=0)

    h_by = det.groupby(["bucket", "horizontal_label"])["amount_mop"].sum().reset_index() \
              .sort_values(["bucket", "amount_mop"], ascending=[True, False])
    ng_by = det.groupby(["ng_code", "ng_label", "bucket"])["amount_mop"].sum().reset_index() \
               .sort_values(["ng_code", "bucket"])

    tot_alloc = det["amount_mop"].sum()
    tot_summary = rec["Total"].sum()
    summ = pd.DataFrame({
        "metric": ["summary側(WD,opex+小capex)", "+f1_CAPEX(建設,待拆)", "TOTAL",
                   "golden_total(對標)", "Δ_vs_golden",
                   "—", "WD1", "WD2分配", "WD3人工", "WD4餐飲", "Patron",
                   "bucket_capex", "bucket_opex"],
        "amount": [round(tot_summary), round(cap_sum), round(tot_alloc),
                   2086090000, round(tot_alloc - 2086090000),
                   0, round(rec["WD1"].sum()), round(rec["WD2分配"].sum()), round(rec["WD3人工"].sum()),
                   round(rec["WD4餐飲"].sum()), round(rec["Patron"].sum()),
                   round(det.loc[det.bucket.eq("Capex"), "amount_mop"].sum()),
                   round(det.loc[det.bucket.eq("Opex"), "amount_mop"].sum())],
    })

    out_dir = ROOT / "data/review"; out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"mgm_投資方向_{args.year}_summary.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        pd.DataFrame([
            ("0_index", "this map"),
            ("1_NG×H_pivot", "NG × H cross-tab (WD3→人工, WD4→餐飲, Patron→comp, WD1/WD2→待拆) + 總計"),
            ("2_by_項目_consol", "CONSOLIDATED per 項目 (capex+opex 合一行) — NG / bucket totals / Total"),
            ("2b_by_Item", "per Item-Code (raw -CAPEX/-OPEX lines): NG / WD1..Patron / Total + ΣWD tie"),
            ("3_橫向", "(bucket, H) × Σ amount"),
            ("4_縱向", "(NG, bucket) × Σ amount"),
            ("5_大表", "row-level: (Item, 項目, NG, bucket, H) × amount"),
            ("6_tie", "ΣWD vs summary Total — 注意: 只係 OPEX side, CAPEX(~1,097M) 喺 f1 未加"),
        ], columns=["sheet", "contents"]).to_excel(w, sheet_name="0_index", index=False)
        pv.to_excel(w, sheet_name="1_NG×H_pivot")
        consol.to_excel(w, sheet_name="2_by_項目_consol", index=False)
        rec.to_excel(w, sheet_name="2b_by_Item", index=False)
        h_by.to_excel(w, sheet_name="3_橫向", index=False)
        ng_by.to_excel(w, sheet_name="4_縱向", index=False)
        det.to_excel(w, sheet_name="5_大表", index=False)
        summ.to_excel(w, sheet_name="6_tie", index=False)

    print(f"\n=== tie vs golden 2,086,090,000 ===")
    print(f"  summary側(WD)={tot_summary:,.0f}  + f1 CAPEX={cap_sum:,.0f}  = TOTAL {tot_alloc:,.0f}")
    print(f"  Δ_vs_golden={tot_alloc-2086090000:,.0f}  (~82M payroll 尾數先 ignore)")
    print(f"  WD3人工={rec['WD3人工'].sum():,.0f}  WD4餐飲={rec['WD4餐飲'].sum():,.0f}  "
          f"Patron={rec['Patron'].sum():,.0f}  待拆(WD1+WD2)={rec['WD1'].sum()+rec['WD2分配'].sum():,.0f}")
    print("\n  NG 分佈 (Σamount):")
    for _, r in ng_by.groupby(["ng_code", "ng_label"])["amount_mop"].sum().reset_index() \
            .sort_values("ng_code").iterrows():
        print(f"    {r['ng_code']:5} {r['ng_label']:8} {r['amount_mop']:>15,.0f}")
    print(f"\n→ {out}")


if __name__ == "__main__":
    main()
