#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_sources.py — 盤點【報告三個 input】到底有幾多可用欄，同埋我哋而家用咗幾多。

背景：清單 Database 有 ~184 欄，我哋 build_narrative 只 map 咗 10 個概念；
表2 有標準 33 欄（關注事項／建議調整金額／調整原因／跨司兩輪意見…），但 biao2.py 而家係
「盲抓」（見到 ≥30 字嘅 cell 就當 finding，唔知邊欄係咩）。呢個 script 出一份逐欄清單，
等我可以改成【按欄名 structured 抽】，把未用嘅補充資料寫入報告。

用法（Windows，kpi-main 底下）：
    python scripts\\report\\inspect_sources.py
    python scripts\\report\\inspect_sources.py --entity mgm --rows 2500

輸出：console（分批，每批 ≤--batch 行，方便 paste）+ 檔 sources_audit.txt（同一內容）。
每行 = TSV：來源 / 檔或sheet / 欄index / 欄名 / 有值行數 / 相異值數 / 我哋用咗未 / 樣本
「用咗未」：USED=已接入報告；—=未用（＝可以攞嚟寫嘢）。
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

QINGDAN_DIR = "data/投資項目清單"
BIAO2_DIR = "data/表2"
FEED = "tableau_combined_25.csv"

# 我哋而家真係接入咗報告嘅欄（關鍵字）→ 其餘全部係「未用」
USED_QINGDAN = ["承批公司項目序號", "項目名稱", "項目類型", "項目性質", "項目狀況",
                "實際實施地點", "計劃投資內容", "實際投資內容", "分析發現", "投資偏離",
                "管理層解釋", "變更原因", "期後調整内容", "期後調整內容", "調整事項備註",
                "調整備註", "跨司工作組的回", "跨司工作组的回", "跨司工作組回", "第二輪意見",
                "KPMG回覆", "KPMG回复", "預計投資金額"]
USED_FEED = ["entity", "dicj code", "project", "year_bucket", "報告年", "ng_scope", "ng_label",
             "ng_code", "vertical_label", "final_capex_opex", "調整一級",
             "調整前_萬", "調整_萬", "調整後_萬"]


def _clean(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


def _used(name, used_list):
    n = _clean(name)
    return "USED" if any(k in n for k in used_list) else "—"


def _col_stats(rows, ci, hdr, used_list, max_sample=2):
    vals, seen = [], set()
    n = 0
    for r in rows:
        if ci >= len(r) or r[ci] is None:
            continue
        s = _clean(r[ci])
        if not s:
            continue
        n += 1
        seen.add(s[:60])
        if len(vals) < max_sample and len(s) > 1:
            vals.append(s[:90])
    if n == 0:
        return None
    return [str(ci), _clean(hdr)[:60] or "(無標題)", str(n), str(len(seen)),
            _used(hdr, used_list), " ⁄ ".join(vals)]


def _find_header(rows, keys, maxscan=14):
    """揾表頭行：包含任一 key 嘅第一行。回 (row_idx, header_list)。"""
    for ri in range(min(maxscan, len(rows))):
        row = rows[ri] or []
        if any(any(k in _clean(v) for k in keys) for v in row):
            return ri, [_clean(v) for v in row]
    return None, []


def audit_qingdan(entity, maxrows, out):
    import openpyxl
    d = Path(QINGDAN_DIR)
    files = [p for p in sorted(d.rglob("*.xlsx"))
             if entity.lower() in p.name.lower() and not p.name.startswith("~$")] if d.exists() else []
    if not files:
        out.append(f"# 清單：揾唔到（{QINGDAN_DIR} / entity={entity}）")
        return
    p = files[0]
    out.append(f"# 清單檔：{p.name}")
    wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    ws = next((wb[s] for s in wb.sheetnames if s.lower().startswith("database")), wb[wb.sheetnames[0]])
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        rows.append(r)
        if i > maxrows:
            break
    hr, hdr = _find_header(rows, ["承批公司項目序號"])
    if hr is None:
        out.append("# ⚠ 清單揾唔到『承批公司項目序號』表頭"); return
    body = rows[hr + 1:]
    ncol = max((len(r) for r in rows if r), default=0)
    out.append(f"# 清單 sheet={ws.title} 表頭喺第{hr+1}行，{ncol} 欄，{len(body)} 行資料")
    for ci in range(ncol):
        st = _col_stats(body, ci, hdr[ci] if ci < len(hdr) else "", USED_QINGDAN)
        if st:
            out.append("清單\t" + ws.title + "\t" + "\t".join(st))


def audit_biao2(entity, maxrows, out):
    import biao2 as B2
    import inspect_biao2 as IB
    d = Path(BIAO2_DIR)
    if not d.exists():
        out.append(f"# 表2：揾唔到 {BIAO2_DIR}"); return
    files = [p for p in sorted(d.rglob("*.xls*"))
             if not p.name.startswith("~$") and B2._match_entity(p.name, entity.lower())
             and "提供附件" not in p.name]
    out.append(f"# 表2：{len(files)} 檔 match「{entity}」")
    for p in files:
        try:
            wb = IB.load_wb(p)
        except Exception as e:
            out.append(f"# ⚠ 開唔到 {p.name}: {e}"); continue
        for sn in wb.sheetnames:
            try:
                ws = wb[sn]
                rows = []
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    rows.append(r)
                    if i > maxrows:
                        break
                ncol = max((len(r) for r in rows if r), default=0)
                if ncol == 0:
                    continue
                # 表2 兩層表頭：揾「投資項目」/「序號」嗰行，同上一行（group 標題）夾埋做欄名
                hr, hdr = _find_header(rows, ["投資項目", "項目序號", "項目名稱"])
                if hr is None:
                    hr, hdr = 0, [_clean(v) for v in (rows[0] or [])]
                grp = [_clean(v) for v in (rows[hr - 1] or [])] if hr > 0 else []
                body = rows[hr + 1:]
                out.append(f"# 表2 {p.name}｜{sn}：表頭第{hr+1}行，{ncol} 欄，{len(body)} 行")
                for ci in range(ncol):
                    h = hdr[ci] if ci < len(hdr) else ""
                    g = grp[ci] if ci < len(grp) else ""
                    name = (f"{g}／{h}" if g and h and g != h else (h or g))
                    st = _col_stats(body, ci, name, [])      # 表2 全部當「未用」（而家係盲抓）
                    if st:
                        out.append(f"表2\t{p.name[:28]}｜{sn[:22]}\t" + "\t".join(st))
            except Exception as e:
                out.append(f"# ⚠ {p.name}｜{sn}: {e}")


def audit_feed(entity, out):
    try:
        import pandas as pd
    except ImportError:
        out.append("# feed：冇 pandas"); return
    p = Path(FEED)
    if not p.exists():
        out.append(f"# feed：揾唔到 {FEED}"); return
    df = pd.read_csv(p, low_memory=False, nrows=200000)
    if "entity" in df.columns:
        df = df[df["entity"].astype(str).str.lower() == entity.lower()]
    out.append(f"# feed {p.name}：{len(df)} 行（{entity}），{len(df.columns)} 欄")
    for ci, c in enumerate(df.columns):
        s = df[c]
        nn = int(s.notna().sum())
        if nn == 0:
            continue
        samp = " ⁄ ".join(_clean(x)[:60] for x in s.dropna().unique()[:2])
        out.append("\t".join(["feed", "-", str(ci), _clean(c)[:60], str(nn),
                              str(int(s.nunique(dropna=True))), _used(c, USED_FEED), samp]))


def main():
    args = sys.argv[1:]
    entity, maxrows, batch = "mgm", 2500, 900
    for flag, cast in (("--entity", str), ("--rows", int), ("--batch", int)):
        if flag in args:
            i = args.index(flag); v = cast(args[i + 1]); del args[i:i + 2]
            entity = v if flag == "--entity" else entity
            maxrows = v if flag == "--rows" else maxrows
            batch = v if flag == "--batch" else batch
    entity = entity.lower()
    out = [f"# ==== report sources audit｜entity={entity} ====",
           "# 欄位：來源\t檔／sheet\t欄index\t欄名\t有值行數\t相異值\t用咗未\t樣本",
           "# 『USED』＝已接入報告；『—』＝未用（可以攞嚟寫報告內容）"]
    audit_feed(entity, out)
    audit_qingdan(entity, maxrows, out)
    audit_biao2(entity, maxrows, out)

    f = Path(f"{entity}_sources_audit.txt")
    f.write_text("\n".join(out), encoding="utf-8")
    n_un = sum(1 for l in out if "\t—\t" in l)
    n_us = sum(1 for l in out if "\tUSED\t" in l)
    print(f"✓ {f.resolve()}（{len(out)} 行；USED {n_us}、未用 {n_un}）")
    print(f"── 以下分批印出，每批 ≤{batch} 行，逐批 paste 返俾我 ──")
    for i in range(0, len(out), batch):
        print(f"\n===== BATCH {i//batch + 1}/{-(-len(out)//batch)} =====")
        print("\n".join(out[i:i + batch]))


if __name__ == "__main__":
    main()
