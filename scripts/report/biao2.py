#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
biao2.py — 由「表2」(審查底稿，加密 （密碼由環境變數 KPI_XLSX_PW 提供）) 逐項目抽 finding，做第 2 個 narrative source
（配清單一齊餵 LLM）。表2＝最權威嘅調整/發現來源。

★ structured 抽（2026-08-12 起）：表2 有標準兩層表頭 —— group 行 +【下面嗰行】先係 detail
  表頭（概念名），欄序會浮動 36-38 欄，『項目編號』喺 group 行。所以一律【按欄名認】，
  唔用固定 index。實測 mgm 7 檔 × 17 sheet 全部 detail 表頭 r6、code 欄 col35。
  （舊版 load_biao2 係盲抓：見到 ≥30 字 cell 就當 finding，冇 label，已 deprecated。）

key = (gaming, 正規化碼)。gaming 由檔名判：『博監局』檔＝博彩，其餘＝非博彩。
→ 修返「博彩娛樂場」冇內容（清單博彩碼撞號攞唔到，但表2 博監局有）。

用法：
    python scripts\\report\\biao2.py data\\表2 --entity mgm      # 逐概念覆蓋率 + 逐項目內容
    from biao2 import load_biao2_struct, b2rec, b2text
    b2 = load_biao2_struct("data/表2", "mgm")
    b2rec(b2, "gaming", "項目19")   → {"關注事項": …, "建議調整金額": …}
    b2text(b2, "gaming", "項目19")  → 有 label 嘅文字（餵 LLM）
"""
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import inspect_biao2 as IB      # load_wb（msoffcrypto 解密）

# 乾淨 code：項目N / 字母碼(B11.1/OP005/IV008)；唔匹配金額(27206)避免揀錯欄
_CODE_RE = re.compile(r"^(項目\s*\d+|[A-Za-z]{1,5}\d+(?:\.\d+)?)$")
_JUNK_RE = re.compile(r"^(無新增問題|無|是|否|已回覆|未回覆|不適用|n/?a|請參閱附件)")

# 表2 檔名可能用中文行名（非 mgm/galaxy）→ 別名匹配
_ENT_ALIASES = {
    "mgm": ["mgm", "美高梅"],
    "galaxy": ["galaxy", "銀河"],
    "sjm": ["sjm", "澳博", "新葡京"],
    "wynn": ["wynn", "永利"],
    "vml": ["vml", "威尼斯", "金沙"],
    "melco": ["melco", "新濠"],
}


def _match_entity(fname, entity):
    fl = fname.lower()
    for a in _ENT_ALIASES.get(entity, [entity]):
        if a.lower() in fl:
            return True
    return False


def _norm(v):
    s = re.sub(r"\s+", "", str(v if v is not None else ""))
    s = re.sub(r"^項目", "", s)
    m = re.match(r"^0*(\d+)$", s)
    return m.group(1) if m else s.lower()


def _txt(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


# ── 表2 標準欄（實測 mgm 7 檔 × 14 sheet，2026-08-12 sources audit）────────────
# detail 表頭行喺 group 行【下面】嗰行；欄位順序會浮動（36-38 欄）→ 一律按欄名認，唔用固定 index。
B2_FIELDS = {
    "關注事項": ["畢馬威關注事項"],                 # ★ 調整嘅正文
    "調整類型": ["需溝通關注事項"],                 # e.g.「一般支持性部門的人工成本」
    "關注事項金額": ["該關注事項涉及調整金額"],
    "建議調整金額": ["建議調整金額"],
    "調整原因": ["調整原因"],
    "建議調整後金額": ["建議調整後金額"],
    "KPMG分析": ["KPMG分析"],                     # 第一/二輪諮詢各一，兩個都收
    "管理層解釋": ["承批公司管理層解釋"],
    "承批公司反饋": ["承批公司的反饋意見"],
    "跨司回覆": ["跨司工作組的回覆", "跨司工作組最新反饋意見", "跨司工作組的反饋意見",
                 "跨司工作組主責部門"],
    # 跨司工作組審閱意見 block：實測 mgm 成欄係空（跨司未填，標題仲係 2026.07.XX）→ 抽到 0 係正常
    "項目分析意見": ["項目分析意見"],
    "建議接納調整後金額": ["建議接納之調整後金額"],
    # ⚠ 唔好加「擬投資金額／已投放金額」：嗰啲係【該性質範疇各項目之加總金額】，唔掛喺項目碼上，
    #   forward-fill 落去會亂咁派。per-project 投資內容一律由清單『實際投資內容』攞。
}
# 出 prompt 時嘅次序（最有用擺前）
B2_ORDER = ["調整類型", "關注事項", "調整原因", "關注事項金額", "建議調整金額",
            "建議調整後金額", "KPMG分析", "管理層解釋", "承批公司反饋", "跨司回覆",
            "項目分析意見", "建議接納調整後金額"]
# 金額類：多值要用「／」分開，唔可以黐埋（27832　26404 睇落似一個數）
_NUMISH = {"關注事項金額", "建議調整金額", "建議調整後金額", "建議接納調整後金額"}
_HDR_HINT = [k for ks in B2_FIELDS.values() for k in ks] + ["項目編號", "資料要求", "問題狀態"]


def _detail_header_row(rows, band=10):
    """detail 表頭行 = 頭 band 行入面 match 到最多概念嗰行（group 行喺佢上面）。"""
    best, bn = 0, 0
    for ri in range(min(band, len(rows))):
        n = sum(1 for v in (rows[ri] or []) if any(k in _txt(v) for k in _HDR_HINT))
        if n > bn:
            bn, best = n, ri
    return best, bn


def _field_cols(hdr):
    """detail 表頭 → {概念: [col_idx]}。"""
    cols = {}
    for ci, v in enumerate(hdr):
        s = _txt(v)
        if not s:
            continue
        for concept, keys in B2_FIELDS.items():
            if any(k in s for k in keys):
                cols.setdefault(concept, []).append(ci)
    return cols


def load_biao2_struct(folder, entity, log=lambda *a: None):
    """{(gaming, 正規化碼): {概念: 文字}} —— 按【欄名】structured 抽（唔再盲抓）。
    表2 layout：group 行 + detail 表頭行（概念名）+ 逐項目 data 行；欄序浮動 → 全部按名認。
    同一個碼有多行（matrix）→ 每個概念收最長嗰個非空值，其餘唔同值就接落去。"""
    out = {}
    d = Path(folder)
    if not d.exists():
        log(f"（冇 {folder}）"); return out
    allx = [p for p in sorted(d.rglob("*.xls*")) if not p.name.startswith("~$")]
    files = [p for p in allx if _match_entity(p.name, entity.lower()) and "提供附件" not in p.name]
    log(f"表2 folder {folder}：共 {len(allx)} 個 xls*，match「{entity}」{len(files)} 檔")
    n_field = 0
    for p in files:
        gaming = ("博監局" in p.name)
        try:
            wb = IB.load_wb(p)
        except Exception as e:
            log(f"  ⚠ 開唔到 {p.name}: {e}"); continue
        for sn in wb.sheetnames:
            try:
                rows = []
                for i, r in enumerate(wb[sn].iter_rows(values_only=True)):
                    rows.append(r)
                    if i > 700:
                        break
                ncol = max((len(r) for r in rows if r), default=0)
                if ncol == 0:
                    continue
                hr, nmatch = _detail_header_row(rows)
                if nmatch < 3:
                    continue                     # 唔似標準表2 sheet（附件/圖片頁）→ 跳
                fcols = _field_cols(rows[hr])
                if not fcols:
                    continue
                # code 欄：優先『項目編號』表頭（實測佢喺 group 行，唔喺 detail 行）→ 兩行都揾；
                # 都揾唔到就用 _CODE_RE 命中最多嗰欄
                hdr_band = list(rows[hr]) + list(rows[hr - 1] if hr else [])
                code_c = next((ci % max(ncol, 1) for ci, v in enumerate(hdr_band)
                               if "項目編號" in _txt(v)), None)
                if code_c is None:
                    best, bestn = None, 0
                    for ci in range(ncol):
                        n = sum(1 for r in rows if ci < len(r) and r[ci] is not None
                                and _CODE_RE.match(re.sub(r"\s+", "", str(r[ci]))))
                        if n > bestn:
                            bestn, best = n, ci
                    code_c = best
                if code_c is None:
                    continue
                log(f"  · {p.name}｜{sn}：detail表頭 r{hr+1}、code欄 col{code_c}、"
                    f"{len(fcols)} 個概念（{'博彩' if gaming else '非博彩'}）")
                for r in rows[hr + 1:]:
                    if code_c >= len(r) or r[code_c] is None:
                        continue
                    if not _CODE_RE.match(re.sub(r"\s+", "", str(r[code_c]))):
                        continue
                    rec = out.setdefault((gaming, _norm(r[code_c])), {})
                    for concept, cis in fcols.items():
                        for ci in cis:
                            if ci >= len(r):
                                continue
                            s = _txt(r[ci])
                            if len(s) < 3 or _JUNK_RE.match(s):
                                continue
                            cur = rec.get(concept, "")
                            if s in cur:
                                continue
                            sep = "／" if concept in _NUMISH else "　"
                            rec[concept] = (cur + sep + s).strip(sep) if cur else s
                            n_field += 1
            except Exception as e:
                log(f"  ⚠ {p.name}｜{sn}: {e}")
    n_all = len(out)
    out = {k: v for k, v in out.items() if v}      # 冇任何欄值 = 該項目冇 finding，唔留空 key
    log(f"表2（structured）：{len(files)} 檔 → {len(out)} 個項目有內容"
        f"（另 {n_all - len(out)} 個碼冇 finding）、{n_field} 個欄值")
    return out


# ── 調整類型層：主要發現嗰章嘅「事項描述」原文 ──────────────────────────
# 2026-09-14 真係讀通表2 之後確認嘅嘢：報告『主要發現』每個調整類型嗰段事項描述，
# 原文就喺表2『畢馬威關注事項』欄，八個類型一個唔少，段落形狀同報告一樣
# （事實 → 引指引條文 → 建議剔除）。同一段會逐個項目抄一次，所以喺幾十行、
# 跨檔逐字重複 →【重複最多嗰段】＝類型層 boilerplate，只出現一兩次嗰啲係該項目
# 獨有補充。我哋本來只攞 per-project 碎片餵 LLM 叫佢自己歸納，原文喺度都冇攞。
#
# 另一個要點：表2 好多行嘅項目碼欄係合併儲存格（空），load_biao2_struct 靠碼認行
# 會成批漏 → 呢度唔靠碼，逐行收。
#
# 一格可能寫兩個類型：「1、其他日常營運支出調整：… 2、未完全實現投資目的…：…」
# → 先切編號（只切後面真係跟住「短標題＋冒號」嗰啲，正文入面嘅「1）」唔會中），
#   再由開頭「名（金額）：」認類型。類型名喺呢度【唔】canonical 化 —— B.CANON 喺
#   報告嗰層，biao2.py 保持唔依賴報告。
_ADJ_NUM = r"\d{1,2}[、．]"
_ADJ_HEAD = re.compile(r"^(?:" + _ADJ_NUM + r")?\s*([^：:]{4,40}?)\s*[：:]\s*(.+)$", re.S)
_ADJ_AMT = re.compile(r"[（(][^）)]*(?:萬|億)[^）)]*[）)]\s*$")     # 類型名後面嗰個「（N萬澳門元）」
_ADJ_EXTRA = ["承批公司反饋", "跨司回覆", "KPMG分析", "管理層解釋", "調整原因"]


def split_adj_cell(cell):
    """一格『畢馬威關注事項』→ [(類型名, 段落)]。"""
    out = []
    for p in re.split(r"(?:^|\s)(?=" + _ADJ_NUM + r"[^：:]{4,40}[：:])", _txt(cell)):
        m = _ADJ_HEAD.match(p.strip())
        if not m:
            continue
        name = _ADJ_AMT.sub("", m.group(1)).strip()
        body = m.group(2).strip()
        if len(name) >= 4 and len(body) >= 20:
            out.append((name, body))
    return out


def _iter_rows(folder, entity, log=lambda *a: None):
    """逐檔逐 sheet 逐行 → (檔名, sheet, {概念: [文字]})，唔理項目碼。
    by_adj 同 corpus 共用 —— 兩者都唔需要逐項目對號，需要嘅係全部敘述原文。"""
    d = Path(folder)
    if not d.exists():
        log(f"（冇 {folder}）"); return
    files = [p for p in sorted(d.rglob("*.xls*"))
             if not p.name.startswith("~$") and _match_entity(p.name, entity.lower())
             and "提供附件" not in p.name]
    log(f"表2 folder {folder}：match「{entity}」{len(files)} 檔")
    for p in files:
        try:
            wb = IB.load_wb(p)
        except Exception as e:
            log(f"  ⚠ 開唔到 {p.name}: {e}"); continue
        for sn in wb.sheetnames:
            try:
                rows = []
                for i, r in enumerate(wb[sn].iter_rows(values_only=True)):
                    rows.append(r)
                    if i > 700:
                        break
                if not rows:
                    continue
                hr, nmatch = _detail_header_row(rows)
                if nmatch < 3:
                    continue                   # 附件／圖片頁，唔似標準表2
                fcols = _field_cols(rows[hr])
                if not fcols:
                    continue
                for r in rows[hr + 1:]:
                    rec = {}
                    for concept, cis in fcols.items():
                        for ci in cis:
                            if ci >= len(r):
                                continue
                            s = _txt(r[ci])
                            if len(s) >= 3 and not _JUNK_RE.match(s):
                                rec.setdefault(concept, []).append(s)
                    if rec:
                        yield p.name, sn, rec
            except Exception as e:
                log(f"  ⚠ {p.name}｜{sn}: {e}")


def load_biao2_by_adj(folder, entity, log=lambda *a: None):
    """{類型名: {"事項描述": 原文, "n": 重複次數, "變體": [...], "承批公司反饋": [...], …}}

    「事項描述」＝該類型重複最多嗰段（同分就攞最長）。變體＝其餘寫法（多數係
    針對單一項目嗰版，render 項目層 bullet 時有用）。"""
    desc, extra, n_row = {}, {}, 0
    for _fn, _sn, rec in _iter_rows(folder, entity, log):
        n_row += 1
        names = []
        for cell in rec.get("關注事項", []):
            for nm, body in split_adj_cell(cell):
                desc.setdefault(nm, Counter())[body] += 1
                names.append(nm)
        for t in rec.get("調整類型", []):          # 『需溝通關注事項』＝該行嘅類型標籤
            names.append(_ADJ_AMT.sub("", t).strip())
        for nm in set(names):
            for k in _ADJ_EXTRA:
                for v in rec.get(k, []):
                    extra.setdefault(nm, {}).setdefault(k, Counter())[v] += 1
    out = {}
    for nm, c in desc.items():
        body, n = max(c.items(), key=lambda kv: (kv[1], len(kv[0])))
        out[nm] = {"事項描述": body, "n": n,
                   "變體": [t for t, _ in c.most_common() if t != body]}
        for k, v in (extra.get(nm) or {}).items():
            out[nm][k] = [t for t, _ in v.most_common(4)]
    log(f"表2（類型層）：{n_row} 行 → {len(out)} 個調整類型有『事項描述』原文")
    return out


def load_biao2_corpus(folder, entity, log=lambda *a: None):
    """表2 全部敘述欄原文（去重）。量『報告嗰段文字係咪表2 有原料』用。"""
    seen, out = set(), []
    for _fn, _sn, rec in _iter_rows(folder, entity, log):
        for vs in rec.values():
            for v in vs:
                if v not in seen:
                    seen.add(v); out.append(v)
    log(f"表2（corpus）：{len(out)} 段、{sum(len(x) for x in out):,} 字")
    return out


# 藝術品清單（報告 slide 101「藝術品展出情況清單」）：喺表2 附件 sheet，唔跟標準 33 欄
ART_FIELDS = ["購入時間", "Artwork", "名稱", "類別", "產權歸屬", "Artist",
              "購入", "添置原因", "展出紀錄", "當前位置", "當前狀態", "未來處置"]


def load_artwork(folder, entity, log=lambda *a: None):
    """表2 附件『藝術品』sheet → (欄名 list, 行 list)。冇就回 (None, [])。"""
    d = Path(folder)
    if not d.exists():
        return None, []
    for p in sorted(d.rglob("*.xls*")):
        if p.name.startswith("~$") or not _match_entity(p.name, entity.lower()):
            continue
        try:
            wb = IB.load_wb(p)
        except Exception:
            continue
        for sn in wb.sheetnames:
            if "藝術品" not in sn and "艺术品" not in sn:
                continue
            rows = [r for i, r in enumerate(wb[sn].iter_rows(values_only=True)) if i < 400]
            hr = next((i for i, r in enumerate(rows[:8])
                       if sum(1 for v in (r or []) if any(k in _txt(v) for k in ART_FIELDS)) >= 4), None)
            if hr is None:
                continue
            hdr = [_txt(v) for v in rows[hr]]
            keep = [i for i, h in enumerate(hdr) if h]
            body = []
            for r in rows[hr + 1:]:
                vals = [_txt(r[i]) if i < len(r) else "" for i in keep]
                if sum(1 for v in vals if v) >= 3:
                    body.append(vals)
            log(f"  · 藝術品清單：{p.name}｜{sn} → {len(body)} 件 × {len(keep)} 欄")
            return [hdr[i] for i in keep], body
    return None, []


def b2rec(b2s, ng_scope, code):
    """由 (ng_scope, code) 攞 structured rec（撞號先試 exact）。"""
    g = (ng_scope == "gaming")
    c = _norm(code)
    return b2s.get((g, c)) or b2s.get((not g, c)) or {}


def b2text(b2s, ng_scope, code, limit=1400):
    """structured rec → 有 label 嘅文字（餵 LLM 用；比盲抓清楚好多）。"""
    rec = b2rec(b2s, ng_scope, code)
    if not rec:
        return ""
    parts = [f"{k}：{rec[k]}" for k in B2_ORDER if rec.get(k)]
    s = "；".join(parts)
    return s[:limit]


def load_biao2(folder, entity, log=lambda *a: None):
    """{(gaming, 正規化碼): [finding 文字…]}。best-effort，逐檔逐 sheet try。
    （舊版盲抓；新 code 請用 load_biao2_struct + b2text。）"""
    out = {}
    d = Path(folder)
    if not d.exists():
        log(f"（冇 {folder}）"); return out
    allx = [p for p in sorted(d.rglob("*.xls*")) if not p.name.startswith("~$")]
    files = [p for p in allx if _match_entity(p.name, entity.lower()) and "提供附件" not in p.name]
    log(f"表2 folder {folder}：共 {len(allx)} 個 xls*，match「{entity}」{len(files)} 檔")
    for p in files:
        log(f"  ✓ {p.name}")
    if not files and allx:
        log(f"  ⚠ 一個都 match 唔到！全部檔名：" + " | ".join(p.name for p in allx[:20]))
    for p in files:
        gaming = ("博監局" in p.name)     # 博監局檔 = 博彩項目；其餘範疇 = 非博彩
        try:
            wb = IB.load_wb(p)
        except Exception as e:
            log(f"  ⚠ 開唔到 {p.name}: {e}"); continue
        for sn in wb.sheetnames:
            try:
                ws = wb[sn]
                rows = []
                for i, r in enumerate(ws.iter_rows(values_only=True)):
                    rows.append(r)
                    if i > 700:
                        break
                ncol = max((len(r) for r in rows if r), default=0)
                if ncol == 0:
                    continue
                # code 欄 = 匹配 _CODE_RE 最多嘅欄
                best, bestn = None, 0
                for ci in range(ncol):
                    n = sum(1 for r in rows if ci < len(r) and r[ci] is not None
                            and _CODE_RE.match(re.sub(r"\s+", "", str(r[ci]))))
                    if n > bestn:
                        bestn, best = n, ci
                if best is None or bestn < 2:
                    continue
                log(f"  · {p.name}｜{sn}：code欄 col{best}（{bestn}個碼，{'博彩' if gaming else '非博彩'}）")
                for r in rows:
                    if best < len(r) and r[best] is not None \
                            and _CODE_RE.match(re.sub(r"\s+", "", str(r[best]))):
                        code = _norm(r[best])
                        snips = []
                        for c in r:
                            if c is None:
                                continue
                            s = str(c).replace("\n", " ").strip()
                            if len(s) >= 30 and not _JUNK_RE.match(s):
                                snips.append(s[:400])
                        if snips:
                            out.setdefault((gaming, code), [])
                            for s in snips[:4]:
                                if s not in out[(gaming, code)]:
                                    out[(gaming, code)].append(s)
            except Exception as e:
                log(f"  ⚠ {p.name}｜{sn}: {e}")
    log(f"表2：{len(files)} 檔 → {len(out)} 個 (gaming,碼) 有 finding")
    return out


def b2look(b2, ng_scope, code, joiner="　"):
    """由 (ng_scope, code) 攞表2 finding 文字（合併），撞號用 exact 先。"""
    g = (ng_scope == "gaming")
    c = _norm(code)
    snips = b2.get((g, c)) or b2.get((not g, c)) or []
    return joiner.join(snips)


def main():
    args = sys.argv[1:]
    entity = None
    if "--entity" in args:
        i = args.index("--entity"); entity = args[i + 1].lower(); del args[i:i + 2]
    folder = args[0] if args else "data/表2"
    print(f"=== biao2 structured 診斷（entity={entity or 'mgm'}）===")
    b2 = load_biao2_struct(folder, entity or "mgm", log=print)
    from collections import Counter
    cnt = Counter(k for rec in b2.values() for k in rec)
    print(f"\n>>> {len(b2)} 個 (博彩?,碼)；逐個概念抽到幾多個項目：")
    for k in B2_ORDER:
        if cnt.get(k):
            print(f"   {k:<12} {cnt[k]:>4} 個項目")
    miss = [k for k in B2_ORDER if not cnt.get(k)]
    if miss:
        print(f"   ○ 0 個項目：{'、'.join(miss)}"
              f"（『跨司工作組審閱意見』block 源頭仲係空白，唔係抽唔到）")
    print("\n>>> 頭 6 個【有內容】嘅項目逐欄：")
    for (g, c), rec in sorted(b2.items(), key=lambda kv: -len(kv[1]))[:6]:
        print(f"\n[{'博彩' if g else '非博彩'} 項目{c}]")
        for k in B2_ORDER:
            if rec.get(k):
                print(f"   {k}：{rec[k][:160]}")


if __name__ == "__main__":
    main()
