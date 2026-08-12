#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
biao2.py — 由「表2」(審查底稿，加密 dicj_kpmg) 逐項目抽 finding 文字，做第 2 個 narrative source
（配清單一齊餵 LLM）。表2 layout 複雜（每 entity×範疇 一檔，per-project 行，欄浮動），
用 best-effort：揾 code 欄（項目N / 字母碼）→ 每有 code 嘅行收長文字 cell（關注事項/分析/管理層/跨司）。

key = (gaming, 正規化碼)。gaming 由檔名判：『博監局』檔＝博彩，其餘＝非博彩。
→ 修返「博彩娛樂場」冇內容（清單博彩碼撞號攞唔到，但表2 博監局有）。

用法（module）：from biao2 import load_biao2; b2 = load_biao2("data/表2", "mgm")
              b2[(True, "19")] → 博彩項目19 嘅 finding 文字 list
"""
import re
import sys
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
    "項目分析意見": ["項目分析意見"],
    "建議接納調整後金額": ["建議接納之調整後金額"],
    "實際情況": ["已投放金額", "截至"],
    "擬投資內容": ["擬投資金額", "擬落實"],
}
# 出 prompt 時嘅次序（最有用擺前）
B2_ORDER = ["調整類型", "關注事項", "調整原因", "關注事項金額", "建議調整金額",
            "建議調整後金額", "KPMG分析", "管理層解釋", "承批公司反饋", "跨司回覆",
            "項目分析意見", "建議接納調整後金額", "實際情況", "擬投資內容"]
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
                            rec[concept] = (cur + "　" + s).strip("　") if cur else s
                            n_field += 1
            except Exception as e:
                log(f"  ⚠ {p.name}｜{sn}: {e}")
    log(f"表2（structured）：{len(files)} 檔 → {len(out)} 個 (gaming,碼)、{n_field} 個欄值")
    return out


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
        print(f"   ⚠ 完全抽唔到：{'、'.join(miss)}")
    print("\n>>> 頭 6 個項目逐欄內容：")
    for (g, c), rec in sorted(b2.items())[:6]:
        print(f"\n[{'博彩' if g else '非博彩'} 項目{c}]")
        for k in B2_ORDER:
            if rec.get(k):
                print(f"   {k}：{rec[k][:160]}")


if __name__ == "__main__":
    main()
