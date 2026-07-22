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


def load_biao2(folder, entity, log=lambda *a: None):
    """{(gaming, 正規化碼): [finding 文字…]}。best-effort，逐檔逐 sheet try。"""
    out = {}
    d = Path(folder)
    if not d.exists():
        log(f"（冇 {folder}）"); return out
    files = [p for p in sorted(d.rglob("*.xls*"))
             if _match_entity(p.name, entity.lower()) and not p.name.startswith("~$")
             and "提供附件" not in p.name]
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
    b2 = load_biao2(folder, entity or "mgm", log=print)
    for (g, c), snips in list(b2.items())[:12]:
        print(f"\n[{'博彩' if g else '非博彩'} 項目{c}] {len(snips)} 段")
        for s in snips[:2]:
            print("  ", s[:110])


if __name__ == "__main__":
    main()
