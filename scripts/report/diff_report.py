#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
diff_report.py — 收斂用：項目組原報告 pptx（golden） vs 我哋生成嘅 pptx，出一張差異清單。

點解唔做幾何 diff：原報告啲數字表【全部係 Tableau 截圖】（inspect_pptx --fmt 實測），
我哋出 native 表，shape 類型本身唔同，逐格對冇意義。但原報告嘅【敘述文字】入面
帶住全部關鍵數字（［項目數］、［金額］、［金額］調整、［金額］人工成本…），
嗰啲先係真 tie target。所以呢個工具對三樣：

  ① 章節對照   golden 有邊啲子節 → 我哋出咗未（捉「成節冇做」）
  ② 數字收斂   golden 敘述每個帶單位嘅數字 → 我哋同章有冇同一個數（捉「數唔啱／冇講」）
  ③ 罐頭文字   golden 有成段、我哋完全冇 → 直接就係要抄嘅 boilerplate 清單

用法：
    python scripts\\report\\diff_report.py "MGM…報告.pptx" mgm_report_llm.pptx
    python scripts\\report\\diff_report.py golden.pptx ours.pptx --canned   # 淨係出 ③

出 results\\diff_report.txt（UTF-8）。每輪修完重跑，✗ 數應該一路跌 —— 呢個就係停機條件。
"""
import re
import sys
from collections import defaultdict
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("✗ pip install python-pptx"); sys.exit(1)

EMU_IN = 914400.0

# 初稿／模板留低嘅工作痕跡同 placeholder —— 唔算內容
MARKERS = ("已更新表格", "定稿後還需手動更新", "目錄手動修改為繁體字",
           "DO NOT DELETE", "Workspace (", "Click to edit", "单击以编辑",
           "此处插入", "此处添加", "点击图标", "Click icon", "‹#›")

# 帶單位先算「報告數字」；淨係一個裸數字多數係頁碼／編號，噪音太大。
NUM = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*(億|萬|%|個|項|次|家|間|版)")


def _in(v):
    return (v or 0) / EMU_IN


def _walk(shapes):
    """遞迴攤平 group。"""
    for sh in shapes:
        if sh.shape_type is not None and str(sh.shape_type).startswith("GROUP"):
            try:
                yield from _walk(sh.shapes)
                continue
            except Exception:
                pass
        yield sh


def _texts(slide, W):
    """一版嘅有效文字（畫布內、非 marker）。表格 cell 都收。"""
    out = []
    for sh in _walk(slide.shapes):
        try:
            x, w = _in(sh.left), _in(sh.width)
        except Exception:
            x = w = 0
        if x + w < 0.05 or x > W - 0.05:          # 完全喺畫布外（原報告泊住嘅舊版本）
            continue
        chunks = []
        if getattr(sh, "has_table", False):
            for r in sh.table.rows:
                for c in r.cells:
                    chunks.append(c.text)
        elif sh.has_text_frame:
            chunks.append(sh.text_frame.text)
        for t in chunks:
            t = (t or "").strip()
            if t and not any(m in t for m in MARKERS):
                out.append(t)
    return out


def _crumb(slide, W):
    """「章節 | 子題」麵包屑（原報告固定 y≈0.50、含 ' | '）。回 (章, 子題)。"""
    best = None
    for sh in _walk(slide.shapes):
        if not sh.has_text_frame:
            continue
        t = (sh.text_frame.text or "").strip()
        if "|" not in t and "｜" not in t:
            continue
        y = _in(sh.top)
        if 0.2 < y < 0.7 and (best is None or y < best[0]):
            best = (y, t)
    if not best:
        return ("", "")
    parts = re.split(r"\s*[|｜]\s*", best[1], maxsplit=1)
    return (parts[0].strip(), parts[1].strip() if len(parts) > 1 else "")


def _nums(texts):
    """{(數字, 單位): 出處片段}"""
    got = {}
    for t in texts:
        for m in NUM.finditer(t):
            key = (m.group(1).replace(",", ""), m.group(2))
            if key not in got:
                s = max(0, m.start() - 24)
                got[key] = t[s:m.end() + 16].replace("\n", " ")
    return got


def _scan(path):
    """→ [(版號, 章, 子題, [文字…])]"""
    prs = Presentation(str(path))
    W = _in(prs.slide_width)
    out = []
    for i, sl in enumerate(prs.slides, 1):
        ch, sub = _crumb(sl, W)
        out.append((i, ch, sub, _texts(sl, W)))
    return out


def _norm_sub(s):
    """子題比對用：去節號（要有點，唔好連「2025年度…」個年份都食咗）、去（1/2）、去空白。"""
    s = re.sub(r"^\d+(?:\.\d+)+\s*", "", s or "")
    s = re.sub(r"[（(]\s*\d+\s*/\s*\d+\s*[)）]", "", s)
    return re.sub(r"\s+", "", s)


def run(gold_path, ours_path, canned_only=False):
    L = []
    P = L.append
    gold, ours = _scan(gold_path), _scan(ours_path)
    P(f"### golden  {Path(gold_path).name}  {len(gold)} 版")
    P(f"### ours    {Path(ours_path).name}  {len(ours)} 版\n")

    # ── ① 章節對照 ───────────────────────────────────────────────
    g_by_sub, o_by_sub = defaultdict(list), defaultdict(list)
    for i, ch, sub, _ in gold:
        if sub:
            g_by_sub[(ch, _norm_sub(sub))].append(i)
    for i, ch, sub, _ in ours:
        if sub:
            o_by_sub[_norm_sub(sub)].append(i)

    if not canned_only:
        P("══ ① 章節對照（golden 每個子節 → 我哋出咗未）")
        miss = 0
        for (ch, sub), gi in g_by_sub.items():
            oi = o_by_sub.get(sub, [])
            mark = "✓" if oi else "✗ 未做"
            miss += 0 if oi else 1
            P(f"  {mark}  [{ch}] {sub}"
              f"　golden {_rng(gi)}　→ 我哋 {_rng(oi) if oi else '—'}")
        extra = [s for s in o_by_sub if not any(s == k[1] for k in g_by_sub)]
        for s in extra:
            P(f"  +   我哋多咗：{s}　({_rng(o_by_sub[s])})")
        P(f"\n  → golden {len(g_by_sub)} 個子節，未做 {miss} 個，我哋多出 {len(extra)} 個")

        # ── ② 數字收斂 ───────────────────────────────────────────
        P("\n\n══ ② 數字收斂（golden 敘述帶單位嘅數字 → 我哋【同章】有冇同一個）")
        g_ch, o_ch = defaultdict(list), defaultdict(list)
        for _, ch, _, tx in gold:
            g_ch[ch] += tx
        for _, ch, _, tx in ours:
            o_ch[ch] += tx
        tot_hit = tot_miss = 0
        for ch in sorted(g_ch):
            if not ch:
                continue
            gn, on = _nums(g_ch[ch]), set(_nums(o_ch.get(ch, [])))
            hit = [k for k in gn if k in on]
            mis = [k for k in gn if k not in on]
            tot_hit += len(hit); tot_miss += len(mis)
            P(f"\n  ── {ch}　golden {len(gn)} 個數　對到 {len(hit)}　對唔到 {len(mis)}")
            if ch not in o_ch:
                P("     （我哋成章都未做）")
                continue
            for k in mis:
                P(f"     ✗ {k[0]}{k[1]}　…{gn[k]}…")
        P(f"\n  → 合計 對到 {tot_hit}、對唔到 {tot_miss}"
          f"（{tot_hit / max(1, tot_hit + tot_miss) * 100:.1f}% 收斂）")

    # ── ③ 罐頭文字 ───────────────────────────────────────────────
    P("\n\n══ ③ 罐頭文字（golden 有成段 ≥60 字、我哋全份都搵唔到）")
    ours_all = "".join(t for _, _, _, tx in ours for t in tx)
    ours_all = re.sub(r"\s+", "", ours_all)
    n = 0
    for i, ch, sub, tx in gold:
        for t in tx:
            body = re.sub(r"\s+", "", t)
            if len(body) < 60 or body[:40] in ours_all:
                continue
            n += 1
            P(f"\n  [golden s{i}｜{ch or '—'}｜{sub or '—'}]  {len(body)} 字")
            P(f"    {t[:180].replace(chr(10), ' ⏎ ')}…")
    P(f"\n  → {n} 段罐頭未接")

    txt = "\n".join(L)
    dest = Path("results") if Path("results").is_dir() else Path(".")
    f = dest / "diff_report.txt"
    f.write_text(txt, encoding="utf-8")
    print(txt)
    print(f"\n✓ 已寫 {f}（UTF-8）")


def _rng(nums):
    if not nums:
        return "—"
    out, s, p = [], nums[0], nums[0]
    for x in nums[1:]:
        if x == p + 1:
            p = x; continue
        out.append(f"{s}-{p}" if p > s else f"{s}"); s = p = x
    out.append(f"{s}-{p}" if p > s else f"{s}")
    return ",".join(out)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    a = [x for x in sys.argv[1:] if not x.startswith("--")]
    if len(a) < 2:
        print(__doc__); return
    run(a[0], a[1], canned_only="--canned" in sys.argv)


if __name__ == "__main__":
    main()
