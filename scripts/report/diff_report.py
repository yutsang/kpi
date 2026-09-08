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
    python scripts\\report\\diff_report.py golden.pptx ours.pptx --brief    # 只出摘要（貼返用）
    python scripts\\report\\diff_report.py golden.pptx ours.pptx --canned   # 淨係出 ③

⚠ --brief 只影響【印出嚟】嘅嘢；`results/diff_report.txt` 一樣係嗰份摘要，
  想睇逐項就唔好加 --brief。

出 results\\diff_report.txt（UTF-8）。每輪修完重跑，✗ 數應該一路跌 —— 呢個就係停機條件。
"""
import re
import sys
from collections import defaultdict
from difflib import SequenceMatcher as SM
from pathlib import Path

try:
    from pptx import Presentation
except ImportError:
    print("✗ pip install python-pptx"); sys.exit(1)

EMU_IN = 914400.0

# 報告六大章（用嚟由「主要發現」呢類短標籤還原返章名）
SECTIONS = ["2025年度投資計劃執行情況概述", "過往年度投資計劃在2025年繼續執行的審查跟進",
            "本年度審查工作的主要發現", "其他信息", "投資計劃執行報告的六項KPI分析", "附件"]

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
    """一版嘅 (章, 子題)。子題優先讀【畫布外嗰個 UpSlide 章節標記】(y<0)：
    原報告主要發現同附件嗰批版冇麵包屑，節名淨係收喺嗰度（目錄亦係靠佢生成）。
    冇標記先退返讀 y≈0.50 嘅「章節 | 子題」麵包屑。章名喺分隔頁（≥30pt 大字）攞，
    之後逐版帶落去 —— 由 caller 用 _scan() 串起。"""
    marker, crumb, big = "", None, ""
    for sh in _walk(slide.shapes):
        if not sh.has_text_frame:
            continue
        t = (sh.text_frame.text or "").strip()
        if not t or any(m in t for m in MARKERS):
            continue
        y = _in(sh.top)
        if y < 0 and len(t) <= 40 and not marker:          # UpSlide 章節標記
            marker = t
        if ("|" in t or "｜" in t) and 0.2 < y < 0.7 and (crumb is None or y < crumb[0]):
            crumb = (y, t)
        # 分隔頁大標題 = 章名；分隔頁仲有個 48pt 嘅章號「1.」，要隔走（至少 2 個中文字）
        if not big and 2 <= len(t) <= 30 and len(re.findall(r"[一-鿿]", t)) >= 2:
            try:
                sz = max((r.font.size.pt for p in sh.text_frame.paragraphs
                          for r in p.runs if r.font.size), default=0)
            except Exception:
                sz = 0
            if sz >= 30:
                big = t
    ch = sub = ""
    if crumb:
        parts = re.split(r"\s*[|｜]\s*", crumb[1], maxsplit=1)
        ch = parts[0].strip()
        sub = parts[1].strip() if len(parts) > 1 else ""
    if not ch and not big:
        # 主要發現版（原報告 s30-44 同我哋）冇麵包屑，只喺 y≈0.30 寫「主要發現」。
        # 攞嗰個短標籤去 SECTIONS 做「包含」配對，還原章名 —— 唔還原就成章 91 個數
        # 當「未做」，收斂率會無端跌 18 個百分點。
        for sh in _walk(slide.shapes):
            if not sh.has_text_frame:
                continue
            t = (sh.text_frame.text or "").strip()
            if not (2 <= len(t) <= 12) or not (0.2 < _in(sh.top) < 0.45):
                continue
            hit = [s for s in SECTIONS if t in s or s in t]
            if hit:
                ch = hit[0]
                break
    return (big or ch, marker or sub, bool(big))


BARE = re.compile(r"\d[\d,]*(?:\.\d+)?")


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


def _bare(texts):
    """我哋表格入面嘅裸數字（單位喺欄名，唔會逐格寫「萬」）。"""
    out = set()
    for t in texts:
        for m in BARE.finditer(t):
            out.add(m.group(0).replace(",", ""))
    return out


def source_nums(entity):
    """掃我哋【源頭】（表2 + 投資項目清單）嘅文字，出一個數字集。
    用嚟拆開「對唔到」：源頭有 = 我哋 pipeline 掉咗（可修）；源頭都冇 = auditor 現場
    做出嚟嘅數（feed 冇，收唔到，唔好再追）。"""
    import glob
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None, None, []
    # 由 data/ 遞歸搵，唔好寫死 folder 名（之前寫死 "data/表2" 撞唔到，得清單入到，
    # 表2 七個檔全部漏咗，令【源冇】報大數）。
    root = Path("data")
    if not root.is_dir():
        return None, None, []
    # ⚠ 只可以掃【項目組畀嘅源頭】—— 表2 同 投資項目清單。
    #   data/ 下面仲有一大堆我哋自己 pipeline 嘅輸出（tableau_*/[ent]_master_audit_*/
    #   *_kpi_report_* 等），掃咗佢哋等於自己對自己，咩數都撞到，【源有】會發脹。
    SRC_KEY = ("表二", "表2", "投资项目清单", "投資項目清單", "投资項目清單", "投資项目清单")
    e, files = entity.lower(), []
    for f in root.rglob("*.xls*"):
        if f.name.startswith("~$") or e not in f.name.lower():
            continue
        hay = f.name + "|" + f.parent.name
        if not any(k in hay for k in SRC_KEY):
            continue
        files.append(str(f))
    files = sorted(set(files))
    if not files:
        return None, None, []
    pairs, bare = set(), set()
    for f in files:
        try:
            wb = load_workbook(f, read_only=True, data_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            try:
                rows = ws.iter_rows(values_only=True)
            except Exception:
                continue
            for row in rows:
                for v in row:
                    if v is None:
                        continue
                    s = str(v)
                    if not any(c.isdigit() for c in s):
                        continue
                    for m in NUM.finditer(s):
                        pairs.add((m.group(1).replace(",", ""), m.group(2)))
                    for m in BARE.finditer(s):
                        bare.add(m.group(0).replace(",", ""))
        try:
            wb.close()
        except Exception:
            pass
    return pairs, bare, files


def _fmt(v):
    """3414.0 → '3414'；1.9 → '1.9'"""
    return f"{v:.10g}"


def _matched_src(key, pairs, bare):
    """判斷 golden 個數喺唔喺我哋源頭 —— 比 _matched 嚴好多。
    源頭係 spreadsheet，成千上萬個裸數字，隨便一格有個 70 就會令「70%」當中咗。所以：
      · %／次／個／項／家／間／版 —— 一定要【連單位一齊】喺源頭文字出現先算
      · 萬／億 —— 可以夾裸數字，但值要 ≥100（唔好俾「12萬」撞到一個 12）"""
    n, u = key
    if key in pairs:
        return True
    # ★ 2026-09-08 最後收緊：一律要【連單位】喺源頭文字出現，唔再夾裸數字。
    #   之前萬／億 容許夾裸數字（值≥100），結果表2 任何一格有個 356，golden 嘅
    #   「［金額］」就當【源有】—— 但「某格有 356」≠「表2 有段文字寫住
    #   『［金額］』並講明佢係乜」。照住嗰張清單改 LLM prompt，改極都冇反應
    #   （對到 161 → 161），就係因為呢批根本唔係真嘅「源頭有」。
    if u != "億":
        return False
    try:                                   # 億 ⇄ 萬 只做【單位換算】，仍然要 pair
        return (_fmt(float(n) * 10000), "萬") in pairs
    except ValueError:
        return False


def _matched(key, pairs, bare):
    """golden 一個 (數, 單位) 喺我哋度算唔算對到。
    除咗原樣，仲要試【單位換算】—— golden 敘述寫「1.9億」，我哋表寫「19,000」（萬）；
    golden 寫「3,414萬」，我哋表格淨寫「3,414」。呢啲係同一個數，唔應該當差異。"""
    n, u = key
    if key in pairs or n in bare:
        return True
    try:
        v = float(n)
    except ValueError:
        return False
    alts = []
    if u == "億":
        alts += [(_fmt(v * 10000), "萬"), (_fmt(v * 10000), None)]
    elif u == "萬":
        alts += [(_fmt(v / 10000), "億"), (_fmt(v / 10000), None)]
    for a, au in alts:
        if au and (a, au) in pairs:
            return True
        if au is None and a in bare:
            return True
    # golden 寫「約7.0億」係約數 —— 我哋表出 69,860（萬）就應該當對到。
    # 只有【明寫小數】嘅先開容差（3,414萬 呢類整數仍然要一模一樣）。
    if "." in n:
        dec = len(n.split(".")[1])
        for b in bare:
            try:
                bv = float(b)
            except ValueError:
                continue
            if u == "億" and round(bv / 10000, dec) == v:
                return True
            if u in ("萬", "%") and round(bv, dec) == v:
                return True
    return False


def _scan(path):
    """→ [(版號, 章, 子題, [文字…])]。章名喺分隔頁定落，之後逐版帶落去
    （原報告主要發現／附件嗰批版本身冇章名）。"""
    prs = Presentation(str(path))
    W = _in(prs.slide_width)
    out, cur = [], ""
    for i, sl in enumerate(prs.slides, 1):
        ch, sub, is_div = _crumb(sl, W)
        if ch:
            cur = ch
        out.append((i, cur, "" if is_div else sub, _texts(sl, W)))
    return out


def _norm_sub(s):
    """子題比對用：去節號（要有點，唔好連「2025年度…」個年份都食咗）、去（1/2）、去空白。"""
    s = re.sub(r"^\d+(?:\.\d+)+\s*", "", s or "")
    s = re.sub(r"[（(]\s*\d+\s*/\s*\d+\s*[)）]", "", s)
    return re.sub(r"\s+", "", s)


def run(gold_path, ours_path, canned_only=False, entity=None, brief=False):
    L = []
    P = L.append
    entity = entity or re.split(r"[_.]", Path(ours_path).stem)[0].lower()
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
        miss, near = 0, set()
        for (ch, sub), gi in g_by_sub.items():
            oi, note = o_by_sub.get(sub, []), ""
            if not oi:                                  # 差一兩個字（簡繁／用詞）都當對到，另外標出嚟改
                cand = max(o_by_sub, key=lambda s: SM(None, sub, s).ratio(), default=None)
                if cand and SM(None, sub, cand).ratio() >= 0.8:
                    oi, note = o_by_sub[cand], f"　⚠ 我哋叫「{cand}」"
                    near.add(cand)
            mark = "✓" if (oi and not note) else ("≈ 名唔同" if oi else "✗ 未做")
            miss += 0 if oi else 1
            if not brief or mark != "✓":       # brief：淨印有問題嗰啲
                P(f"  {mark}  [{ch}] {sub}"
                  f"　golden {_rng(gi)}　→ 我哋 {_rng(oi) if oi else '—'}{note}")
        extra = [s for s in o_by_sub
                 if s not in near and not any(s == k[1] for k in g_by_sub)]
        for s in extra:
            P(f"  +   我哋多咗：{s}　({_rng(o_by_sub[s])})")
        P(f"\n  → golden {len(g_by_sub)} 個子節，未做 {miss} 個，我哋多出 {len(extra)} 個")

        # ── ② 數字收斂 ───────────────────────────────────────────
        s_pairs, s_bare, s_files = source_nums(entity)
        P("\n\n══ ② 數字收斂（golden 敘述帶單位嘅數字 → 我哋【同章】有冇同一個）")
        if s_files:
            P(f"   源頭掃咗 {len(s_files)} 個檔（表2＋清單）→ 對唔到嘅會標【源有】/【源冇】")
            for f in s_files:
                P(f"     · {Path(f).name}")
        else:
            P("   （搵唔到 data/表2、data/投資項目清單 → 唔標源頭）")
        g_ch, o_ch = defaultdict(list), defaultdict(list)
        for _, ch, _, tx in gold:
            g_ch[ch] += tx
        for _, ch, _, tx in ours:
            o_ch[ch] += tx
        tot_hit = tot_miss = 0
        fixable, unfixable = [], []
        for ch in sorted(g_ch):
            if not ch:
                continue
            gn = _nums(g_ch[ch])
            pairs, bare = set(_nums(o_ch.get(ch, []))), _bare(o_ch.get(ch, []))
            hit = [k for k in gn if _matched(k, pairs, bare)]
            mis = [k for k in gn if not _matched(k, pairs, bare)]
            tot_hit += len(hit); tot_miss += len(mis)
            P(f"\n  ── {ch}　golden {len(gn)} 個數　對到 {len(hit)}　對唔到 {len(mis)}")
            if ch not in o_ch:
                P("     （我哋成章都未做）")
                continue
            for k in mis:
                tag = ""
                if s_pairs is not None:
                    tag = "【源有】" if _matched_src(k, s_pairs, s_bare) else "【源冇】"
                    (fixable if tag == "【源有】" else unfixable).append((ch, k))
                if not brief:
                    P(f"     ✗ {tag}{k[0]}{k[1]}　…{gn[k]}…")
        P(f"\n  → 合計 對到 {tot_hit}、對唔到 {tot_miss}"
          f"（{tot_hit / max(1, tot_hit + tot_miss) * 100:.1f}% 收斂）")
        if s_pairs is not None:
            P(f"\n  ── 對唔到嘅拆開 ──")
            P(f"  【源有】{len(fixable)} 個　表2／清單搵到，但我哋冇帶出報告 → 可修，呢批先係真 todo")
            P(f"  【源冇】{len(unfixable)} 個　源頭都冇（auditor 現場／訪談得出）→ 收唔到，唔好再追")
            if fixable:
                P("\n  【源有】清單：")
                for ch, k in fixable:
                    P(f"    · [{ch}] {k[0]}{k[1]}")

    # ── ③ 罐頭文字 ───────────────────────────────────────────────
    P("\n\n══ ③ 罐頭文字（golden 有成段 ≥60 字、我哋全份都搵唔到）")
    ours_all = "".join(t for _, _, _, tx in ours for t in tx)
    ours_all = re.sub(r"\s+", "", ours_all)
    seen = {}                       # 同一段喺 golden 連續版重複（導語）→ 合併，唔好列幾次
    for i, ch, sub, tx in gold:
        for t in tx:
            body = re.sub(r"\s+", "", t)
            if len(body) < 60 or body[:40] in ours_all:
                continue
            e = seen.setdefault(body[:80], {"t": t, "n": len(body), "ch": ch,
                                            "sub": sub, "s": []})
            e["s"].append(i)
    for e in (() if brief else seen.values()):
        P(f"\n  [golden s{_rng(e['s'])}｜{e['ch'] or '—'}｜{e['sub'] or '—'}]  {e['n']} 字"
          + ("　（重複 %d 版）" % len(e["s"]) if len(e["s"]) > 1 else ""))
        P(f"    {e['t'][:180].replace(chr(10), ' ⏎ ')}…")
    P(f"\n  → {len(seen)} 段罐頭未接（去重後）")

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
    run(a[0], a[1], canned_only="--canned" in sys.argv, brief="--brief" in sys.argv)


if __name__ == "__main__":
    main()
