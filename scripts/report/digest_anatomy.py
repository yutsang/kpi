#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
digest_anatomy.py — 把原報告逐版量度出嚟嗰份大 JSON（幾 MB）榨成一份幾百行嘅摘要。

點解要：anatomy.json 逐版逐個 shape 記晒，成 5MB，傳唔到。但我需要嘅唔係原始記錄，
係【聚合】—— 樣式普查、圖片清單、兩欄幾何、marker 全表、導語字數分佈。呢個 script
本機做聚合，出一份貼得返嘅 txt。

★ 唔假設 schema：開頭一定會印返個 JSON 實際嘅結構（第 0 節）。就算 key 名同我預期
  唔同，睇返第 0 節就知點改，唔使傳成份檔。

用法：
    python scripts\\report\\digest_anatomy.py                       # 預設 results/anatomy.json
    python scripts\\report\\digest_anatomy.py 邊個檔.json
    python scripts\\report\\digest_anatomy.py --schema              # 只出第 0 節（最短，先睇結構）

出 results\\anatomy_digest.txt（UTF-8）。
"""
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

CAND = ["results/anatomy.json", "anatomy.json", "results/report_anatomy.json"]
OUT = []


def P(s=""):
    OUT.append(str(s))


def g(d, *names, default=None):
    """逐個 key 名試（schema 唔跟足都撈得返）。支援 'a.b' 一層 nested。"""
    if not isinstance(d, dict):
        return default
    for n in names:
        if "." in n:
            a, _, b = n.partition(".")
            v = g(d.get(a), b, default=None)
            if v is not None:
                return v
        elif d.get(n) is not None:
            return d[n]
    return default


def _slides(data):
    """{"slides":[…]} / [ … ] / {"1":{…}} 都認。"""
    if isinstance(data, list):
        return data
    for k in ("slides", "pages", "data", "records"):
        v = data.get(k) if isinstance(data, dict) else None
        if isinstance(v, list):
            return v
    if isinstance(data, dict) and all(str(k).isdigit() for k in list(data)[:5]):
        return [dict(v, n=v.get("n", int(k))) for k, v in data.items() if isinstance(v, dict)]
    return []


def _lst(sl, *names):
    v = g(sl, *names, default=[])
    return v if isinstance(v, list) else []


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def schema_section(data, sl):
    P("══ 0. JSON 實際結構（我照呢個嚟解讀；key 名唔啱睇呢節就知）")
    P(f"  top-level: {type(data).__name__}"
      + (f"　keys={list(data)[:8]}" if isinstance(data, dict) else ""))
    P(f"  版數: {len(sl)}")
    keys = Counter()
    for s in sl:
        if isinstance(s, dict):
            keys.update(s.keys())
    P(f"  逐版 key（出現次數）：")
    for k, c in keys.most_common():
        P(f"     {k:<18} {c}")
    # 逐個 key 嘅形狀 + 一個非空樣本 —— 我事先唔知 agent 出咗咩 field，
    # 靠呢節先知點寫 extractor（唔使傳成份 5MB JSON 過嚟）。
    P("\n  逐個 key 嘅形狀（type｜非空版數｜樣本）：")
    for k in keys:
        vals = [s.get(k) for s in sl if isinstance(s, dict)]
        nonempty = [v for v in vals if v not in (None, "", [], {})]
        t = type(nonempty[0]).__name__ if nonempty else "全部空"
        extra = ""
        if nonempty and isinstance(nonempty[0], list):
            inner = next((x for v in nonempty for x in v), None)
            extra = (f"　元素={type(inner).__name__}"
                     + (f" keys={list(inner)[:10]}" if isinstance(inner, dict) else ""))
            extra += f"　最長 {max(len(v) for v in nonempty)} 個"
        elif nonempty and isinstance(nonempty[0], dict):
            allk = Counter(kk for v in nonempty if isinstance(v, dict) for kk in v)
            extra = f"　keys={[x for x, _ in allk.most_common(12)]}"
        P(f"     {k:<16} {t:<8} {len(nonempty):>4}/{len(sl)}{extra}")
        if nonempty:
            smp = json.dumps(nonempty[0], ensure_ascii=False)
            P(f"        樣本 {smp[:420]}" + ("…" if len(smp) > 420 else ""))
    P()


def run(path, schema_only=False):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    sl = _slides(data)
    P(f"### {path}　{Path(path).stat().st_size / 1e6:.1f} MB　{len(sl)} 版\n")
    schema_section(data, sl)
    if schema_only or not sl:
        return

    # ── 1. 逐版一行 ────────────────────────────────────────────
    P("══ 1. 逐版摘要（版｜archetype｜text/table/pic/shape｜marker）")
    for s in sl:
        n = g(s, "n", "slide", "index", default="?")
        arch = g(s, "archetype", "type", "kind", default="—")
        c = g(s, "counts", default={}) or {}
        cs = "/".join(str(c.get(k, len(_lst(s, k + "s", k)))) for k in
                      ("text", "table", "picture", "autoshape"))
        mk = (g(s, "marker", "subsection", "upslide", default="") or "")[:34]
        P(f"  s{str(n):>3}  {str(arch):<24} {cs:<12} {mk}")
    P()

    # ── 2. 樣式普查（最重要）──────────────────────────────────
    P("══ 2. 樣式普查（font, size, bold, color → 次數 + 用喺邊）")
    sty, where = Counter(), defaultdict(list)
    for s in sl:
        n = g(s, "n", default="?")
        for t in _lst(s, "texts", "text", "runs"):
            if not isinstance(t, dict):
                continue
            k = (str(g(t, "font", "latin", default="?"))[:22],
                 g(t, "size", "pt", default=None),
                 bool(g(t, "bold", default=False)),
                 str(g(t, "color", "fg", default="?")).upper())
            sty[k] += 1
            if len(where[k]) < 4:
                y = _num(g(t, "y", "top"))
                where[k].append(f"s{n}@y{y:.2f}" if y is not None else f"s{n}")
    P(f"  {'font':<24}{'pt':>6}{'B':>3}  {'色':<8}{'次':>6}  用喺邊")
    for k, c in sty.most_common(40):
        f, sz, b, col = k
        P(f"  {f:<24}{(sz if sz is not None else '?'):>6}{'B' if b else '':>3}  {col:<8}{c:>6}  "
          + ", ".join(where[k]))
    if len(sty) > 40:
        P(f"  …仲有 {len(sty) - 40} 個組合（次數 ≤ {sty.most_common(40)[-1][1]}）")
    # 主題色／繼承字體解唔到 = 呢批數據等於冇 —— 要出聲，唔好當已經量度到
    unres_c = Counter(); unres_f = Counter()
    for (f, sz, b, col), c in sty.items():
        if col.startswith("THEME") or col == "?":
            unres_c[col] += c
        if f in ("?",) or f.startswith("+"):
            unres_f[f] += c
    tot_run = sum(sty.values())
    if unres_c or unres_f:
        P(f"\n  ⚠ 未解析（占 run 總數 {tot_run}）：")
        for k, v in unres_c.most_common():
            P(f"     色 {k:<28} {v:>5} run（{v / tot_run * 100:.0f}%）→ 要查 theme1.xml clrScheme+clrMap 換 RGB")
        for k, v in unres_f.most_common():
            P(f"     字體 {k:<26} {v:>5} run（{v / tot_run * 100:.0f}%）→ 要沿 layout/master/theme 逐層繼承解析")
    P()

    # ── 3. 圖片 ────────────────────────────────────────────────
    # ⚠ 好多版喺畫布外泊住一份舊圖（x 係負數）—— 嗰啲唔算內容，唔好當版式證據。
    P("══ 3. 圖片（畫布內先算；畫布外泊住嘅舊圖另計）")
    npic, off, ext_c = 0, 0, Counter()
    for s in sl:
        for p in _lst(s, "pictures", "pics", "images"):
            if not isinstance(p, dict):
                continue
            x, y = _num(g(p, "x", "left")), _num(g(p, "y", "top"))
            w, h = _num(g(p, "w", "width")), _num(g(p, "h", "height"))
            if x is not None and w is not None and x + w < 0.1:
                off += 1; continue
            npic += 1
            ext_c[str(g(p, "ext", default="?")).lower()] += 1
            px = g(p, "px_w", "pixel_w", "img_w", default="?")
            py = g(p, "px_h", "pixel_h", "img_h", default="?")
            geo = (f"x{x:.2f} y{y:.2f} {w:.2f}x{h:.2f}in"
                   if None not in (x, y, w, h) else "幾何缺")
            P(f"  s{str(g(s, 'n', default='?')):>3}  {geo}　{px}x{py}px　{g(p, 'ext', default='')}")
    P(f"  → 畫布內 {npic} 張（格式：{'、'.join(f'{k} {v}' for k, v in ext_c.most_common())}）"
      f"；畫布外泊住 {off} 張（舊版本，唔理）")
    P("  ※ wmf = 由 Excel／Tableau 貼入嚟嘅向量圖 → 嗰啲『表』冇 cell 資料可抄。\n")

    # ── 4. 原生表 ──────────────────────────────────────────────
    P("══ 4. 原生表（唔係截圖嗰啲）")
    ntb = 0
    for s in sl:
        for t in _lst(s, "tables", "table"):
            if not isinstance(t, dict):
                continue
            ntb += 1
            x, y = _num(g(t, "x", "left")), _num(g(t, "y", "top"))
            w = _num(g(t, "w", "width"))
            P(f"  s{str(g(s, 'n', default='?')):>3}  "
              f"{g(t, 'rows', 'nrow', default='?')}r x {g(t, 'cols', 'ncol', default='?')}c"
              + (f"　x{x:.2f} y{y:.2f} w{w:.2f}" if None not in (x, y, w) else ""))
    P(f"  → 合共 {ntb} 張原生表\n")

    # ── 5. 兩欄版幾何 ──────────────────────────────────────────
    # golden 啲數據表【係圖唔係表】，所以左邊嗰嚿要連 picture 一齊當「表」睇，
    # 唔係就淨係撈到一兩版（第一版寫得只睇 tables，結果得 s31 而且係畫布外嗰個）。
    P("══ 5. 兩欄版（左邊表／圖 ＋ 右邊敘述）：左邊闊 / 右欄起點 / 間距")
    n2 = 0
    for s in sl:
        blocks = []
        for t in _lst(s, "tables", "table") + _lst(s, "pictures", "pics", "images"):
            if not isinstance(t, dict):
                continue
            x, w = _num(g(t, "x", "left")), _num(g(t, "w", "width"))
            y = _num(g(t, "y", "top"))
            if None in (x, w, y) or x + w < 0.1 or y < 0.9:       # 畫布外／頁首唔算
                continue
            blocks.append((x, w, y))
        if not blocks:
            continue
        lx, lw, _ = min(blocks, key=lambda b: b[0])
        if lx + lw > 7.6:                                   # 佔成版闊 → 唔係兩欄版
            continue
        right = [_num(g(q, "x", "left")) for q in _lst(s, "texts", "text")
                 if isinstance(q, dict)
                 and (_num(g(q, "x", "left")) or -9) > lx + lw - 0.05
                 and 0.9 < (_num(g(q, "y", "top")) or 0) < 6.6]
        if not right:
            continue
        rx = min(right); n2 += 1
        P(f"  s{str(g(s, 'n', default='?')):>3}  左邊 x{lx:.2f} w{lw:.2f}（右邊界 {lx + lw:.2f}）"
          f"　右欄 x{rx:.2f}　間距 {rx - lx - lw:.2f}")
    P(f"  → {n2} 版兩欄\n")

    # ── 6. marker 全表 + 異名偵測 ──────────────────────────────
    P("══ 6. UpSlide 章節標記全表（相鄰版名唔同 = 要跟返；繁簡混用會標出）")
    seq = [(g(s, "n", default=i + 1), (g(s, "marker", "subsection", default="") or "").strip())
           for i, s in enumerate(sl)]
    grp, cur = [], None
    for n, m in seq:
        if m != cur:
            grp.append([m, [n]]); cur = m
        else:
            grp[-1][1].append(n)
    for m, ns in grp:
        rng = f"{ns[0]}-{ns[-1]}" if len(ns) > 1 else str(ns[0])
        P(f"  s{rng:<9} {m or '（冇）'}")
    P()
    SIMP = "内为产从众优体医华协单卖厂历压县参双发";  # 常見簡體，繁體稿入面出現就係混用
    mixed = [(n, m) for n, m in seq if any(ch in m for ch in SIMP)]
    if mixed:
        P("  ⚠ marker 有簡體字（繁體稿混用）：")
        for n, m in mixed:
            P(f"     s{n}  {m}")
    near = []
    names = [m for m, _ in grp if m]        # grp 元素係 [marker, [版號…]]
    for i, a in enumerate(names):
        for b in names[i + 1:]:
            if a != b and (a in b or b in a or
                           len(set(a) & set(b)) >= min(len(a), len(b)) * 0.75):
                near.append((a, b))
    if near:
        P("  ⚠ 講同一件事但名唔同（要逐版跟返，唔可以齊用一個名）：")
        for a, b in near[:12]:
            P(f"     「{a}」　vs　「{b}」")
    P()

    # ── 7. 導語字數 ────────────────────────────────────────────
    P("══ 7. 導語（headline）字數分佈")
    by_ch = defaultdict(list)
    for s in sl:
        hl = g(s, "headline", "head", default="") or ""
        if not hl:
            continue
        ch = g(s, "crumb.chapter", "chapter", default="—") or "—"
        by_ch[ch].append(len(re.sub(r"\s+", "", hl)))
    for ch, v in sorted(by_ch.items(), key=lambda kv: -len(kv[1])):
        v.sort()
        P(f"  {ch[:30]:<32} n={len(v):>3}  min {v[0]:>4}  中位 {v[len(v) // 2]:>4}  max {v[-1]:>4}")
    P()


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    av = sys.argv[1:]
    pos = [a for a in av if not a.startswith("--")]
    src = pos[0] if pos else next((c for c in CAND if Path(c).exists()), None)
    if not src or not Path(src).exists():
        print(f"✗ 搵唔到 anatomy JSON（試過 {', '.join(CAND)}）；用第一個引數指明"); return
    run(src, schema_only="--schema" in av)
    txt = "\n".join(OUT)
    dest = Path("results") if Path("results").is_dir() else Path(".")
    f = dest / "anatomy_digest.txt"
    f.write_text(txt, encoding="utf-8")
    print(txt)
    print(f"\n✓ 已寫 {f}（UTF-8，{len(OUT)} 行）")


if __name__ == "__main__":
    main()
