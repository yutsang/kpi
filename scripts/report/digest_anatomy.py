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
    if sl:
        one = min(sl, key=lambda s: len(json.dumps(s, ensure_ascii=False)) if isinstance(s, dict) else 0)
        txt = json.dumps(one, ensure_ascii=False, indent=1)
        P(f"  最細嗰版嘅完整記錄（睇 field 形狀）：")
        for line in txt.splitlines()[:40]:
            P("     " + line[:150])
        if len(txt.splitlines()) > 40:
            P("     …")
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
    P()

    # ── 3. 圖片 ────────────────────────────────────────────────
    P("══ 3. 圖片（邊版有圖、幾大、像素）")
    npic = 0
    for s in sl:
        for p in _lst(s, "pictures", "pics", "images"):
            if not isinstance(p, dict):
                continue
            npic += 1
            x, y = _num(g(p, "x", "left")), _num(g(p, "y", "top"))
            w, h = _num(g(p, "w", "width")), _num(g(p, "h", "height"))
            px = g(p, "px_w", "pixel_w", "img_w", default="?")
            py = g(p, "px_h", "pixel_h", "img_h", default="?")
            geo = (f"x{x:.2f} y{y:.2f} {w:.2f}x{h:.2f}in"
                   if None not in (x, y, w, h) else "幾何缺")
            P(f"  s{str(g(s, 'n', default='?')):>3}  {geo}　{px}x{py}px　{g(p, 'ext', default='')}")
    P(f"  → 合共 {npic} 張圖\n")

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
    P("══ 5. 兩欄版（表左＋敘述右）：左表闊 / 右欄起點 / 間距")
    for s in sl:
        tb = [t for t in _lst(s, "tables", "table") if isinstance(t, dict)]
        if len(tb) != 1:
            continue
        t = tb[0]
        tx, tw = _num(g(t, "x", "left")), _num(g(t, "w", "width"))
        if None in (tx, tw):
            continue
        right = [q for q in _lst(s, "texts", "text")
                 if isinstance(q, dict) and (_num(g(q, "x", "left")) or 0) > tx + tw - 0.05
                 and 1.0 < (_num(g(q, "y", "top")) or 0) < 6.5]
        if not right:
            continue
        rx = min(_num(g(q, "x", "left")) for q in right)
        P(f"  s{str(g(s, 'n', default='?')):>3}  左表 x{tx:.2f} w{tw:.2f}（右邊界 {tx + tw:.2f}）"
          f"　右欄 x{rx:.2f}　間距 {rx - tx - tw:.2f}")
    P()

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
