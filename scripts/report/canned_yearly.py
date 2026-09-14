#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
canned_yearly.py — 罐頭版入面【明年要改】嘅嘢清單。

點解要：呢條 pipeline 係做明年 report automation 嘅格式底。罐頭 46 版入面，
骨架模式嗰批（架構圖／走訪相／藝術品／補充圖片）已經換成佔位框，但其餘 33 版
係【逐字抄咗 2025 年】—— 前置聲明、4.4 管理流程、4.5 編制基礎、4.6 審查程序匯總
（抽樣 1,015 筆嗰類）、六項KPI、附件工作範圍。明年冇新 golden 可抽，要沿用呢份，
所以要逐個揾出邊度寫住年份／金額／期間去改。

呢個 script 唔改嘢，只係出清單：逐版逐個 shape，標出
    [年]  出現年份（2023-2030）
    [額]  帶單位嘅金額（億／萬）
    [數]  百分比、件數、筆數
並照 shape 分類（標題／表格格仔／方框／文字），方便逐項處理。

用法：
    python scripts\\report\\canned_yearly.py                    # 預設 mgm
    python scripts\\report\\canned_yearly.py --entity mgm --max 70
    python scripts\\report\\canned_yearly.py --json conf\\local\\canned_mgm.json

出 results\\canned_yearly.txt（UTF-8）。
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

YEAR = re.compile(r"20(2[0-9]|30)")
# breadcrumb 六大章名 —— 每張罐頭版都帶住（golden 自己嗰條），重複 4 次 x 40 版 ≈ 160 個，
# 淨係報數量就淹晒真正要改嘅嘢。呢啲唔使逐版改，改 layout.SECTIONS 就一次過搞掂。
try:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from layout import SECTIONS as _SECS
except Exception:
    _SECS = []
CRUMB = {re.sub(r"\s+", "", x) for x in _SECS}
AMT = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:億|萬)")
NUM = re.compile(r"\d[\d,]*(?:\.\d+)?\s*(?:%|個|筆|件|家|場|次|名|間|項)")
OUT = []


def P(s=""):
    OUT.append(str(s))


def _texts(sh):
    """一個 shape 入面所有文字 → [(類型, 文字)]。"""
    k = sh.get("kind")
    if k == "table":
        for ri, row in enumerate(sh.get("cells") or []):
            for c in row:
                t = (c.get("t") or "").strip()
                if t:
                    yield (f"表r{ri}", t)
    elif k in ("text", "shape"):
        t = (sh.get("text") or "").strip()
        if t:
            yield ("方框" if k == "shape" else "文字", t)


def run(path, maxlen):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    slides = data.get("slides", [])
    P(f"### {path}　{len(slides)} 版\n")
    P("明年沿用呢份罐頭嘅話，下面每一項都要人手檢查／更新。")
    P("（骨架版已經換成佔位符，所以唔會出現喺度 —— 出現嘅就係真係抄咗今年內容嗰啲。）\n")
    tot = Counter()
    for s in slides:
        hits = []
        for sh in s.get("shapes") or []:
            for kind, t in _texts(sh):
                flat = re.sub(r"\s+", "", t)
                if "〔" in t or flat in CRUMB:      # 骨架佔位符／breadcrumb 章名，唔算
                    continue
                tags = []
                if YEAR.search(t):
                    tags.append("年")
                if AMT.search(t):
                    tags.append("額")
                if NUM.search(t):
                    tags.append("數")
                if tags:
                    hits.append((kind, tags, t))
        if not hits:
            continue
        title = (s.get("title") or s.get("marker") or "（冇標題）")[:46]
        P(f"── s{s['n']:<4}{title}　{len(hits)} 項")
        for kind, tags, t in hits:
            tot.update(tags)
            body = re.sub(r"\s+", " ", t)
            P(f"     [{'/'.join(tags)}] {kind:<5} {body[:maxlen]}"
              + ("…" if len(body) > maxlen else ""))
        P()
    P(f"→ 合共 年 {tot['年']}、額 {tot['額']}、數 {tot['數']} 項要逐年更新")
    P("  · [年] 多數係改一個數字（2025→2026）")
    P("  · [額]/[數] 係今年嘅實數（抽樣量、涉及項目數…）→ 明年要由當年數據換返")
    P("  · 改完之後 build_report 唔使郁，罐頭 JSON 改咗就即刻生效")
    P(f"  · 已經隔走 breadcrumb 六大章名（{len(CRUMB)} 個，每版重複 4 次）——"
      " 嗰啲改 layout.SECTIONS 一次過搞掂，唔使逐版執")
    # code 層嘅年份硬編碼 —— 罐頭改完都仲要改呢度
    P("\n══ code 入面寫死年份嘅地方（罐頭改完都仲要改）")
    root = Path(__file__).resolve().parent
    for f in sorted(root.glob("*.py")):
        if f.name in ("canned_yearly.py", "build_report.py"):
            continue                              # build_report 係 bundle 產物
        try:
            lines = f.read_text(encoding="utf-8").splitlines()
        except Exception:
            continue
        hit = [(i, ln) for i, ln in enumerate(lines, 1)
               if YEAR.search(ln) and not ln.lstrip().startswith("#")
               and ("年度" in ln or "年繼續" in ln)]
        for i, ln in hit[:6]:
            P(f"  {f.name}:{i}  {ln.strip()[:96]}")
        if len(hit) > 6:
            P(f"  {f.name}    …另外 {len(hit) - 6} 行")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    av = sys.argv[1:]
    ent = av[av.index("--entity") + 1] if "--entity" in av else "mgm"
    maxlen = int(av[av.index("--max") + 1]) if "--max" in av else 90
    src = Path(av[av.index("--json") + 1]) if "--json" in av else \
        Path("conf/local") / f"canned_{ent}.json"
    if not src.exists():
        print(f"✗ 搵唔到 {src}（先跑 extract_canned.py）"); return
    run(str(src), maxlen)
    txt = "\n".join(OUT)
    dest = Path("results") if Path("results").is_dir() else Path(".")
    f = dest / "canned_yearly.txt"
    f.write_text(txt, encoding="utf-8")
    print(txt)
    print(f"\n✓ 已寫 {f}（UTF-8，{len(OUT)} 行）")
    print("⚠ 內容係客戶報告原文 —— results/ 已 gitignore，唔好 commit")


if __name__ == "__main__":
    main()
