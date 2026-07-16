#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
align_to_header.py — 對齊 source_1 全部檔到 表頭.xlsx 標準格式（34 欄）。

背景／結構（合併感知，唔係一行一項目）：
  一個項目佔幾行（行數浮動）；左半(B–F)逐行 key-value 明細；右半(是否該司局範疇→尾)
  每欄垂直合併蓋住成個項目 span = 一個項目一個值。項目邊界 = 「是否該司局範疇的項目」
  欄嘅垂直合併範圍。欄位置逐檔浮動 → 一律靠 row6 子表頭 + row5 分組表頭認欄。

映射（group-aware，因第一輪/第二輪子表頭重複，用 (分組,子表頭) 複合鍵）：見 TARGET_SCHEMA。
  尾段：調整原因 = 各非零調整類型「一格列舉」`1、{類型}：{|金額|}萬澳門元`＼n；
        建議調整金額←潛在調整合計(連正負)、建議調整後金額←調整後投資金額、
        建議接納之調整後金額←跨司工作組確認投資金額；該關注事項涉及調整金額←|潛在調整合計|。
  source_2 per-範疇檔按 承批公司+範疇+項目序號 覆蓋審查欄（--with-overlay，source_2 為準）。

輸出：<root>\\_aligned\\ 鏡像 source_1 樹（**永不寫入 source 樹**）。ss/ 同 附件 = 原檔照抄/best-effort。

Windows 跑：
    set PYTHONIOENCODING=utf-8
    # 1) 先驗一個檔（淨吐文字：逐項目抽咗乜 + 會寫入表頭邊欄；唔寫 xlsx）
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace ^
        --only "source_1\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx" --preview
    # 2) 同一個檔真正寫低（開嚟眼睇）
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace ^
        --only "source_1\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx" --with-overlay
    # 3) 冇問題先批全部
    python scripts\\adhoc\\align_to_header.py --root ad-hoc\\workspace --all --with-overlay
"""
from __future__ import annotations

import argparse
import io
import re
import shutil
from copy import copy
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side

PASSWORD = "dicj_kpmg"
EXCEL_EXT = {".xlsx", ".xlsm", ".xls"}
HDR_SCAN = 12                       # 掃頭幾行揾子表頭行
GATE_NOADJ = True                   # 冇調整嘅項目 → 建議調整欄留空（唔填 0）；--fill-zero-adj 關掉
_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_WRAP_TOP = Alignment(wrap_text=True, vertical="top", horizontal="left")


# ── 正規化 ────────────────────────────────────────────────────────────────
_PUNC = str.maketrans({
    "（": "(", "）": ")", "：": ":", "，": ",", "、": ",", "／": "/",
    "　": "", " ": "", " ": "", "\"": "", "“": "", "”": "", "'": "",
})


def _s(x) -> str:
    return "" if x is None else str(x).strip().replace("\r", "")


def nkey(x) -> str:
    """認欄用嘅正規化鍵：去空白/引號、全形標點→半形。"""
    return _s(x).replace("\n", "").translate(_PUNC)


# 子表頭同義變體（source 名 → 表頭 canonical 名 nkey）
VARIANTS = {
    # 各司局專名「XX整體的回覆」+「跨司工作組整體的回覆」→ 表頭「跨司工作組的回覆」
    nkey("跨司工作組整體的回覆"): nkey("跨司工作組的回覆"),
    nkey("博監局整體的回覆"): nkey("跨司工作組的回覆"),
    nkey("整體的回覆"): nkey("跨司工作組的回覆"),
    nkey("畢馬威的分析"): nkey("KPMG分析"),   # 只喺第一/二輪組；潛在調整組另有規則
}
_RE_WHOLE_REPLY = re.compile(r".*整體的回覆$")


def canon_sub(raw: str) -> str:
    k = nkey(raw)
    k = re.sub(r"\([^()]*\)$", "", k)          # 去尾部括號註（「(萬澳門元)」「(如不一致…)」）
    if k in VARIANTS:
        return VARIANTS[k]
    if _RE_WHOLE_REPLY.match(_s(raw)):
        return nkey("跨司工作組的回覆")
    return k


# ── 表頭目標 schema（34 欄；idx = 表頭資料欄位置；rule 見 resolve）──────────
# group: row5 分組（"" = 無/左半/單欄）；sub: row6 子表頭；rule: 點填
G1 = "與跨司工作組的第一輪問題諮詢"
G2 = "與跨司工作組的第二輪問題諮詢"
GT = "與承批公司溝通潛在調整事項"          # 用前綴 match（表頭全名長）
LEFT = "__LEFT__"

TARGET_SCHEMA = [
    # 左半（positional，右對齊照抄）——只標記，唔逐欄配名
    {"grp": "", "sub": "是否該司局範疇的項目", "rule": ("copy_any", "是否該司局範疇的項目")},
    # 第一輪
    {"grp": G1, "sub": "是否有希望諮詢的問題", "rule": ("copy", G1, "是否有希望諮詢的問題")},
    {"grp": G1, "sub": "問題狀態", "rule": ("copy", G1, "問題狀態")},
    {"grp": G1, "sub": "KPMG提出日期", "rule": ("copy", G1, "KPMG提出日期")},
    {"grp": G1, "sub": "KPMG分析", "rule": ("copy", G1, "KPMG分析")},
    {"grp": G1, "sub": "承批公司管理層解釋", "rule": ("copy", G1, "承批公司管理層解釋")},
    {"grp": G1, "sub": "KPMG希望進一步向跨司工作組瞭解的事項", "rule": ("copy", G1, "KPMG希望進一步向跨司工作組瞭解的事項")},
    {"grp": G1, "sub": "跨司工作組的回覆", "rule": ("copy", G1, "跨司工作組的回覆")},
    # 第二輪
    {"grp": G2, "sub": "問題狀態", "rule": ("copy", G2, "問題狀態")},
    {"grp": G2, "sub": "KPMG提出日期", "rule": ("copy", G2, "KPMG提出日期")},
    {"grp": G2, "sub": "KPMG分析", "rule": ("copy", G2, "KPMG分析")},
    {"grp": G2, "sub": "承批公司管理層解釋", "rule": ("copy", G2, "承批公司管理層解釋")},
    {"grp": G2, "sub": "KPMG希望進一步向跨司工作組瞭解的事項", "rule": ("copy", G2, "KPMG希望進一步向跨司工作組瞭解的事項")},
    {"grp": G2, "sub": "跨司工作組的回覆", "rule": ("copy", G2, "跨司工作組的回覆")},
    # 與承批公司溝通潛在調整事項（source_1 只有前 2 條 + 跨司反饋/畢馬威分析）
    {"grp": GT, "sub": "畢馬威關注事項", "rule": ("copy", GT, "畢馬威關注事項")},
    {"grp": GT, "sub": "承批公司的反饋意見", "rule": ("copy", GT, "承批公司的反饋意見")},
    {"grp": GT, "sub": "是否需進一步與跨司工作組溝通", "rule": ("blank",)},               # source_2
    {"grp": GT, "sub": "需溝通關注事項", "rule": ("blank",)},                             # source_2
    {"grp": GT, "sub": "該關注事項涉及調整金額", "rule": ("abs_total",)},                  # |潛在調整合計|
    {"grp": GT, "sub": "跨司工作組主責部門針對該關注事項已給的反饋意見", "rule": ("copy", GT, "跨司工作組的反饋意見")},
    {"grp": GT, "sub": "KPMG需與跨司工作組進一步確認的問題", "rule": ("copy", GT, "畢馬威的分析")},   # 假設
    {"grp": GT, "sub": "跨司工作組最新反饋意見", "rule": ("blank",)},                      # source_2
    # 畢馬威審查意見
    {"grp": "畢馬威審查意見", "sub": "建議調整金額", "rule": ("seed", "潛在調整合計")},
    {"grp": "畢馬威審查意見", "sub": "調整原因", "rule": ("enum",)},
    {"grp": "畢馬威審查意見", "sub": "建議調整後金額", "rule": ("seed", "調整後投資金額")},
    # 跨司工作組審閱意見
    {"grp": "跨司工作組審閱意見", "sub": "項目分析意見", "rule": ("blank",)},
    {"grp": "跨司工作組審閱意見", "sub": "建議接納之調整後金額", "rule": ("seed", "跨司工作組確認投資金額")},
    {"grp": "", "sub": "對比畢馬威審查意見與跨司工作組審閱意見是否一致", "rule": ("blank",)},
]

# 潛在調整組內「唔係調整類型」嘅子表頭（其餘全部當調整類型入列舉）
ADJ_NON_TYPE = {nkey(x) for x in
                ["申報投資金額", "潛在調整合計", "調整後投資金額", "跨司工作組確認投資金額"]}
# source_2 覆蓋會蓋嘅表頭欄（sub nkey）
OVERLAY_SUBS = [nkey(x) for x in
                ["是否需進一步與跨司工作組溝通", "需溝通關注事項", "該關注事項涉及調整金額",
                 "跨司工作組主責部門針對該關注事項已給的反饋意見",
                 "KPMG需與跨司工作組進一步確認的問題", "跨司工作組最新反饋意見",
                 "畢馬威關注事項", "承批公司的反饋意見"]]


# ── I/O ──────────────────────────────────────────────────────────────────
def load_wb(path: Path):
    """非 read-only（要 merged_cells）；加密就 dicj_kpmg 解。"""
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception:
        import msoffcrypto
        buf = io.BytesIO()
        with open(path, "rb") as f:
            off = msoffcrypto.OfficeFile(f)
            off.load_key(password=PASSWORD)
            off.decrypt(buf)
        buf.seek(0)
        return openpyxl.load_workbook(buf, data_only=True)


def num(v):
    """試 parse 成數字（去逗號）；唔係數字回 None。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = _s(v).replace(",", "").replace("，", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", s):
        return float(s)
    return None


def fmt_amt(x: float) -> str:
    return str(int(x)) if abs(x - round(x)) < 1e-9 else f"{x:.2f}".rstrip("0").rstrip(".")


# ── 表頭解析 ──────────────────────────────────────────────────────────────
def group_map(ws, grow: int, maxcol: int) -> dict[int, str]:
    """row5 分組表頭：橫向合併 fill-forward → {col: group_nkey}。"""
    gm: dict[int, str] = {}
    merged = {}
    for m in ws.merged_cells.ranges:
        if m.min_row <= grow <= m.max_row and m.max_col > m.min_col:
            v = _s(ws.cell(m.min_row, m.min_col).value)
            for c in range(m.min_col, m.max_col + 1):
                merged[c] = v
    for c in range(1, maxcol + 1):
        gm[c] = nkey(merged.get(c, _s(ws.cell(grow, c).value)))
    return gm


def detect_sub_row(ws, maxrow: int, maxcol: int) -> int:
    """揀同已知子表頭重疊最多嗰行做 row6（子表頭）。"""
    known = {nkey(r["sub"]) for r in TARGET_SCHEMA} | {nkey("非博彩投資項目性質")}
    best, best_hit = 1, -1
    for r in range(1, min(HDR_SCAN, maxrow) + 1):
        vals = {canon_sub(ws.cell(r, c).value) for c in range(1, maxcol + 1)}
        hit = len(known & vals)
        if hit > best_hit:
            best_hit, best = hit, r
    return best


def find_anchor_col(ws, subrow: int, maxcol: int) -> int | None:
    tgt = nkey("是否該司局範疇的項目")
    for c in range(1, maxcol + 1):
        if canon_sub(ws.cell(subrow, c).value) == tgt:
            return c
    return None


# ── 抽取一個 sheet 嘅項目 ─────────────────────────────────────────────────
class Project:
    __slots__ = ("r0", "r1", "left", "by_gs", "adj", "seed", "seq", "override")

    def __init__(self, r0, r1):
        self.r0, self.r1 = r0, r1
        self.left: list[list] = []          # 逐行左半原值 [[cell,...],...]
        self.by_gs: dict[tuple, object] = {}  # (grp_nkey, sub_canon) -> value
        self.adj: list[tuple[str, float]] = []  # [(類型顯示名, 金額), ...] 非零
        self.seed: dict[str, object] = {}   # 潛在調整合計/調整後投資金額/跨司工作組確認投資金額
        self.seq: str = ""                  # 項目序號名稱（join key，如 "76-吸引外國客源…"）
        self.override: dict[str, object] = {}  # source_2 覆蓋：canon_sub -> value（為準，蓋過計算值）

    def has_adj(self) -> bool:
        """有冇實際調整：有非零列舉類型、或 潛在調整合計 非零。"""
        if self.adj:
            return True
        n = num(self.seed.get(nkey("潛在調整合計")))
        return n is not None and abs(n) > 1e-9


def merged_val(ws, c: int, r0: int, r1: int):
    for r in range(r0, r1 + 1):
        v = ws.cell(r, c).value
        if _s(v) != "":
            return v
    return None


def project_spans(ws, anchor: int, data0: int, maxrow: int) -> list[tuple[int, int]]:
    """用 anchor 欄垂直合併定項目邊界；未合併但有值 = 單行項目。"""
    spans = []
    covered = set()
    for m in ws.merged_cells.ranges:
        if m.min_col <= anchor <= m.max_col and m.min_row >= data0 and m.max_row > m.min_row:
            spans.append((m.min_row, m.max_row))
            covered.update(range(m.min_row, m.max_row + 1))
    for r in range(data0, maxrow + 1):
        if r in covered:
            continue
        if _s(ws.cell(r, anchor).value) != "":
            spans.append((r, r))
    return sorted(spans)


def extract(ws, log) -> tuple[list[Project], int, int, dict[int, str], int]:
    maxrow, maxcol = ws.max_row or 0, ws.max_column or 0
    subrow = detect_sub_row(ws, maxrow, maxcol)
    grow = max(1, subrow - 1)
    gm = group_map(ws, grow, maxcol)
    anchor = find_anchor_col(ws, subrow, maxcol)
    if anchor is None:
        return [], subrow, 0, gm, maxcol
    # 每欄嘅 (group, sub)
    col_gs: dict[int, tuple] = {}
    for c in range(1, maxcol + 1):
        col_gs[c] = (gm.get(c, ""), canon_sub(ws.cell(subrow, c).value))
    left_cols = [c for c in range(1, anchor)]         # 左半 = anchor 之前
    right_cols = [c for c in range(anchor, maxcol + 1)]
    data0 = subrow + 1
    spans = project_spans(ws, anchor, data0, maxrow)
    projs: list[Project] = []
    for r0, r1 in spans:
        p = Project(r0, r1)
        for r in range(r0, r1 + 1):
            p.left.append([ws.cell(r, c).value for c in left_cols])
        # 序號名稱 join key = 左半第一行第一個似 "數字-名" 或次欄
        first = [ _s(x) for x in (p.left[0] if p.left else []) ]
        p.seq = next((x for x in first if re.match(r"^\d+\s*[-–]", x)), (first[1] if len(first) > 1 else (first[0] if first else "")))
        for c in right_cols:
            grp, sub = col_gs[c]
            val = merged_val(ws, c, r0, r1)
            # 潛在調整組：拆調整類型 / seed
            if sub in {nkey("潛在調整合計"), nkey("調整後投資金額"), nkey("跨司工作組確認投資金額"), nkey("申報投資金額")}:
                p.seed[sub] = val
            elif _is_adj_type(grp, sub, ws.cell(subrow, c).value):
                n = num(val)
                if n is not None and abs(n) > 1e-9:
                    p.adj.append((_s(ws.cell(subrow, c).value), n))
            else:
                if val is not None:
                    p.by_gs[(grp, sub)] = val
        projs.append(p)
    return projs, subrow, anchor, gm, maxcol


def _is_adj_type(grp: str, sub: str, raw: str) -> bool:
    """只認「投資金額的潛在調整事項及調整後金額」拆解組嘅調整類型欄。
    要排除「與承批公司溝通潛在調整事項…」溝通組（組名一樣含『潛在調整』，
    但佢啲欄係文字工作流，唔係金額拆解）。"""
    if "潛在調整" not in grp or "溝通" in grp:
        return False
    return sub not in ADJ_NON_TYPE


def build_enum(p: Project) -> str:
    lines = []
    for i, (typ, amt) in enumerate(p.adj, 1):
        lines.append(f"{i}、{typ}：{fmt_amt(abs(amt))}萬澳門元")
    return "\n".join(lines)


# ── 解析目標欄值 ──────────────────────────────────────────────────────────
def cell_value(tpl, c: int, p: Project):
    """一個表頭右半欄嘅最終值：source_2 覆蓋為準，否則按 rule 計算。"""
    _grp, sub = tpl.col_gs[c]
    if sub in p.override and _s(p.override[sub]) != "":
        return p.override[sub]
    return resolve(tpl.rules[c], p)


def resolve(rule, p: Project):
    kind = rule[0]
    if kind == "blank":
        return None
    if kind == "copy_any":
        sub = canon_sub(rule[1])
        for (g, s), v in p.by_gs.items():
            if s == sub:
                return v
        return None
    if kind == "copy":
        want_sub = canon_sub(rule[2])          # 分組用前綴寬鬆 match、子表頭正規化 match
        for (g, s), v in p.by_gs.items():
            if s == want_sub and _grp_match(g, rule[1]):
                return v
        return None
    if kind == "enum":
        if GATE_NOADJ and not p.has_adj():
            return None
        return build_enum(p) or None
    if kind == "abs_total":
        if GATE_NOADJ and not p.has_adj():
            return None
        n = num(p.seed.get(nkey("潛在調整合計")))
        return abs(n) if n is not None else None
    if kind == "seed":
        key = nkey(rule[1])
        if (GATE_NOADJ and not p.has_adj()
                and key in (nkey("潛在調整合計"), nkey("調整後投資金額"))):
            return None                      # 冇調整 → 建議調整/建議調整後金額 留空（唔填 0）
        return p.seed.get(key)
    return None


def _grp_match(got: str, want: str) -> bool:
    """分組名寬鬆 match：want 係短前綴（如 GT），got 係檔內全名 nkey。"""
    w = nkey(want)
    return w in got or got in w or got.startswith(w[:6])


# ── 表頭 template（header rows 1..subrow）───────────────────────────────────
class Template:
    def __init__(self, hf: Path, log):
        wb = load_wb(hf)
        ws = wb.active
        self.maxcol = ws.max_column or 0
        self.subrow = detect_sub_row(ws, ws.max_row or 0, self.maxcol)
        self.grow = max(1, self.subrow - 1)
        self.gm = group_map(ws, self.grow, self.maxcol)
        self.anchor = find_anchor_col(ws, self.subrow, self.maxcol)
        # 逐欄 (grp, sub) + 每欄嘅 rule（跟 TARGET_SCHEMA 對）
        self.col_gs = {c: (self.gm.get(c, ""), canon_sub(ws.cell(self.subrow, c).value))
                       for c in range(1, self.maxcol + 1)}
        self.rules = self._attach_rules(ws, log)
        # header 區塊（值+style）、col widths、header 內 merges
        self.hdr_cells = {}
        for r in range(1, self.subrow + 1):
            for c in range(1, self.maxcol + 1):
                cell = ws.cell(r, c)
                self.hdr_cells[(r, c)] = (cell.value, copy(cell._style), cell.number_format)
        self.widths = {get_column_letter(c): (ws.column_dimensions[get_column_letter(c)].width)
                       for c in range(1, self.maxcol + 1)}
        self.hdr_merges = [str(m) for m in ws.merged_cells.ranges if m.max_row <= self.subrow]
        self.data_style = copy(ws.cell(self.subrow, self.maxcol)._style)
        wb.close()

    def _attach_rules(self, ws, log):
        """按 (grp,sub) 幫每個表頭欄搭 TARGET_SCHEMA rule；左半 = LEFT。"""
        rules = {}
        left_cols = [c for c in range(1, (self.anchor or 1))]
        for c in left_cols:
            rules[c] = ("__left__", c)
        # 建 schema lookup by (grp_match, sub)
        for c in range(self.anchor or 1, self.maxcol + 1):
            grp, sub = self.col_gs[c]
            if not sub:                      # row6 空（縱向合併表頭）→ 用 row5 文字做 sub
                sub = canon_sub(ws.cell(self.grow, c).value)
            match = None
            for row in TARGET_SCHEMA:
                if canon_sub(row["sub"]) == sub and (row["grp"] == "" or _grp_match(grp, row["grp"])):
                    match = row["rule"]
                    break
            if match is None:
                match = ("blank",)
                log(f"      ⚠ 表頭欄 {get_column_letter(c)} ({grp}|{sub}) 冇對應 rule → 留空")
            rules[c] = match
        return rules


# ── 寫一個對齊 sheet ───────────────────────────────────────────────────────
def write_sheet(ws_out, tpl: Template, projs: list[Project], src_left_anchor: int, log):
    # header 區
    for (r, c), (v, st, nf) in tpl.hdr_cells.items():
        cell = ws_out.cell(r, c, v)
        cell._style = copy(st)
        cell.number_format = nf
    for col, w in tpl.widths.items():
        if w:
            ws_out.column_dimensions[col].width = w
    for m in tpl.hdr_merges:
        ws_out.merge_cells(m)

    tgt_left = [c for c in range(1, (tpl.anchor or 1))]     # 表頭左半欄
    row = tpl.subrow + 1
    for p in projs:
        span = p.r1 - p.r0 + 1
        # 左半：右對齊 positional（兩邊都以 anchor 前一欄為尾）
        for i in range(span):
            src_row = p.left[i] if i < len(p.left) else []
            n = min(len(tgt_left), len(src_row))
            for k in range(1, n + 1):
                v = src_row[-k]
                if v is not None:
                    cell = ws_out.cell(row + i, tgt_left[-k], v)
                    cell._style = copy(tpl.data_style)
                    cell.alignment = _WRAP_TOP
        # 右半：每欄取值（source_2 覆蓋優先）、span 內合併
        for c in range(tpl.anchor or 1, tpl.maxcol + 1):
            val = cell_value(tpl, c, p)
            cell = ws_out.cell(row, c, val)
            cell._style = copy(tpl.data_style)
            cell.alignment = _WRAP_TOP
            cell.border = _BORDER
            if span > 1:
                ws_out.merge_cells(start_row=row, end_row=row + span - 1,
                                   start_column=c, end_column=c)
        row += span
    return row


def light_copy_sheet(ws_out, ws_src):
    """附件/非標準 sheet → 原樣 best-effort 複製（值+合併+欄寬）。"""
    for r in range(1, (ws_src.max_row or 0) + 1):
        for c in range(1, (ws_src.max_column or 0) + 1):
            v = ws_src.cell(r, c).value
            if v is not None:
                ws_out.cell(r, c, v)
    for m in ws_src.merged_cells.ranges:
        ws_out.merge_cells(str(m))
    for c in range(1, (ws_src.max_column or 0) + 1):
        L = get_column_letter(c)
        w = ws_src.column_dimensions[L].width
        if w:
            ws_out.column_dimensions[L].width = w


# ── source_2 overlay ──────────────────────────────────────────────────────
def find_overlay_file(root: Path, scope: str, company: str) -> Path | None:
    base = root / "source_2"
    if not base.exists():
        return None
    cands = []
    for p in base.rglob("*"):
        if p.is_dir() or p.name.startswith("~$") or p.suffix.lower() not in EXCEL_EXT:
            continue
        rel = p.relative_to(base).as_posix()
        if scope and scope in rel and company and company in rel:
            cands.append(p)
    return sorted(cands, key=lambda x: len(x.as_posix()))[0] if cands else None


def build_overlay(path: Path, log) -> dict[str, dict]:
    """source_2 per-範疇檔 → {項目序號: {sub_canon: value}}（掃全部 sheet，只攞 OVERLAY_SUBS）。"""
    try:
        wb = load_wb(path)
    except Exception as e:
        log(f"      ⚠ overlay 開唔到 {path.name}: {e}")
        return {}
    out: dict[str, dict] = {}
    log(f"      overlay 檔 sheets: {wb.sheetnames}")
    for sn in wb.sheetnames:
        ws = wb[sn]
        projs, subrow, anchor, gm, maxcol = extract(ws, log)
        if anchor is None or not projs:
            log(f"      overlay sheet {sn!r}: anchor={anchor} 項目={len(projs)} → 跳過")
            continue
        found = 0
        for p in projs:
            d = {}
            for (g, s), v in p.by_gs.items():
                if s not in OVERLAY_SUBS or _s(v) == "":
                    continue
                # 該關注事項涉及調整金額：source_2 呢格若係文字列舉，唔好蓋我哋算好嘅數字
                if s == nkey("該關注事項涉及調整金額") and num(v) is None:
                    continue
                d[s] = v
            if d:
                out[_seqkey(p.seq)] = d
                found += 1
        log(f"      overlay sheet {sn!r}: {len(projs)} 項目, {found} 個有覆蓋值 "
            f"(序號: {[_seqkey(p.seq) for p in projs]})")
    wb.close()
    return out


def _seqkey(seq: str) -> str:
    m = re.match(r"^\s*(\d+)", _s(seq))
    return m.group(1) if m else nkey(seq)


def apply_overlay(projs, overlay: dict, tpl: Template, log):
    hits = 0
    for p in projs:
        d = overlay.get(_seqkey(p.seq))
        if not d:
            continue
        hits += 1
        for s, v in d.items():
            p.override[s] = v               # 為準，寫入時蓋過計算值
    log(f"      overlay 命中 {hits}/{len(projs)} 個項目")


# ── preview ───────────────────────────────────────────────────────────────
def preview_sheet(sn, projs, tpl: Template, log):
    log(f"\n  ── sheet {sn!r}：{len(projs)} 個項目 ──")
    for p in projs[:6]:
        log(f"\n    ▸ 項目「{p.seq}」 rows {p.r0}-{p.r1}（{p.r1-p.r0+1}行）")
        log(f"       左半首行: {[_s(x) for x in (p.left[0] if p.left else [])][:6]}")
        if p.adj:
            log(f"       調整類型(非零): " + "; ".join(f"{t}={fmt_amt(a)}" for t, a in p.adj))
        log(f"       seed: 潛在調整合計={_s(p.seed.get(nkey('潛在調整合計')))} "
            f"調整後={_s(p.seed.get(nkey('調整後投資金額')))} "
            f"跨司確認={_s(p.seed.get(nkey('跨司工作組確認投資金額')))}")
        # 表頭尾段會寫乜
        for c in range(tpl.anchor or 1, tpl.maxcol + 1):
            g, s = tpl.col_gs[c]
            if s in {nkey("調整原因"), nkey("該關注事項涉及調整金額"), nkey("建議調整金額"),
                     nkey("建議調整後金額"), nkey("建議接納之調整後金額"),
                     nkey("需溝通關注事項"), nkey("是否需進一步與跨司工作組溝通"),
                     nkey("畢馬威關注事項"), nkey("承批公司的反饋意見"),
                     nkey("跨司工作組主責部門針對該關注事項已給的反饋意見"),
                     nkey("KPMG需與跨司工作組進一步確認的問題"),
                     nkey("跨司工作組最新反饋意見")}:
                v = cell_value(tpl, c, p)
                if _s(v):
                    tag = " [source_2]" if s in p.override and _s(p.override[s]) != "" else ""
                    disp = _s(v).replace("\n", " ⏎ ")
                    log(f"       → 表頭[{get_column_letter(c)} {s}]{tag} = {disp[:120]}")
    if len(projs) > 6:
        log(f"    …（另 {len(projs)-6} 個項目略）")


# ── 每檔處理 ───────────────────────────────────────────────────────────────
SCOPE_HINTS = ["旅遊局", "博監局", "經濟局", "文化局", "體育局", "郵電局", "其他範疇", "治安警"]


def infer_scope_company(rel: str) -> tuple[str, str]:
    parts = Path(rel).parts
    scope = next((s for s in SCOPE_HINTS if any(s in pp for pp in parts)), "")
    m = re.search(r"([A-Za-z]{2,})[-\-]", Path(rel).name)
    company = m.group(1) if m else ""
    return scope, company


def is_attachment_sheet(name: str) -> bool:
    return any(k in name for k in ("附件", "承批公司附件", "attach", "Attach"))


def process_file(root: Path, rel: str, tpl: Template, out_dir: Path,
                 preview: bool, with_overlay: bool, log):
    src = root / rel
    log(f"\n{'='*74}\n# {rel}")
    if not src.exists():
        log("  ✗ 唔存在"); return
    wb = load_wb(src)
    scope, company = infer_scope_company(rel)
    overlay = {}
    if with_overlay:
        ofile = find_overlay_file(root, scope, company)
        if ofile:
            log(f"  overlay 檔: {ofile.relative_to(root).as_posix()}")
            overlay = build_overlay(ofile, log)
            if not overlay:
                log("  ⚠ overlay 抽唔到任何覆蓋值（下面 base 為準）")
        else:
            log(f"  （冇 source_2 per-範疇檔 match scope={scope} company={company}）")

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)
    for sn in wb.sheetnames:
        ws = wb[sn]
        if is_attachment_sheet(sn):
            log(f"  ── sheet {sn!r}：附件 → 原樣照抄（best-effort）")
            if not preview:
                light_copy_sheet(out_wb.create_sheet(sn), ws)
            continue
        projs, subrow, anchor, gm, maxcol = extract(ws, log)
        if anchor is None or not projs:
            log(f"  ── sheet {sn!r}：揾唔到「是否該司局範疇」anchor 或冇項目 → 原樣照抄")
            if not preview:
                light_copy_sheet(out_wb.create_sheet(sn), ws)
            continue
        if overlay:
            apply_overlay(projs, overlay, tpl, log)
        if preview:
            preview_sheet(sn, projs, tpl, log)
        else:
            write_sheet(out_wb.create_sheet(sn[:31]), tpl, projs, anchor, log)
    wb.close()
    if preview:
        return
    out_path = out_dir / rel
    out_path.parent.mkdir(parents=True, exist_ok=True)
    if len(out_wb.sheetnames) == 0:
        out_wb.create_sheet("空")
    out_wb.save(out_path)
    log(f"  ✓ 寫入 {out_path.relative_to(root).as_posix()}")


def copy_asis(root: Path, rel: str, out_dir: Path, log):
    src, dst = root / rel, out_dir / rel
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    log(f"  ⧉ 照抄 {rel}")


def iter_source1(root: Path):
    base = root / "source_1"
    for p in sorted(base.rglob("*")):
        if p.is_dir() or p.name.startswith("~$"):
            continue
        if p.suffix.lower() in EXCEL_EXT:
            yield p.relative_to(root).as_posix()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="ad-hoc/workspace")
    ap.add_argument("--header-file", default="表頭.xlsx")
    ap.add_argument("--only", help="淨處理呢一個相對 root 嘅檔（驗證用）")
    ap.add_argument("--all", action="store_true", help="批 source_1 全部檔")
    ap.add_argument("--preview", action="store_true", help="只吐抽取+映射文字，唔寫 xlsx")
    ap.add_argument("--with-overlay", action="store_true", help="套 source_2 per-範疇覆蓋")
    ap.add_argument("--fill-zero-adj", action="store_true",
                    help="冇調整都照抄 0（預設留空）")
    ap.add_argument("--out", default="_aligned", help="輸出子資料夾（root 底下）")
    a = ap.parse_args()
    global GATE_NOADJ
    GATE_NOADJ = not a.fill_zero_adj
    root = Path(a.root)
    out_dir = root / a.out
    lines = []

    def log(s=""):
        lines.append(s); print(s)

    hf = root / a.header_file
    if not hf.exists():
        raise SystemExit(f"✗ 冇表頭: {hf}")
    tpl = Template(hf, log)
    log(f"# 表頭 template：{tpl.maxcol} 欄, 子表頭 row{tpl.subrow}, anchor 欄 "
        f"{get_column_letter(tpl.anchor) if tpl.anchor else '?'}")

    if a.only:
        targets = [a.only]
    elif a.all:
        targets = list(iter_source1(root))
    else:
        raise SystemExit("俾 --only <rel> 或 --all")

    for rel in targets:
        parts = Path(rel).parts
        if any(pp.lower() == "ss" for pp in parts):
            if not a.preview:
                copy_asis(root, rel, out_dir, log)
            continue
        try:
            process_file(root, rel, tpl, out_dir, a.preview, a.with_overlay, log)
        except Exception as e:
            import traceback
            log(f"  ✗ 出錯: {type(e).__name__}: {e}")
            log(traceback.format_exc())

    rp = root / ("_align_preview.txt" if a.preview else "_align_log.txt")
    rp.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n✓ log 寫入 {rp.resolve()}（貼返嚟俾我）")


if __name__ == "__main__":
    main()
