"""Table 2 pipeline logic.

Two steps, no caching — every kedro run is a fresh build:

  table2_extract:  parses data/table_2/2025_*_*.xlsx (vertical 表二 forms)
                   → data/table_2/_extracted.xlsx
                   Each row = one NG section (anchor "1./2./..."), with
                   plan_* (left) vs actual_* (right) paired columns and
                   computed 達成率 + pct_of_company_category / pct_of_industry.

  table2_analyze:  reads _extracted.xlsx → LLM per project →
                   data/table_2/_analyzed.xlsx (projects_analyzed + flagged sheets).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

from kpi.lib.conf import ROOT, load_master
from kpi.lib.io_setup import force_unbuffered_io
from kpi.lib.llm import LLMClient

force_unbuffered_io()


SRC_DIR = ROOT / "data" / "table_2"
EXTRACTED = SRC_DIR / "_extracted.xlsx"
ANALYZED = SRC_DIR / "_analyzed.xlsx"
LLM_CACHE = SRC_DIR / "_llm_cache.jsonl"   # cleared at the start of analyze step
SHEET_CONFIG = ROOT / "conf" / "table_2" / "parameters.yml"


def _load_sheet_config() -> dict[str, list[dict]]:
    """Load per-file sheet selection from conf/table_2/parameters.yml.
    Returns {filename: [{sheet, data_year, is_subsequent_year}]}, or {} if missing."""
    if not SHEET_CONFIG.exists():
        return {}
    import yaml
    with SHEET_CONFIG.open(encoding="utf-8") as f:
        cfg = yaml.safe_load(f) or {}
    raw = cfg.get("files") or {}
    out: dict[str, list[dict]] = {}
    for fname, entries in raw.items():
        if not isinstance(entries, list):
            continue
        out[fname] = [
            {
                "sheet": e.get("sheet", ""),
                "data_year": e.get("data_year", 2025),
                "is_subsequent_year": bool(e.get("is_subsequent_year", False)),
            }
            for e in entries
            if isinstance(e, dict) and e.get("sheet")
        ]
    return out


# ── Schema ────────────────────────────────────────────────────────────────────
FNAME_RE = re.compile(r"^2025_([A-Za-z]+)_(Gaming|Non-gaming)\.xlsx$")
ANCHOR_RE = re.compile(r"^\s*(\d+)\.\s*$")
# Read whole sheet — 表二 are small (a few hundred rows max). The previous
# 400-row cap silently truncated long sheets like Melco Non-gaming.
MAX_ROWS = None

# Section header row markers — appear between NG categories (e.g. "5. B5文化藝術"
# → "6. B6健康養生"). The header cell embeds the section subtotal of all projects
# within that section, on plan and actual sides.
SECTION_HEADER_KEYWORDS = (
    "性質範疇", "各項目之加總", "投資項目性質 / 金額",
)
# Matches "6. B6健康養生" / "1. B1.1吸引外國客源" / "1. A1博彩設施"
SECTION_CODE_RE = re.compile(
    r"^\s*(\d+)\.\s*([AB]\d+(?:\.\d+)?)\s*(.+?)\s*$"
)
# Matches "設施150.0萬澳門元" / "活動1,020.0萬澳門元" / "總投資1,170.0萬澳門元"
SECTION_SUBTOTAL_RE = re.compile(
    r"(設施|活動|總投資)[\s:：]*([\d,]+\.?\d*)\s*萬澳門元"
)
SUBTOTAL_KEY_MAP = {"設施": "capex", "活動": "opex", "總投資": "total"}

FIELD_MAP = {
    "company_name":            ["承批公司"],
    "project_id_name":         ["投資項目序號及名稱"],
    "not_started_explanation": ["未開展情況説明", "未開展情況說明"],
    "implementation_time":     ["實施時間"],
    "amt_capex":               ["該年度「設施建設」投資金額（萬澳門元）"],
    "amt_opex":                ["該年度「活動舉辦」投資金額（萬澳門元）"],
    "amt_total":               ["該年度投資金額（萬澳門元）"],
    "cumulative_capex":        ["批給期內總「設施建設」投資金額（萬澳門元）"],
    "cumulative_opex":         ["批給期內總「活動舉辦」投資金額（萬澳門元）"],
    "cumulative_total":        ["批給期內總投資金額（萬澳門元）"],
    "location":                ["地點、空間"],
    "kpi_indicators":          ["每年可量化的指標及經濟效益估算"],
    "content_brief":           ["投資項目內容、實施計劃的重點簡介"],
    "content_ref_doc":         ["投資項目內容、實施計劃詳情之參考文件"],
}
LABEL_TO_CANON: dict[str, str] = {}
for canon, variants in FIELD_MAP.items():
    for v in variants:
        LABEL_TO_CANON[v] = canon

NUMERIC_FIELDS = {"amt_capex", "amt_opex", "amt_total",
                  "cumulative_capex", "cumulative_opex", "cumulative_total"}
NULL_TOKENS = {"-", "—", "不適用", "暫未有相關資料", "N/A", "n/a", ""}


# ── Extract helpers ───────────────────────────────────────────────────────────
def _cell(row, col: int) -> str:
    if col < 0 or col >= len(row):
        return ""
    v = row.iloc[col]
    if pd.isna(v):
        return ""
    return str(v).strip()


def _strip_label(s: str) -> str:
    return s.rstrip("：:").strip()


def _to_num(s: str) -> float | None:
    if not s or s in NULL_TOKENS:
        return None
    s = s.replace(",", "").replace(" ", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _detect_anchor_col(df: pd.DataFrame) -> int | None:
    for col in range(min(4, len(df.columns))):
        for _, row in df.head(60).iterrows():
            if ANCHOR_RE.match(_cell(row, col)):
                return col
    return None


def _parse_id_name(raw: str) -> tuple[str, str]:
    if not raw:
        return "", ""
    m = re.match(r"^\s*(.+?)\s*[-–—]\s*(.+)$", raw, re.DOTALL)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return raw.strip(), ""


def _parse_section_subtotals(text: str) -> dict[str, float]:
    """From a section-header cell like '... 6. B6健康養生 設施150 活動1020 總投資1170 ...',
    extract {capex: 150, opex: 1020, total: 1170}."""
    out: dict[str, float] = {}
    for m in SECTION_SUBTOTAL_RE.finditer(text):
        key = SUBTOTAL_KEY_MAP.get(m.group(1))
        val = m.group(2).replace(",", "").strip()
        try:
            out[key] = float(val)
        except (ValueError, TypeError):
            pass
    return out


def _detect_sections(df: pd.DataFrame, anchor_col: int) -> list[dict]:
    """Find section-header rows. Returns [{row, ng_section_seq, ng_section_code, ng_section_name,
    plan_*, actual_*}].

    Section headers appear between project blocks and contain a marker phrase
    like '性質範疇' / '各項目之加總'. Multi-line cells embed both the section name
    (e.g. '6. B6健康養生') and the subtotal amounts.
    """
    plan_col = anchor_col + 2
    actual_col = anchor_col + 4
    sections: list[dict] = []
    for i in range(len(df)):
        row = df.iloc[i]
        cells = [_cell(row, c) for c in range(len(row))]
        joined = " ".join(cells)
        if not any(kw in joined for kw in SECTION_HEADER_KEYWORDS):
            continue
        # Find the section identifier across all cells
        seq, code, name = None, "", ""
        for cell in cells:
            for line in cell.split("\n"):
                m = SECTION_CODE_RE.match(line.strip())
                if m:
                    seq = int(m.group(1))
                    code = m.group(2).strip()
                    name = m.group(3).strip()
                    break
            if code:
                break
        if not code:
            continue
        plan_text = _cell(row, plan_col)
        actual_text = _cell(row, actual_col)
        plan_amts = _parse_section_subtotals(plan_text)
        actual_amts = _parse_section_subtotals(actual_text)
        sections.append({
            "row": i,
            "ng_section_seq": seq,
            "ng_section_code": code,
            "ng_section_name": name,
            "plan_section_capex": plan_amts.get("capex"),
            "plan_section_opex": plan_amts.get("opex"),
            "plan_section_total": plan_amts.get("total"),
            "actual_section_capex": actual_amts.get("capex"),
            "actual_section_opex": actual_amts.get("opex"),
            "actual_section_total": actual_amts.get("total"),
        })
    return sections


def _section_for_row(sections: list[dict], row_idx: int) -> dict | None:
    """The section that owns project block starting at row_idx (closest preceding header)."""
    chosen = None
    for s in sections:
        if s["row"] <= row_idx:
            chosen = s
        else:
            break
    return chosen


def _section_from_pid(pid: str) -> str:
    """Fallback: extract NG section code from project_id prefix (Galaxy/Wynn style).
    E.g. 'B6.1' → 'B6', 'A1.2' → 'A1'. Returns '' if no match."""
    m = re.match(r"^([AB]\d+)", (pid or "").strip())
    return m.group(1) if m else ""


def _parse_sheet(df: pd.DataFrame, company: str, category: str,
                 file_name: str, sheet_name: str,
                 data_year: int = 2025, is_subsequent_year: bool = False) -> list[dict]:
    anchor_col = _detect_anchor_col(df)
    if anchor_col is None:
        return []
    label_col = anchor_col + 1
    plan_col = anchor_col + 2
    actual_col = anchor_col + 4
    n = len(df)

    anchors: list[tuple[int, int]] = []
    for i in range(n):
        m = ANCHOR_RE.match(_cell(df.iloc[i], anchor_col))
        if m:
            anchors.append((i, int(m.group(1))))
    if not anchors:
        return []

    sections = _detect_sections(df, anchor_col)

    records: list[dict] = []
    for j, (start, anchor_no) in enumerate(anchors):
        end = anchors[j + 1][0] if j + 1 < len(anchors) else n
        plan_data: dict[str, str] = {}
        actual_data: dict[str, str] = {}
        for r in range(start, end):
            row = df.iloc[r]
            label = _strip_label(_cell(row, label_col))
            canon = LABEL_TO_CANON.get(label)
            if not canon:
                continue
            v_plan = _cell(row, plan_col)
            v_actual = _cell(row, actual_col)
            if v_plan:
                prev = plan_data.get(canon, "")
                plan_data[canon] = (prev + "\n" + v_plan).strip() if prev else v_plan
            if v_actual:
                prev = actual_data.get(canon, "")
                actual_data[canon] = (prev + "\n" + v_actual).strip() if prev else v_actual

        if not (plan_data.get("project_id_name") or actual_data.get("project_id_name")):
            continue

        plan_pid, plan_pname = _parse_id_name(plan_data.get("project_id_name", ""))
        act_pid, act_pname = _parse_id_name(actual_data.get("project_id_name", ""))

        # Attach section info (from detected boundary header, fall back to project_id prefix)
        sec = _section_for_row(sections, start) or {}
        pid_section = _section_from_pid(plan_pid or act_pid)
        ng_section_code = sec.get("ng_section_code") or pid_section
        ng_section_name = sec.get("ng_section_name", "")
        ng_section_seq = sec.get("ng_section_seq")
        # Gaming projects (ng_section_code=0) — auto-fill section_name if blank
        if str(ng_section_code).strip() == "0" and not str(ng_section_name).strip():
            ng_section_name = "博彩項目"

        rec = {
            "file": file_name,
            "sheet": sheet_name,
            "company": company,
            "category": category,
            "data_year": data_year,
            "is_subsequent_year": is_subsequent_year,
            "ng_section_seq": ng_section_seq,
            "ng_section_code": ng_section_code,
            "ng_section_name": ng_section_name,
            "anchor_no": anchor_no,                   # within-section sequence ("1./2./3.")
            "anchor_row_1based": start + 1,
            "project_id": plan_pid or act_pid,
            "plan_project_name": plan_pname,
            "actual_project_name": act_pname if act_pname and act_pname != plan_pname else "",
            "names_differ": (plan_pid != act_pid or plan_pname != act_pname)
                            and bool(plan_pname) and bool(act_pname),
            "plan_implementation_time": plan_data.get("implementation_time", ""),
            "actual_implementation_time": actual_data.get("implementation_time", ""),
            # Section-level subtotals from the boundary header (for filter / drill-down)
            "ng_section_plan_total": sec.get("plan_section_total"),
            "ng_section_plan_capex": sec.get("plan_section_capex"),
            "ng_section_plan_opex": sec.get("plan_section_opex"),
            "ng_section_actual_total": sec.get("actual_section_total"),
            "ng_section_actual_capex": sec.get("actual_section_capex"),
            "ng_section_actual_opex": sec.get("actual_section_opex"),
        }
        for canon in FIELD_MAP:
            if canon in ("project_id_name", "company_name", "implementation_time"):
                continue
            plan_v = plan_data.get(canon, "")
            act_v = actual_data.get(canon, "")
            if canon in NUMERIC_FIELDS:
                rec[f"plan_{canon}"] = _to_num(plan_v)
                rec[f"actual_{canon}"] = _to_num(act_v)
            else:
                rec[f"plan_{canon}"] = plan_v
                rec[f"actual_{canon}"] = act_v
        for short in ("amt_total", "amt_capex", "amt_opex"):
            p = rec.get(f"plan_{short}")
            a = rec.get(f"actual_{short}")
            rec[f"achievement_{short}"] = round(a / p, 4) if (p and a and p != 0) else None
        rec["plan_company_name"] = plan_data.get("company_name", "")
        rec["actual_company_name"] = actual_data.get("company_name", "")
        records.append(rec)
    return records


# ── Step 1: extract ───────────────────────────────────────────────────────────
def extract_main() -> None:
    if EXTRACTED.exists():
        EXTRACTED.unlink()
    SRC_DIR.mkdir(parents=True, exist_ok=True)

    files = sorted(p for p in SRC_DIR.glob("2025_*_*.xlsx") if not p.name.startswith("_"))
    print(f"Found {len(files)} files in {SRC_DIR}")

    sheet_config = _load_sheet_config()
    if sheet_config:
        print(f"Loaded sheet config: {len(sheet_config)} file entries from "
              f"{SHEET_CONFIG.relative_to(ROOT)}")
    else:
        print(f"No {SHEET_CONFIG.relative_to(ROOT)} found — "
              "falling back to 'all non-表一 sheets'. Copy parameters.yml.template to "
              "parameters.yml in conf/table_2/ to control which sheets are parsed.")

    all_records: list[dict] = []
    for fp in files:
        m = FNAME_RE.match(fp.name)
        if not m:
            print(f"  [skip] {fp.name} — name mismatch")
            continue
        company, category = m.group(1), m.group(2)
        print(f"\n── {fp.name} ──")
        xl = pd.ExcelFile(fp, engine="openpyxl")
        available = set(xl.sheet_names)

        # Determine which sheets to process
        if fp.name in sheet_config:
            sheet_entries = sheet_config[fp.name]
            # Warn on any configured sheet that doesn't exist in the file
            for e in sheet_entries:
                if e["sheet"] not in available:
                    print(f"  WARNING: configured sheet {e['sheet']!r} not in xlsx (skipped)")
        else:
            sheet_entries = [
                {"sheet": s, "data_year": 2025, "is_subsequent_year": False}
                for s in xl.sheet_names if "表一" not in s
            ]

        for entry in sheet_entries:
            sheet = entry["sheet"]
            if sheet not in available:
                continue
            df = pd.read_excel(fp, sheet_name=sheet, header=None, dtype=str,
                               nrows=MAX_ROWS, engine="openpyxl")
            recs = _parse_sheet(
                df, company, category, fp.name, sheet,
                data_year=entry["data_year"],
                is_subsequent_year=entry["is_subsequent_year"],
            )
            tag = " (SY)" if entry["is_subsequent_year"] else ""
            print(f"  [{sheet}]{tag} year={entry['data_year']} → {len(recs)} projects")
            all_records.extend(recs)

    if not all_records:
        print("\nNo records extracted.")
        return

    out_df = pd.DataFrame(all_records)
    grp = out_df.groupby(["company", "category"], dropna=False)
    company_totals = grp["actual_amt_total"].transform("sum")
    out_df["pct_of_company_category"] = (out_df["actual_amt_total"] / company_totals).round(4)
    industry_total = out_df["actual_amt_total"].sum()
    out_df["pct_of_industry"] = (out_df["actual_amt_total"] / industry_total).round(4) if industry_total else None

    with pd.ExcelWriter(EXTRACTED, engine="xlsxwriter") as w:
        out_df.to_excel(w, sheet_name="projects", index=False)
        summary = (
            out_df.groupby(["company", "category"], dropna=False)
            .agg(projects=("project_id", "count"),
                 plan_total=("plan_amt_total", "sum"),
                 actual_total=("actual_amt_total", "sum"))
            .reset_index()
        )
        summary["achievement"] = (summary["actual_total"] / summary["plan_total"]).round(4)
        summary.to_excel(w, sheet_name="_summary", index=False)

    print(f"\nWrote {EXTRACTED.relative_to(ROOT)}  ({len(out_df)} project rows)")
    print("\nSummary by company × category:")
    print(summary.to_string(index=False))


# ── Step 2: LLM analyze ───────────────────────────────────────────────────────
SYSTEM = (
    "你是審計顧問，正在審查澳門博彩承批公司 2025 年度投資計劃 vs 實際執行報告（表二）。\n"
    "我會給你一個項目的計劃敘述（plan_content_brief）與實際執行敘述（actual_content_brief）。\n\n"
    "**唯一任務：判斷實際做嘅嘢有冇本質上超出計劃範圍。**\n\n"
    "判定原則 — 平衡 sensitivity：\n"
    "  • **細化 / 詳述**：plan 寫 A，actual 寫 A1/A2/A3（A 嘅子步驟、具體 supplier、具體場地、\n"
    "    具體 KPI、具體月份活動、具體 sub-project）→ **within_plan**\n"
    "  • **縮減 / 部分完成 / delay / 做少咗** → **within_plan**\n"
    "  • **換對象 / 換 nature**：plan 講 A 類別，actual 突然講 X 類別（不同 nature）→ **exceeds_plan**\n"
    "  • **新增 distinct activity**：plan 冇明確列嘅 standalone item（非 A 嘅自然延伸）→ **exceeds_plan**\n"
    "  • **金額顯著超支**：actual 比 plan 多 30%+（且新增了 plan 冇明確列嘅 sub-item）→ **exceeds_plan**\n\n"
    "判定例子：\n"
    "  ✓ within_plan：plan「2025 演唱會系列」→ actual「Leo Ku / Eric Moo 演唱會」（具體藝人＝細化）\n"
    "  ✓ within_plan：plan「翻新酒店大堂」→ actual「翻新大堂 + lobby bar + 走廊裝修」（同範疇）\n"
    "  ✓ within_plan：plan「3 個 roadshow」→ actual「1 個 roadshow」（少做 / delay）\n"
    "  ✗ exceeds_plan：plan「翻新酒店大堂」→ actual「翻新大堂 + 採購新 slot machine」（新範疇）\n"
    "  ✗ exceeds_plan：plan「演唱會」→ actual「演唱會 + 全新藝術展覽」（新增 distinct activity）\n"
    "  ✗ exceeds_plan：plan「廣告 campaign」→ actual「贊助體育賽事」（換 nature）\n"
    "  ✗ exceeds_plan：plan「翻新 A 區」金額 1000萬 → actual 1500萬 加埋整咗 B 區（金額+50% 新範疇）\n\n"
    "資料不足無法判斷 → unclear。\n\n"
    "**重要輸出要求** — 對於 exceeds_plan 嘅 case：\n"
    "  • 必須喺 `exceeded_items_detail` 列出每件具體做多咗嘅 item，包括：\n"
    "      - item_name：item 嘅具體名（中文）\n"
    "      - planned：plan 入面有冇對應描述（'未提及' 或者 plan 入面相關描述）\n"
    "      - actual：actual 入面講嘅嘢（包括金額/規模 如有）\n"
    "      - estimated_amount：你估計呢 item 嘅金額（萬MOP），如 actual 文本有提及金額就用；冇就寫 '未明示'\n"
    "      - diff_type：'new_item'（plan 完全冇）/ 'amount_overrun'（金額超）/ 'scope_change'（換 nature）\n\n"
    "輸出原則：\n"
    "  • deviation_summary：一兩句中文摘要\n"
    "  • exceeded_items_detail：list of {item_name, planned, actual, estimated_amount, diff_type}\n"
    "  • key_concerns：本質範疇換咗（A → 唔同 nature B）嘅明確矛盾；無則空 array\n\n"
    "回答必須是合法 JSON。"
)

SCHEMA_HINT = (
    "JSON schema:\n"
    "{\n"
    '  "deviation_status":   "within_plan|exceeds_plan|unclear",\n'
    '  "deviation_summary":  "<一兩句中文摘要>",\n'
    '  "exceeded_items_detail": [\n'
    '    {\n'
    '      "item_name":         "<具體 item 中文名>",\n'
    '      "planned":           "<plan 入面有冇對應描述，冇就寫「未提及」>",\n'
    '      "actual":            "<actual 入面講嘅描述，包括金額/規模>",\n'
    '      "estimated_amount":  "<萬MOP 數字 或者「未明示」>",\n'
    '      "diff_type":         "new_item|amount_overrun|scope_change"\n'
    '    }\n'
    '  ],\n'
    '  "key_concerns":       [<本質範疇換咗嘅明確矛盾，無則空 array>]\n'
    "}"
)


def _fmt_num(v) -> str:
    if pd.isna(v) or v is None:
        return "—"
    try:
        return f"{float(v):,.0f}"
    except (ValueError, TypeError):
        return str(v)


def _fmt_text(v, max_len: int = 800) -> str:
    if pd.isna(v) or v is None:
        return "—"
    s = str(v).strip()
    if not s:
        return "—"
    return s[:max_len] + ("…" if len(s) > max_len else "")


def _build_user_prompt(row: dict) -> str:
    lines = [
        f"公司: {row.get('company')}    類別: {row.get('category')}",
        f"NG 範疇: {row.get('ng_section_code') or '—'} {row.get('ng_section_name') or ''}   "
        f"範疇內項目序號: {row.get('anchor_no')}   項目編號: {row.get('project_id')}",
        f"計劃項目名稱: {_fmt_text(row.get('plan_project_name'), 200)}",
    ]
    if row.get("actual_project_name"):
        lines.append(f"實際項目名稱: {_fmt_text(row.get('actual_project_name'), 200)}  "
                     f"(NOTE: 名稱與計劃不同)")
    lines.extend([
        "",
        "── 時間 ──",
        f"計劃實施時間: {row.get('plan_implementation_time') or '—'}",
        f"實際實施時間: {row.get('actual_implementation_time') or '—'}",
        "",
        "── 本年度金額 (萬MOP) ──",
        f"  CAPEX:  計劃={_fmt_num(row.get('plan_amt_capex'))}  "
        f"實際={_fmt_num(row.get('actual_amt_capex'))}  "
        f"達成率={row.get('achievement_amt_capex') or '—'}",
        f"  OPEX:   計劃={_fmt_num(row.get('plan_amt_opex'))}  "
        f"實際={_fmt_num(row.get('actual_amt_opex'))}  "
        f"達成率={row.get('achievement_amt_opex') or '—'}",
        f"  Total:  計劃={_fmt_num(row.get('plan_amt_total'))}  "
        f"實際={_fmt_num(row.get('actual_amt_total'))}  "
        f"達成率={row.get('achievement_amt_total') or '—'}",
        "",
        "── 累計金額 (萬MOP) ──",
        f"  Total:  計劃={_fmt_num(row.get('plan_cumulative_total'))}  "
        f"實際={_fmt_num(row.get('actual_cumulative_total'))}",
        "",
        f"── 所屬 NG 範疇小計 ({row.get('ng_section_code') or '—'} "
        f"{row.get('ng_section_name') or ''}) ──",
        f"  Total:  計劃={_fmt_num(row.get('ng_section_plan_total'))}  "
        f"實際={_fmt_num(row.get('ng_section_actual_total'))}",
        "",
        "── 未開展情況 ──",
        f"計劃側: {_fmt_text(row.get('plan_not_started_explanation'), 300)}",
        f"實際側: {_fmt_text(row.get('actual_not_started_explanation'), 300)}",
        "",
        "── 計劃內容簡介 ──",
        _fmt_text(row.get("plan_content_brief"), 1500),
        "",
        "── 實際執行情況 ──",
        _fmt_text(row.get("actual_content_brief"), 1500),
        "",
        "── KPI 指標 ──",
        f"計劃: {_fmt_text(row.get('plan_kpi_indicators'), 300)}",
        f"實際: {_fmt_text(row.get('actual_kpi_indicators'), 300)}",
        "",
        SCHEMA_HINT,
    ])
    return "\n".join(lines)


def analyze_main() -> None:
    if not EXTRACTED.exists():
        raise FileNotFoundError(f"Run table2_extract first. Missing: {EXTRACTED}")
    # Fresh build: drop old analyzed output + LLM cache so every run is from scratch.
    if ANALYZED.exists():
        ANALYZED.unlink()
    if LLM_CACHE.exists():
        LLM_CACHE.unlink()

    df = pd.read_excel(EXTRACTED, sheet_name="projects")
    print(f"Loaded {len(df)} projects from {EXTRACTED.name}")

    master = load_master()
    # Table 2 sends long-context prompts. Start at 4 concurrent; LLMClient halves it on each
    # retry round (4 → 2 → 1) so we push hard then back off when the endpoint complains.
    llm_cfg = {**(master.get("llm") or {}), "concurrency": 4}
    cfg = {"llm": llm_cfg}
    llm = LLMClient(cfg, cache_path=LLM_CACHE)
    payloads = df.to_dict("records")

    def call_one(row):
        return llm.chat_json(system=SYSTEM, user=_build_user_prompt(row), schema_hint=SCHEMA_HINT)

    results = llm.map_parallel(payloads, call_one, desc="table2 analyze")

    def _get(r, k, default=""):
        if not isinstance(r, dict):
            return default
        v = r.get(k, default)
        return v if v is not None else default

    def _join_list(v) -> str:
        """Render a list of strings as multi-line text for spreadsheet readability."""
        if isinstance(v, list):
            return "\n".join(f"• {str(x).strip()}" for x in v if str(x).strip())
        if v is None or v == "":
            return ""
        return str(v)

    # Deviation columns — project team only cares about actual EXCEEDS plan (scope expansion).
    df["llm_deviation_status"]  = [_get(r, "deviation_status") for r in results]
    df["llm_deviation_summary"] = [_get(r, "deviation_summary") for r in results]
    df["llm_key_concerns"]      = [_join_list(_get(r, "key_concerns", [])) for r in results]

    # Structured exceeded items (NEW) — flatten list of {item_name, planned, actual, ...}
    # into 5 parallel string columns for Excel readability.
    def _items_col(results_list, field):
        out = []
        for r in results_list:
            items = _get(r, "exceeded_items_detail", [])
            if not isinstance(items, list):
                out.append("")
                continue
            vals = []
            for it in items:
                if isinstance(it, dict):
                    v = it.get(field, "")
                    vals.append(str(v).strip())
            out.append(" || ".join(vals))
        return out

    df["llm_exceeded_item_names"]     = _items_col(results, "item_name")
    df["llm_exceeded_item_planned"]   = _items_col(results, "planned")
    df["llm_exceeded_item_actual"]    = _items_col(results, "actual")
    df["llm_exceeded_item_amount"]    = _items_col(results, "estimated_amount")
    df["llm_exceeded_item_diff_type"] = _items_col(results, "diff_type")
    # Backward-compat: keep the old text-summary column as concatenated item_names.
    df["llm_exceeds_items"] = df["llm_exceeded_item_names"]

    df["llm_error"] = [
        r.get("_error", "") if isinstance(r, dict) else "non-dict response"
        for r in results
    ]

    # flagged: actual EXCEEDS plan (scope expansion) OR key_concerns flagged.
    flagged = df[
        df["llm_deviation_status"].isin(["exceeds_plan", "unclear"])
        | df["llm_key_concerns"].fillna("").astype(str).str.len().gt(0)
    ]

    # Lean detail view columns (original English keys; renamed to Chinese for output)
    detail_cols = [
        "company", "category", "ng_section_code", "ng_section_name", "anchor_no", "project_id",
        "plan_project_name", "actual_project_name",
        "plan_amt_total", "actual_amt_total", "achievement_amt_total",
        "llm_deviation_status",
        "llm_deviation_summary",
        "llm_exceeded_item_names",
        "llm_exceeded_item_planned",
        "llm_exceeded_item_actual",
        "llm_exceeded_item_amount",
        "llm_exceeded_item_diff_type",
        "llm_key_concerns",
    ]
    detail_cols = [c for c in detail_cols if c in df.columns]

    # Chinese column-name rename for Excel readability (project team feedback).
    COL_RENAME = {
        "company":               "公司",
        "category":              "類別",
        "ng_section_code":       "NG範疇代碼",
        "ng_section_name":       "NG範疇名稱",
        "anchor_no":             "範疇內項目序號",
        "project_id":            "項目編號",
        "plan_project_name":     "計劃項目名稱",
        "actual_project_name":   "實際項目名稱",
        "plan_amt_total":        "計劃金額_合計",
        "actual_amt_total":      "實際金額_合計",
        "achievement_amt_total": "達成率_合計",
        "plan_amt_capex":        "計劃金額_Capex",
        "plan_amt_opex":         "計劃金額_Opex",
        "actual_amt_capex":      "實際金額_Capex",
        "actual_amt_opex":       "實際金額_Opex",
        "achievement_amt_capex": "達成率_Capex",
        "achievement_amt_opex":  "達成率_Opex",
        "plan_cumulative_total":   "計劃累計金額",
        "actual_cumulative_total": "實際累計金額",
        "ng_section_plan_total":   "範疇小計_計劃",
        "ng_section_actual_total": "範疇小計_實際",
        "plan_implementation_time":   "計劃實施時間",
        "actual_implementation_time": "實際實施時間",
        "plan_not_started_explanation":   "計劃側_未開展說明",
        "actual_not_started_explanation": "實際側_未開展說明",
        "plan_content_brief":   "計劃內容簡介",
        "actual_content_brief": "實際執行情況",
        "plan_kpi_indicators":   "計劃KPI",
        "actual_kpi_indicators": "實際KPI",
        "llm_deviation_status":          "LLM_偏離狀態",
        "llm_deviation_summary":         "LLM_偏離摘要",
        "llm_exceeded_item_names":       "LLM_超出項目_名稱",
        "llm_exceeded_item_planned":     "LLM_超出項目_計劃描述",
        "llm_exceeded_item_actual":      "LLM_超出項目_實際描述",
        "llm_exceeded_item_amount":      "LLM_超出項目_估算金額(萬MOP)",
        "llm_exceeded_item_diff_type":   "LLM_超出項目_差異類型",
        "llm_exceeds_items":             "LLM_超出計劃事項(舊欄)",
        "llm_key_concerns":              "LLM_重要矛盾",
        "llm_error":                     "LLM_錯誤",
    }

    with pd.ExcelWriter(ANALYZED, engine="xlsxwriter") as w:
        df_out = df.rename(columns=COL_RENAME)
        df_out.to_excel(w, sheet_name="projects_analyzed", index=False)

        flagged_out = flagged.rename(columns=COL_RENAME)
        flagged_out.to_excel(w, sheet_name="flagged", index=False)

        detail_out = df[detail_cols].rename(columns=COL_RENAME)
        detail_out.to_excel(w, sheet_name="project_details", index=False)

    print(f"\nWrote {ANALYZED.relative_to(ROOT)}")
    print(f"  flagged rows (exceeds_plan/unclear OR key_concerns non-empty): {len(flagged)}")
