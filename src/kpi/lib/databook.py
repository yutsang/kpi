"""Databook 註釋引擎 — 由報告表格生成 headline + 右側 commentary（notes）。

Table-agnostic：入任何一張表（DataFrame）＋ context，出結構化 notes，grounded 喺表內數字。
之後 MGM 25 的實際表格直接 plug 入 `generate_notes`。

設計原則（v0 草稿，見到 MGM 25 databook 風格再收緊 prompt）：
  - 每句都要 grounded 喺表內數字，可引用具體金額/百分比；**唔准作表外數據**。
  - 聚焦重大變動：最大額 / 最大佔比 / 調整前→後變化 / NG·類別集中度。
  - 輸出 JSON：{headline, commentary[2-4], flags[]}。
  - 語言預設繁體中文（audit tone），可切 English。
  - 數字單位由 caller 明示（本 project 一般係「萬」= 1萬 MOP）。
"""
from __future__ import annotations

import json
from typing import Any

import pandas as pd

from .workbench import Workbench

NOTE_SYSTEM_ZH = """你係資深審計/顧問分析員，正為一份「databook」報告表格撰寫旁邊的註釋（commentary）。
對象：澳門博彩承批公司「非博彩（NG）投資 KPI」報告。

嚴格規則：
1. 每一句都必須 grounded 喺我提供嘅表格數字；可引用具體金額同百分比。
2. 絕對唔准作出表格以外嘅數據、假設或前瞻預測（除非我明確叫你估）。
3. 聚焦重大訊息：最大額項目、最大佔比、調整前→調整後嘅變化、NG 或類別集中度。
4. 若見到數據異常（例如調整後為負、佔比 >100% 或 <0、合計對唔上），放入 flags。
5. 語氣專業精煉，似寫俾管理層/董事會睇嘅 report note。

只輸出 JSON（唔好有其他文字），格式：
{
  "headline": "一句標題（≤30字，點出最重要訊息）",
  "commentary": ["要點1（含數字）", "要點2", "…（2-4 條）"],
  "flags": ["數據異常或需覆核（冇就空陣列）"]
}"""

NOTE_SYSTEM_EN = """You are a senior audit/advisory analyst writing the commentary next to a
databook report table for a Macau gaming concessionaire's non-gaming (NG) investment KPI.

Strict rules:
1. Every sentence must be grounded in the numbers I provide; cite specific amounts / percentages.
2. Never invent data, assumptions, or forward-looking statements beyond the table (unless asked).
3. Focus on material messages: largest items, largest shares, pre→post-adjustment movements, NG/category concentration.
4. Put data anomalies (negative post-adjustment, share >100% or <0, totals not tying) into flags.
5. Concise, board-report tone.

Output JSON only:
{"headline": "one line (<=20 words)", "commentary": ["point 1 with a number", "..."], "flags": [...]}"""


def df_to_markdown(df: pd.DataFrame, max_rows: int = 40, floatfmt: str = ",.1f") -> str:
    """把表轉成精簡 markdown（俾 LLM 睇）。太多行就截，保留合計行。"""
    d = df.copy()
    if len(d) > max_rows:
        tail = d[d.index.astype(str).str.contains("合計|total|Total")] if d.index.dtype == object else d.iloc[0:0]
        d = pd.concat([d.head(max_rows - len(tail)), tail])
    try:
        return d.to_markdown(floatfmt=floatfmt)
    except Exception:
        return d.round(1).to_string()


def build_user_prompt(table_md: str, *, context: dict[str, Any] | None = None,
                      unit: str = "萬 MOP", extra_instructions: str = "") -> str:
    ctx = context or {}
    ctx_lines = "\n".join(f"- {k}：{v}" for k, v in ctx.items())
    parts = [
        f"單位：{unit}。",
        f"背景：\n{ctx_lines}" if ctx_lines else "",
        f"報告表格：\n{table_md}",
        f"額外要求：{extra_instructions}" if extra_instructions else "",
        "請按 system 指示，只輸出 JSON。",
    ]
    return "\n\n".join(p for p in parts if p)


def generate_notes(table: pd.DataFrame | str, *, context: dict[str, Any] | None = None,
                   unit: str = "萬 MOP", lang: str = "zh", extra_instructions: str = "",
                   wb: Workbench | None = None, model: str | None = None,
                   dry: bool = False) -> dict:
    """由表生成 {headline, commentary[], flags[]}。dry=True 唔出網，只回構建好嘅 prompt。"""
    table_md = table if isinstance(table, str) else df_to_markdown(table)
    system = NOTE_SYSTEM_ZH if lang.lower().startswith("zh") else NOTE_SYSTEM_EN
    user = build_user_prompt(table_md, context=context, unit=unit, extra_instructions=extra_instructions)
    wb = wb or Workbench(model=model)
    if dry:
        return {"_dry": True, "system": system, "user": user, "model": wb.resolve_model(model)}
    out = wb.chat_json(user, system=system, model=model)
    out.setdefault("headline", "")
    out.setdefault("commentary", [])
    out.setdefault("flags", [])
    return out


def _cli() -> None:
    import argparse
    from pathlib import Path
    ap = argparse.ArgumentParser(description="databook 註釋生成（demo）")
    ap.add_argument("table", nargs="?", help="表格檔（.tsv/.csv/.xlsx；index 為第一欄）")
    ap.add_argument("--lang", default="zh", help="zh / en")
    ap.add_argument("--unit", default="萬 MOP")
    ap.add_argument("--model", default=None, help="alias/實名；預設由 config")
    ap.add_argument("--context", default="", help="k=v;k=v，例如 entity=mgm;year=25")
    ap.add_argument("--dry", action="store_true", help="唔出網，只印會送出嘅 prompt")
    a = ap.parse_args()

    if a.table:
        p = Path(a.table)
        if p.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(p, index_col=0)
        else:
            df = pd.read_csv(p, sep="\t" if p.suffix.lower() == ".tsv" else ",", index_col=0)
    else:  # 內置 sample（同 adj_pivot 輸出同形）
        df = pd.DataFrame({
            "調整前": [186142.0, 64264.0, 5529.0],
            "超支調整": [-50000.0, -12000.0, -800.0],
            "重複入賬": [-18522.0, -4000.0, -100.0],
            "調整合計": [-68522.0, -16000.0, -900.0],
            "調整後": [117620.0, 48264.0, 4629.0],
            "調整後/前%": [63.2, 75.1, 83.7],
        }, index=["NG0 博彩項目", "NG5 文化藝術", "NG8 美食"])
        df.loc["合計"] = df.drop(columns=["調整後/前%"]).sum()
        df.loc["合計", "調整後/前%"] = df.loc["合計", "調整後"] / df.loc["合計", "調整前"] * 100

    ctx = dict(kv.split("=", 1) for kv in a.context.split(";") if "=" in kv) if a.context else {}
    res = generate_notes(df, context=ctx, unit=a.unit, lang=a.lang, model=a.model, dry=a.dry)
    if res.get("_dry"):
        print("=== SYSTEM ===\n" + res["system"])
        print("\n=== USER ===\n" + res["user"])
        print(f"\n=== MODEL === {res['model']}   （--dry：未出網。去 KPMG 網內去 --dry 拎走跑真）")
    else:
        print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _cli()
