#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_llm_narrative.py — 用 KPMG Workbench LLM 由底層數據（feed + 清單 +（可）表2）寫報告式
summary，出 {entity}_llm_narrative.json。make_report 見到呢個檔就用 LLM 文字（唔再淨抄清單原文）。

連接方式參考 python-pptx/fdd_utils/ai.py（Workbench = Azure OpenAI gateway，Ocp-Apim-Subscription-Key
header + charge-code/region）。本 repo 已有 client：src/kpi/lib/workbench.py（config-driven）。
「utilisation」＝ ThreadPoolExecutor 併發（--workers），一次過批多個 summary。

要 KPMG 網 + creds：conf/local/credentials.yml『workbench:』或 env WB_API_KEY/WB_BASE_URL/…。
用法（Windows，KPMG 網內）：
    python scripts\\report\\build_llm_narrative.py "tableau_combined_25.csv" --entity mgm ^
        --qingdan "data\\投資項目清單\\3. MGM.…投资项目清单.xlsx" [--workers 3] [--model 5.5]
    # 唔想出網，先驗 config： ... --config
"""
import json
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "src"))
try:
    import pandas as pd
except ImportError:
    print("✗ pip install pandas openpyxl"); sys.exit(1)

import build_summary_tables as S
import build_project_review_table as B
import build_overview_tables as O
import build_narrative as N
from kpi.lib.workbench import Workbench

SYS = ("你係畢馬威（KPMG）投資計劃執行情況審查報告嘅專業撰稿員。用【繁體中文】書面語，"
       "審查報告語氣：精簡、客觀、專業、第三人稱（用『我們』，唔好口語）。"
       "只可根據所提供嘅資料撰寫，嚴禁虛構、誇大或加入未提供嘅事實/數字。"
       "★重要：直接寫有嘅內容，切勿寫『未獲提供』『資料不足』『未能說明』等 meta 或免責語句；"
       "若某類資料缺，就只寫有嘅部分。輸出淨係一段連貫文字（唔好標題/項目符號/開場白/結語），"
       "貼近正式審查報告書面語，忌冗長。")


def _adj_prompt(adj_type, amt_wan, projects):
    lines = [f"潛在調整類型：{adj_type}", f"涉及潛在調減金額：約{abs(amt_wan):,.0f}萬澳門元", "涉及項目及審查發現："]
    for name, find, mgmt in projects[:6]:
        seg = f"- 項目「{name}」"
        if find:
            seg += f"；KPMG分析發現：{find[:320]}"
        if mgmt:
            seg += f"；管理層解釋：{mgmt[:220]}"
        lines.append(seg)
    ctx = "\n".join(lines)
    return (f"以下係一項『潛在調整事項』嘅底層資料。請寫一段報告摘要（約100-180字），"
            f"說明該調整類型、金額、主要涉及嘅投資項目同調減原因（綜合 KPMG 分析與管理層解釋），"
            f"並帶出審查建議（通常為建議剔除／調減）。\n\n{ctx}")


def _cat_prompt(sub, rate_pct, content, reason):
    ctx = (f"投資範疇：{sub}\n投資計劃金額完成率：{rate_pct}\n"
           f"該範疇實際投資內容：{content[:450]}\n管理層解釋／變更原因：{reason[:320]}")
    return (f"請為投資執行概況寫一句『按範疇的項目概況』（約60-130字），"
            f"格式類似「{sub}：主要包括……。投資計劃金額完成率為{rate_pct}，主要由於……」，"
            f"綜合實際投資內容同完成率原因。\n\n{ctx}")


def _gen(wb, prompt, effort):
    return wb.chat(prompt, SYS, reasoning_effort=effort).strip()


def main():
    args = sys.argv[1:]
    entity = qingdan = model = None
    workers = 3
    cfg_only = "--config" in args
    if cfg_only:
        args.remove("--config")
    for flag in ("--entity", "--qingdan", "--model", "--workers"):
        if flag in args:
            i = args.index(flag); val = args[i + 1]; del args[i:i + 2]
            if flag == "--entity":
                entity = val.lower()
            elif flag == "--qingdan":
                qingdan = val
            elif flag == "--model":
                model = val
            elif flag == "--workers":
                workers = int(val)
    wb = Workbench(model=model)
    print("Workbench config（key 遮蔽）:")
    for k, v in wb.config_masked().items():
        print(f"  {k}: {v}")
    if cfg_only:
        print("\n（--config：只驗 config，唔出網）"); return
    if not args:
        print("俾 tableau feed csv 路徑（--qingdan 清單）"); return

    df = S._load(Path(args[0]), entity)
    narr = N.load_narrative(Path(qingdan)) if qingdan else {}
    plan = B.load_plan(Path(qingdan)) if qingdan else None
    ov = O.overview_by_bucket(df, "2025年度投資計劃", plan)
    adj = O.adjustment_bridge(df)
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])

    # 併裝 tasks（(kind, key, prompt, effort)）
    tasks = []
    for _, r in adj.iterrows():
        t = r["潛在調整事項"]
        if t in ("合計", "跨年及其他調整"):
            continue
        amt = r.get("合計", 0)
        if not isinstance(amt, (int, float)) or abs(amt) < 0.5:
            continue
        sub = d[(d["_adj"] == t) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        projs = []
        for _, pp in sub.drop_duplicates("dicj code").iterrows():
            nr = N.nlook(narr, pp["ng_scope"], pp["dicj code"])
            projs.append((str(pp["project"]), nr.get("KPMG分析發現", ""), nr.get("管理層解釋", "")))
        tasks.append(("adj", t, _adj_prompt(t, amt, projs), "medium"))

    proj = d.groupby(["_sub", "dicj code"])["調整前_萬"].sum().reset_index()
    for _, r in ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].iterrows():
        sub = str(r["範疇"]); rate = r.get("投資計劃完成率")
        if not isinstance(rate, (int, float)) or pd.isna(rate):
            continue
        scope = "gaming" if sub.startswith("博彩") else "non_gaming"
        content = reason = ""
        for _, pp in proj[proj["_sub"] == sub].sort_values("調整前_萬", ascending=False).iterrows():
            nr = N.nlook(narr, scope, pp["dicj code"])
            content = content or nr.get("實際投資內容", "")
            reason = reason or nr.get("管理層解釋", "") or nr.get("KPMG分析發現", "")
            if content and reason:
                break
        tasks.append(("cat", sub, _cat_prompt(sub, f"{rate*100:.1f}%", content, reason), "low"))

    print(f"\n（{entity}）批 {len(tasks)} 個 summary，workers={workers}…")
    out = {"adj": {}, "cat": {}}
    with ThreadPoolExecutor(max_workers=workers) as ex:
        fut = {ex.submit(_gen, wb, p, eff): (kind, key) for kind, key, p, eff in tasks}
        for f in as_completed(fut):
            kind, key = fut[f]
            try:
                out[kind][key] = f.result()
                print(f"  ✓ {kind}｜{key[:22]}")
            except Exception as e:
                print(f"  ⚠ {kind}｜{key[:22]}: {type(e).__name__}: {e}")

    outp = Path(f"{entity or 'all'}_llm_narrative.json")
    outp.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n✓ {outp.resolve()}（adj {len(out['adj'])}、cat {len(out['cat'])} 段）→ make_report 會自動用")


if __name__ == "__main__":
    main()
