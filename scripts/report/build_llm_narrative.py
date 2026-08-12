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
import biao2 as B2
from kpi.lib.workbench import Workbench

# 調整事項節（S16-17）：KPMG 審計角度，可講剔除/調減建議
SYS_ADJ = ("你係畢馬威（KPMG）投資計劃執行情況審查報告嘅專業撰稿員。用【繁體中文】書面語，"
           "審查報告語氣：精簡、客觀、專業、第三人稱（用『我們』）。"
           "只可根據所提供嘅資料撰寫，嚴禁虛構、誇大或加入未提供嘅事實/數字。"
           "直接寫有嘅內容，切勿寫『未獲提供』『資料不足』等 meta/免責語句。"
           "輸出淨係一段連貫文字（唔好標題/項目符號/開場白/結語），忌冗長。")

# 按範疇概況節（S13-15）：承批公司『2025年度投資執行報告』角度，唔好審計腔
SYS_CAT = ("你係為承批公司『2025年度投資執行報告』撰寫『按範疇的項目概況』嘅撰稿員。用【繁體中文】書面語，"
           "站喺投資執行角度：描述該範疇實際投資咗啲乜（可含具體項目、子項目、活動／賽事／音樂會場次），"
           "再講完成率點解係咁。語氣自然、貼近企業投資執行報告，唔好似審計底稿、唔好似機器砌 list。"
           "★嚴禁用審計／調整用語：『剔除』『調減』『申報口徑』『可計入範圍』『超出範圍』『再次申報』"
           "『偏離』『不符合定義』『未在計劃中列示』等 —— 呢啲屬另一節（調整事項），概況絕不出現。"
           "完成率原因只用管理層嘅業務解釋（例：施工進度較預期延遲、實際所需資金低於預算、"
           "進度高於預期、活動如期舉辦、發現結構性問題增加成本），唔好用審計理由。"
           "直接寫有嘅內容，切勿寫『未獲提供』『資料不足』等 meta 語。"
           "輸出淨係一段連貫文字（唔好項目符號／開場白／結語），語句要順，忌逐點堆砌。")


def _adj_prompt(adj_type, amt_wan, projects):
    lines = [f"潛在調整類型：{adj_type}", f"涉及潛在調減金額：約{abs(amt_wan):,.0f}萬澳門元",
             "涉及項目及審查發現（審查底稿表2 為最權威來源，優先採用其跨司裁決及具體內容）："]
    for name, find, mgmt, b2, ruling in projects[:6]:
        seg = f"- 項目「{name}」"
        if b2:      # 表2＝審查底稿，最可信，放最前、俾最多
            seg += f"；【審查底稿表2】{b2[:900]}"
        if find:
            seg += f"；KPMG分析發現：{find[:260]}"
        if mgmt:
            seg += f"；管理層解釋：{mgmt[:200]}"
        if ruling:
            seg += f"；跨司工作組／KPMG裁決（清單）：{ruling[:220]}"
        lines.append(seg)
    ctx = "\n".join(lines)
    return (f"以下係一項『潛在調整事項』嘅底層資料（審查底稿表2 內容最詳盡，可用作事實依據）。"
            f"請寫一段報告摘要（約120-200字），說明該調整類型、金額、主要涉及嘅投資項目同調減原因。"
            f"★用字須跟原報告：如有向跨司工作組諮詢得到嘅回覆，用『跨司工作組』集體稱呼帶出其立場"
            f"（例如『根據我們向跨司工作組諮詢得到的回覆，跨司工作組認為／未同意…』），"
            f"【切勿】逐個司局點名（如社會文化司、旅遊局、文化局），亦【切勿】自創『KPMG最終立場』等標籤。"
            f"最後點出審查建議（通常為建議剔除／調減）。\n\n{ctx}")


def _cat_prompt(sub, rate_pct, content, reason, b2=""):
    ctx = (f"投資範疇：{sub}\n投資計劃金額完成率：{rate_pct}\n"
           f"該範疇實際投資內容（項目清單）：{content[:500]}\n"
           f"管理層變更原因／業務解釋：{reason[:340]}\n"
           f"表2 補充（只可攞嚟豐富『投資內容』，例如子項目／活動場次／金額明細；"
           f"切勿抄佢嘅審計措辭或調整理由）：{b2[:700]}")
    return (f"請為承批公司投資執行報告寫一句『按範疇的項目概況』（約70-140字），"
            f"格式：「{sub}：主要包括……（實際投資咗啲乜，如有具體子項目／活動場次請寫）。"
            f"投資計劃金額完成率為{rate_pct}，主要由於……（管理層業務原因）」。"
            f"完成率原因只用管理層業務解釋，唔好用審計／調整措辭。\n\n{ctx}")


def _gen(wb, prompt, effort, sysp):
    return wb.chat(prompt, sysp, reasoning_effort=effort).strip()


def generate_llm_narrative(feed_path, entity, qingdan, biao2_dir="data/表2",
                           model=None, workers=3, out_path=None, log=print):
    """由 feed + 清單 + 表2 用 Workbench 生成 {adj,cat} 敘述；寫 {entity}_llm_narrative.json，回 dict。
    可被 build_report.py --llm 直接調用（唔使另跑 command）。"""
    wb = Workbench(model=model)
    df = S._load(Path(feed_path), entity)
    narr = N.load_narrative(Path(qingdan)) if qingdan else {}
    b2 = B2.load_biao2(biao2_dir, entity or "", log=log)
    plan = B.load_plan(Path(qingdan)) if qingdan else None
    cat = B.load_category(Path(qingdan)) if qingdan else None
    ov = O.overview_by_bucket(df, "2025年度投資計劃", plan, cat)
    adj = O.adjustment_bridge(df)
    d = df.copy()
    d["_adj"] = d["調整一級"].map(B.CANON).fillna(d["調整一級"])

    # 併裝 tasks（(kind, key, prompt, effort, sys)）
    pb = S.BUCKET_ORDER[0]      # 2025計劃 bucket：調整詳述只計 2025年度計劃（期後另計，對返報告）
    tasks = []
    for _, r in adj.iterrows():
        t = r["潛在調整事項"]
        if t in ("合計", "跨年及其他調整"):
            continue
        amt = r.get(pb, 0)
        if not isinstance(amt, (int, float)) or abs(amt) < 0.5:
            continue
        sub = d[(d["_adj"] == t) & (d["_bucket"] == pb) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        projs = []
        for _, pp in sub.drop_duplicates("dicj code").iterrows():
            nr = N.nlook(narr, pp["ng_scope"], pp["dicj code"])
            b2t = B2.b2look(b2, pp["ng_scope"], pp["dicj code"])
            ruling = "；".join(x for x in (nr.get("跨司回覆", ""), nr.get("KPMG回覆", "")) if x)
            projs.append((str(pp["project"]), nr.get("KPMG分析發現", ""),
                          nr.get("管理層解釋", ""), b2t, ruling))
        tasks.append(("adj", t, _adj_prompt(t, amt, projs), "medium", SYS_ADJ))

    proj = d.groupby(["_sub", "dicj code"])["調整前_萬"].sum().reset_index()
    for _, r in ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].iterrows():
        sub = str(r["範疇"]); rate = r.get("投資計劃完成率")
        if not isinstance(rate, (int, float)) or pd.isna(rate):
            continue
        scope = "gaming" if sub.startswith("博彩") else "non_gaming"
        content = reason = b2t = ""
        for _, pp in proj[proj["_sub"] == sub].sort_values("調整前_萬", ascending=False).iterrows():
            nr = N.nlook(narr, scope, pp["dicj code"])
            content = content or nr.get("實際投資內容", "")
            reason = reason or nr.get("管理層解釋", "")   # 業務原因；唔用 KPMG分析發現（審計腔）
            if not b2t:
                b2t = B2.b2look(b2, scope, pp["dicj code"])
            if content and reason and b2t:
                break
        tasks.append(("cat", sub, _cat_prompt(sub, f"{rate*100:.1f}%", content, reason, b2t), "low", SYS_CAT))

    log(f"（{entity}）批 {len(tasks)} 個 summary，workers={workers}…")
    out = {"adj": {}, "cat": {}}
    try:                                    # tqdm 進度條（冇裝就照 log 逐個）
        from tqdm import tqdm
    except ImportError:
        tqdm = None
    with ThreadPoolExecutor(max_workers=workers) as ex:
        fut = {ex.submit(_gen, wb, p, eff, sysp): (kind, key) for kind, key, p, eff, sysp in tasks}
        it = as_completed(fut)
        bar = tqdm(it, total=len(tasks), desc=f"LLM {entity}", unit="段", ncols=90) if tqdm else it
        for f in bar:
            kind, key = fut[f]
            try:
                out[kind][key] = f.result()
                msg = f"  ✓ {kind}｜{key[:22]}"
            except Exception as e:
                msg = f"  ⚠ {kind}｜{key[:22]}: {type(e).__name__}: {e}"
            tqdm.write(msg) if tqdm else log(msg)   # tqdm.write 唔會撞爛進度條

    outp = Path(out_path) if out_path else Path(f"{entity or 'all'}_llm_narrative.json")
    outp.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"✓ {outp.resolve()}（adj {len(out['adj'])}、cat {len(out['cat'])} 段）")
    return out


def main():
    args = sys.argv[1:]
    entity = qingdan = model = None
    biao2_dir = "data/表2"
    workers = 3
    cfg_only = "--config" in args
    if cfg_only:
        args.remove("--config")
    for flag in ("--entity", "--qingdan", "--model", "--workers", "--biao2"):
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
            elif flag == "--biao2":
                biao2_dir = val
    wb = Workbench(model=model)
    print("Workbench config（key 遮蔽）:")
    for k, v in wb.config_masked().items():
        print(f"  {k}: {v}")
    if cfg_only:
        print("\n（--config：只驗 config，唔出網）"); return
    if not args:
        print("俾 tableau feed csv 路徑（--qingdan 清單）"); return
    generate_llm_narrative(args[0], entity, qingdan, biao2_dir, model, workers)


if __name__ == "__main__":
    main()
