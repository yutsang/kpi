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
import re
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


# 表旁 comment（對 scan p-10~p-13：表左 + 敘述右）。兩種語氣：審查（調整/發現）vs 描述（金額/設施活動）
SYS_TBL_ADJ = (SYS_ADJ.replace("輸出淨係一段連貫文字（唔好標題/項目符號/開場白/結語），忌冗長。", "")
               + "你而家寫嘅係【一張報告表格旁邊嘅敘述】：解釋張表講緊乜、關鍵金額同背後原因。"
               "只可引用表格入面真係有嘅數字，唔可以自己計新數或估數。"
"★金額一律用返報告嘅正式名：『報告投資金額』『潛在調整後投資金額』『獲批的計劃投資金額』"
               "『潛在調減』。【嚴禁自創】報告冇嘅字眼，例如『經後續管理檢視後』『經覆核後』"
               "『管理層檢視』『調整後淨額』等 —— 一律用『潛在調整後投資金額』。"
               "輸出 JSON：{\"導語\":\"…\",\"段落\":[{\"小標題\":\"…\",\"內容\":\"…\"}]}。"
               "『導語』＝成版最頂嗰句總結（80-160 字，跟報告句式，見 prompt 內示範）；"
               "『段落』2 至 4 段，每段小標題 ≤14 字、內容 60-130 字。")
SYS_TBL_DESC = (SYS_CAT.replace("輸出淨係一段連貫文字（唔好項目符號／開場白／結語），語句要順，忌逐點堆砌。", "")
                + "你而家寫嘅係【一張報告表格旁邊嘅敘述】：解釋張表講緊乜、邊啲範疇金額最大、"
                "設施建設同活動舉辦嘅比重點樣。只可引用表格入面真係有嘅數字，唔可以自己計新數或估數。"
                "★金額一律用返報告嘅正式名（報告投資金額／潛在調整後投資金額／獲批的計劃投資金額）；"
                "【嚴禁自創】『經後續管理檢視後』『經覆核後』等報告冇嘅字眼。"
 "★金額一律用返報告嘅正式名：『報告投資金額』『潛在調整後投資金額』『獲批的計劃投資金額』"
               "『潛在調減』。【嚴禁自創】報告冇嘅字眼，例如『經後續管理檢視後』『經覆核後』"
               "『管理層檢視』『調整後淨額』等 —— 一律用『潛在調整後投資金額』。"
               "輸出 JSON：{\"導語\":\"…\",\"段落\":[{\"小標題\":\"…\",\"內容\":\"…\"}]}。"
                "『導語』＝成版最頂嗰句總結（80-160 字）；『段落』2 至 4 段，"
                "每段小標題 ≤14 字、內容 60-130 字。")


def tbl_key(kind, arg=""):
    """表旁 comment 嘅 key（generator 同 make_report 必須用同一個）。"""
    return f"{kind}|{arg}"


def proj_key(adj_type, ng_scope, code):
    """主要發現 card 逐項目『事項描述』嘅 key（generator 同 make_report 必須用同一個）。"""
    return f"{adj_type}|{ng_scope}|{B._norm(code)}"


def bkt_key(bucket, adj_type):
    """期後調整事項匯總（scan p-11/p-13）逐類開場句嘅 key。"""
    return f"{bucket}|{adj_type}"


def _bkt_prompt(yr, adj_type, amt_wan, projects):
    lines = [f"年度：{yr}年度投資計劃期後投資（於2025年發生）", f"調整類型：{adj_type}",
             f"該類潛在調減金額：約{abs(amt_wan):,.0f}萬澳門元", "涉及項目："]
    for nm, amt, b2, find in projects[:5]:
        seg = f"- {nm}（{abs(amt):,.0f}萬澳門元）"
        if b2:
            seg += f"；【審查底稿表2】{b2[:600]}"
        elif find:
            seg += f"；KPMG分析發現：{find[:300]}"
        lines.append(seg)
    return ("請寫【一句至兩句】開場描述，講清楚喺該年度期後投資金額中，承批公司申報咗啲乜"
            "而我哋認為要調整。示範句式（要用返下面嘅真數同項目，唔好照抄）：\n"
            "　『在2024年度投資計劃期後投資金額中，MGM申報了澳門美高梅國際旗艦級藝術珍寶博物館"
            "營運後的營運成本（827萬澳門元）。』\n"
            "　『在2024年度投資計劃期後投資金額中，MGM仍申報了酒店客房改造支出，主要包括："
            "1）非博彩項目111多功能娛樂體驗區塊（娛樂表演範疇）的相關支出420萬澳門元；"
            "2）非博彩項目21美獅美高梅高端康養醫療中心的相關支出3,532萬澳門元。』\n"
            "★只寫呢一兩句，唔好寫調整建議／跨司意見／結論（後面有固定句接落去）。\n\n"
            + "\n".join(lines))


def _proj_prompt(adj_type, name, code, rep, adjv, find, mgmt, b2, ruling, content=""):
    ctx = [f"投資項目：{code}　{name}", f"潛在調整類型：{adj_type}",
           f"報告投資金額：{rep:,.0f}萬澳門元；本類潛在調整：{adjv:,.0f}萬澳門元"]
    if content:
        ctx.append(f"實際投資內容（項目清單）：{content[:400]}")
    if b2:      # 表2＝審查底稿，最權威，俾最多
        ctx.append(f"審查底稿表2（關注事項／調整原因／跨司意見，事實依據）：{b2[:1200]}")
    if find:
        ctx.append(f"KPMG分析發現（項目清單）：{find[:400]}")
    if mgmt:
        ctx.append(f"承批公司管理層解釋：{mgmt[:300]}")
    if ruling:
        ctx.append(f"跨司工作組／KPMG回覆（項目清單）：{ruling[:300]}")
    return ("請為報告『本年度審查工作的主要發現』其中一個投資項目，寫一段【事項描述】"
            "（約150-250字）：講清楚該項目投資咗啲乜、我們喺審查中發現咗咩、"
            "點解相關支出不應／只可部分計入報告投資金額、以及調整金額。"
            "★如有跨司工作組回覆，用『跨司工作組』集體稱呼（例：『根據我們向跨司工作組諮詢得到的回覆，"
            "跨司工作組認為…』），切勿逐個司局點名，亦切勿自創『KPMG最終立場』等標籤。"
            "只可用下面提供嘅事實同數字，唔可以虛構。\n\n" + "\n".join(ctx))


def _tbl_text(df, max_rows=40):
    """DataFrame → 精簡 TSV 餵 LLM（數字原封不動）。"""
    cols = list(df.columns)
    lines = ["\t".join(str(c) for c in cols)]
    for _, r in df.head(max_rows).iterrows():
        lines.append("\t".join("" if pd.isna(r[c]) else str(r[c]) for c in cols))
    return "\n".join(lines)


def _tbl_prompt(title, df, sources, unit="萬澳門元"):
    src = ("\n".join(f"- {s}" for s in sources[:6])) if sources else "（無額外資料，只根據表格數字撰寫）"
    return (f"以下係報告入面一張表，請寫佢【旁邊】嘅敘述，同埋成版最頂嗰句【導語】。\n\n"
            f"★導語要跟返呢份報告一貫句式（示範，唔好照抄字眼，要用返下面表格嘅真數）：\n"
            f"　『…在2025年度執行報告中申報的「因發生期後事項需作後續調整之2024年度博彩／非博彩項目」"
            f"投資金額為6.4億澳門元，主要包括…以及…。本次審查工作識別潛在調減金額約4.8億澳門元，"
            f"調減後金額為1.6億澳門元，主要涉及會議展覽、文化藝術、社區旅遊等非博彩投資範疇的37個項目。』\n"
            f"　金額單位跟報告習慣：≥1億寫『X.X億澳門元』（一位小數），唔夠1億寫『X,XXX萬澳門元』"
            f"（【整數、千分位、冇小數】—— 唔可以寫『5,528.9萬澳門元』，要寫『5,529萬澳門元』）。\n\n"
            f"表名：{title}\n金額單位：{unit}（括號 = 負數／調減，「-」= 零）\n\n"
            f"【表格內容】\n{_tbl_text(df)}\n\n"
            f"【其他來源（項目清單／審查底稿表2，用嚟解釋原因，唔好抄佢嘅措辭）】\n{src}")


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


def _cat_prompt(sub, rate_pct, projects, reason, b2=""):
    """projects = [(序號, 名稱, 報告金額萬, 實際投資內容)]，按金額大到細。
    ⚠ 一定要逐個項目【點名】—— 之前只餵一個項目嘅內容去寫成個範疇，讀者分唔清邊句開始
      講新項目（項目組 2026-08-13 反映）。"""
    lines = []
    for code, name, amt, content in projects[:2]:      # 2 個夠：一版要放晒 11 個範疇
        lines.append(f"- {code}「{name}」（報告投資金額 {amt:,.0f} 萬澳門元）：{str(content)[:240]}")
    ctx = (f"投資範疇：{sub}\n投資計劃金額完成率：{rate_pct}\n"
           f"該範疇金額最大嘅投資項目（項目清單）：\n" + "\n".join(lines) + "\n"
           f"管理層變更原因／業務解釋：{str(reason)[:340]}\n"
           f"表2 補充（只可攞嚟豐富『投資內容』，例如子項目／活動場次／金額明細；"
           f"切勿抄佢嘅審計措辭或調整理由）：{str(b2)[:700]}")
    # ⚠ 字數上限係【版面約束】：報告「按範疇的項目概況」博彩／非博彩各佔【一版】，
    #   非博彩 11 個範疇要一版放晒 → 每個範疇最多 ~55 字，寫長咗就會分成三版（同報告唔同）。
    return (f"請為承批公司投資執行報告寫一句『按範疇的項目概況』"
            f"（【嚴格 40-55 字】，超過就會排版爆版，寧短勿長；唔好客套話、唔好重覆範疇名）。\n"
            f"★格式：「主要包括{{項目序號}}「{{項目名稱}}」……；{{項目序號}}「{{項目名稱}}」……。"
            f"完成率{rate_pct}，主要由於……（管理層業務原因，一句起兩句止）」\n"
            f"★【每個項目必須先寫返項目序號同項目名稱】先講佢做咗乜，項目與項目之間用「；」分開，"
            f"令讀者一眼睇到邊句係講邊個項目 —— 唔可以將幾個項目嘅內容混埋一齊寫。\n"
            f"完成率原因只用管理層業務解釋，唔好用審計／調整措辭。\n\n{ctx}")


def _short_err(e):
    """gateway 錯誤成日回一版 HTML → 抽 <h2>/<title> 或者剝晒 tag，最多 120 字。"""
    t = str(e)
    m = re.search(r"<h2[^>]*>(.*?)</h2>|<title[^>]*>(.*?)</title>", t, re.S | re.I)
    if m:
        t = (m.group(1) or m.group(2) or "").strip()
    elif "<html" in t.lower():
        t = re.sub(r"<[^>]+>", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t[:120] + ("…" if len(t) > 120 else "")


def _gen(wb, prompt, effort, sysp, want_json=False):
    if not want_json:
        return wb.chat(prompt, sysp, reasoning_effort=effort).strip()
    d = wb.chat_json(prompt, sysp, reasoning_effort=effort)
    segs = (d or {}).get("段落") or []
    out = [[str(s.get("小標題", "")).strip(), str(s.get("內容", "")).strip()]
           for s in segs if isinstance(s, dict) and s.get("內容")]
    if not out:
        raise ValueError("LLM 冇回到『段落』")
    return {"導語": str((d or {}).get("導語", "")).strip(), "段落": out}


def _proj_sources(d, narr, b2, mask, kind="content", n=5):
    """由 feed 一段 slice 抽 top 項目嘅來源片段（清單 + 表2）→ [str]，餵表旁 comment。"""
    sub = d[mask]
    if sub.empty:
        return []
    key = "調整_萬" if kind == "finding" else "調整前_萬"
    top = (sub.groupby(["ng_scope", "dicj code"])
              .agg(nm=("project", "first"), v=(key, "sum")).reset_index())
    top = top.reindex(top["v"].abs().sort_values(ascending=False).index).head(n)
    out = []
    for _, p in top.iterrows():
        nr = N.nlook(narr, p["ng_scope"], p["dicj code"])
        b2t = B2.b2text(b2, p["ng_scope"], p["dicj code"])
        if kind == "finding":
            txt = nr.get("KPMG分析發現", "") or b2t
            mg = nr.get("管理層解釋", "")
            seg = f"項目「{p['nm']}」（{p['v']:,.0f}萬）：{str(txt)[:300]}"
            if mg:
                seg += f"；管理層解釋：{mg[:160]}"
        else:
            txt = nr.get("實際投資內容", "") or b2t
            seg = f"項目「{p['nm']}」（{p['v']:,.0f}萬）：{str(txt)[:300]}"
        if str(txt).strip():
            out.append(seg)
    return out


def generate_llm_narrative(feed_path, entity, qingdan, biao2_dir="data/表2",
                           model=None, workers=8, out_path=None, log=print):
    """由 feed + 清單 + 表2 用 Workbench 生成 {adj,cat} 敘述；寫 {entity}_llm_narrative.json，回 dict。
    可被 build_report.py --llm 直接調用（唔使另跑 command）。"""
    wb = Workbench(model=model)
    df = S._load(Path(feed_path), entity)
    narr = N.load_narrative(Path(qingdan)) if qingdan else {}
    b2 = B2.load_biao2_struct(biao2_dir, entity or "", log=log)
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
        sub = d[(d["_adj"] == t) & (pd.to_numeric(d["調整_萬"], errors="coerce") != 0)]
        agg = sub.groupby(["ng_scope", "dicj code"]).agg(
            nm=("project", "first"), rep=("調整前_萬", "sum"), adjv=("調整_萬", "sum")).reset_index()
        agg = agg.reindex(agg["adjv"].abs().sort_values(ascending=False).index)
        projs = []
        for _, pp in agg.iterrows():
            nr = N.nlook(narr, pp["ng_scope"], pp["dicj code"])
            b2t = B2.b2text(b2, pp["ng_scope"], pp["dicj code"])
            ruling = "；".join(x for x in (nr.get("跨司回覆", ""), nr.get("KPMG回覆", "")) if x)
            projs.append((str(pp["nm"]), nr.get("KPMG分析發現", ""),
                          nr.get("管理層解釋", ""), b2t, ruling))
            # 逐項目『事項描述』（主要發現 card）：用表2 + 清單寫返報告嗰種 narrative
            tasks.append(("proj", proj_key(t, pp["ng_scope"], pp["dicj code"]),
                          _proj_prompt(t, pp["nm"], pp["dicj code"], pp["rep"], pp["adjv"],
                                       nr.get("KPMG分析發現", ""), nr.get("管理層解釋", ""),
                                       b2t, ruling, nr.get("實際投資內容", "")),
                          "medium", SYS_ADJ))
        tasks.append(("adj", t, _adj_prompt(t, amt, projs), "medium", SYS_ADJ))

    proj = d.groupby(["_sub", "dicj code"])["調整前_萬"].sum().reset_index()
    for _, r in ov[~ov["範疇"].astype(str).str.endswith(("小計", "總計", "項目"))].iterrows():
        sub = str(r["範疇"]); rate = r.get("投資計劃完成率")
        if not isinstance(rate, (int, float)) or pd.isna(rate):
            continue
        scope = "gaming" if sub.startswith("博彩") else "non_gaming"
        projs, reason, b2t = [], "", ""
        for _, pp in proj[proj["_sub"] == sub].sort_values("調整前_萬", ascending=False).iterrows():
            nr = N.nlook(narr, scope, pp["dicj code"])
            if len(projs) < 3:
                projs.append((str(pp["dicj code"]), str(nr.get("項目名稱", "") or pp["dicj code"]),
                              float(pp["調整前_萬"] or 0), nr.get("實際投資內容", "")))
            reason = reason or nr.get("管理層解釋", "")   # 業務原因；唔用 KPMG分析發現（審計腔）
            b2t = b2t or B2.b2text(b2, scope, pp["dicj code"])
            if len(projs) >= 3 and reason and b2t:
                break
        tasks.append(("cat", sub, _cat_prompt(sub, f"{rate*100:.1f}%", projs, reason, b2t),
                      "low", SYS_CAT))

    # ── 表旁 comment（scan p-10~p-13 表左＋敘述右）：由表格數字 + 清單/表2 來源寫 ──
    for bk in S.BUCKET_ORDER[1:]:                       # 2024 / 2023 期後概覽
        ovb = O.overview_by_bucket(df, bk, plan, cat)
        if ovb.empty:
            continue
        src = _proj_sources(d, narr, b2, (d["_bucket"] == bk) &
                            (pd.to_numeric(d["調整_萬"], errors="coerce") != 0), "finding")
        tasks.append(("tbl", tbl_key("期後概覽", bk),
                      _tbl_prompt(f"{(entity or '').upper()} {bk}金額概覽", ovb.fillna(""), src),
                      "medium", SYS_TBL_ADJ, True))
    for bk in S.BUCKET_ORDER:                           # 設施建設 vs 活動舉辦
        fa = S.facility_activity(df, bk)
        if fa.empty:
            continue
        src = _proj_sources(d, narr, b2, d["_bucket"] == bk, "content")
        tasks.append(("tbl", tbl_key("設施活動", bk),
                      _tbl_prompt(f"{(entity or '').upper()} {bk}區分設施建設／活動舉辦的投資金額",
                                  fa.fillna(""), src),
                      "low", SYS_TBL_DESC, True))
    amt = S.summary_amount(df)                          # 4.1 金額匯總
    if not amt.empty:
        tasks.append(("tbl", tbl_key("金額匯總"),
                      _tbl_prompt(f"{(entity or '').upper()} 2025年發生的投資金額匯總",
                                  amt.fillna(""), _proj_sources(d, narr, b2,
                                                                d["_bucket"] == pb, "content")),
                      "low", SYS_TBL_DESC, True))
    for bk in S.BUCKET_ORDER[1:]:                       # 期後調整事項匯總：逐類開場句
        dd = d[d["_bucket"] == bk]
        for t in B.ADJ7:
            sub = dd[dd["_adj"] == t]
            amt = pd.to_numeric(sub["調整_萬"], errors="coerce").sum()
            if abs(amt) < 0.5:
                continue
            g = (sub.groupby(["ng_scope", "dicj code"])
                    .agg(nm=("project", "first"), v=("調整_萬", "sum")).reset_index())
            g = g.reindex(g["v"].abs().sort_values(ascending=False).index)
            projs = [(str(r["nm"]), r["v"], B2.b2text(b2, r["ng_scope"], r["dicj code"]),
                      N.nlook(narr, r["ng_scope"], r["dicj code"]).get("KPMG分析發現", ""))
                     for _, r in g.iterrows()]
            tasks.append(("bkt", bkt_key(bk, t), _bkt_prompt(bk[:4], t, amt, projs),
                          "low", SYS_ADJ))

    fs = O.finding_summary(df)                          # ③ 主要發現摘要
    if not fs.empty:
        tasks.append(("tbl", tbl_key("發現摘要"),
                      _tbl_prompt(f"{(entity or '').upper()} 本年度審查工作的主要發現摘要", fs.fillna(""),
                                  _proj_sources(d, narr, b2,
                                                pd.to_numeric(d["調整_萬"], errors="coerce") != 0,
                                                "finding", n=6)),
                      "medium", SYS_TBL_ADJ, True))

    # ★ preflight：先試一個最平嘅 call。網關擋／key 唔啱嘅話即刻知，唔使等 60 個 task
    #   逐個 fail（2026-08-15 白等 10 分鐘）。
    try:
        wb.chat("ok", "Reply with the single word: ok", reasoning_effort=None, max_tokens=5)
    except Exception as e:
        log(f"  ✗ LLM 連唔到（{type(e).__name__}: {_short_err(e)}）→ 跳過全部 {len(tasks)} 個 summary，"
            "今次用清單／表2 原文 fallback。")
        log("    403／blocked = 網關擋：check 係咪喺 KPMG 內網、key 有冇過期、charge code／region 啱唔啱。")
        out = {"adj": {}, "cat": {}, "tbl": {}, "proj": {}, "bkt": {}}
        outp = Path(out_path) if out_path else Path(f"{entity or 'all'}_llm_narrative.json")
        outp.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
        return out
    log(f"（{entity}）批 {len(tasks)} 個 summary，workers={workers}…")
    out = {"adj": {}, "cat": {}, "tbl": {}, "proj": {}, "bkt": {}}
    try:                                    # tqdm 進度條（冇裝就照 log 逐個）
        from tqdm import tqdm
    except ImportError:
        tqdm = None
    tasks = [(t + (False,))[:6] for t in tasks]         # 補齊 want_json（adj/cat = 純文字）
    with ThreadPoolExecutor(max_workers=workers) as ex:
        fut = {ex.submit(_gen, wb, p, eff, sysp, js): (kind, key)
               for kind, key, p, eff, sysp, js in tasks}
        it = as_completed(fut)
        bar = tqdm(it, total=len(tasks), desc=f"LLM {entity}", unit="段", ncols=90) if tqdm else it
        nfail = 0
        for f in bar:
            kind, key = fut[f]
            try:
                out[kind][key] = f.result()
                msg = f"  ✓ {kind}｜{key[:22]}"
                nfail = 0
            except Exception as e:
                nfail += 1
                # ⚠ err 可能係成版 HTML（KPMG gateway 擋 request 會回錯誤頁）→ 一定要縮短，
                #   否則 console 會俾 60 版 HTML 洗晒版（2026-08-15 實際發生過）
                msg = f"  ⚠ {kind}｜{key[:22]}: {type(e).__name__}: {_short_err(e)}"
            tqdm.write(msg) if tqdm else log(msg)   # tqdm.write 唔會撞爛進度條
        if nfail and not any(out.values()):
            log("  ⚠ LLM 全部失敗 → 今次報告用清單／表2 原文 fallback（唔會空白，但用字唔會似報告）。"
                "\n    常見成因：公司網關擋（並發太多／唔喺 KPMG 網／key 過期）。"
                "\n    試：build_report.py mgm --workers 2；仲係唔得就唔喺內網 or 換 key。")

    outp = Path(out_path) if out_path else Path(f"{entity or 'all'}_llm_narrative.json")
    outp.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"✓ {outp.resolve()}（adj {len(out['adj'])}、cat {len(out['cat'])}、"
        f"tbl {len(out['tbl'])}、proj {len(out['proj'])}、bkt {len(out['bkt'])} 段）")
    return out


def main():
    args = sys.argv[1:]
    entity = qingdan = model = None
    biao2_dir = "data/表2"
    workers = 4
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
