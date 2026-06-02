"""Apply project-team manual labels from raw parquet as H_XXX overrides.

GOLD source: raw parquet has columns where project team manually classified
some rows (comp scope detail / accounting category). These are CUSTOMER-PROVIDED
ground truth — should take precedence over LLM/rules.

For each entity, walk raw parquet:
  1. For each row with manual label → derive target H_XXX via MAPPINGS
  2. Group by signature (account_code|account_desc|desc_norm)
  3. For each sig, pick majority target H (amount-weighted)
  4. Write to feedback.xlsx as override

Total estimated impact: ~9B MOP across 5 entities (Wynn+VML+Melco biggest).

Run:
    python scripts/apply_manual_labels_as_overrides.py --all
    python scripts/apply_manual_labels_as_overrides.py --entity wynn --dry-run
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pandas as pd
import yaml

from kpi.lib.text import signature_key


ENTITIES = {
    "galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
    "vml":    "company_4", "melco":  "company_5", "mgm": "company_6",
}

# V-cols (vertical labels) — NEVER use these to map H. Project type / project nature
# columns describe the V_XXX (vertical) of the row, not the H (horizontal).
# E.g. row with '項目性質=博彩設施及設備的優化' might contain labor/comp lines whose
# H should come from line-level signal, not the project's vertical.
V_LABEL_COLS_SKIP = {
    "galaxy": [],
    "sjm": [],
    "wynn": ["項目類型", "項目性質"],  # 美食之都/娛樂表演 = V cluster
    "vml": ["項目類型", "項目性質"],
    "melco": ["項目性質", "博彩項目標籤", "博彩項目/非博彩項目", "Initiative Name", "initiative ID"],
    "mgm": ["category"],
}

# Per-entity: { raw_col_name: { value: target_H } }
# Single-column rules. Values not in the mapping → row is skipped (no override).
MAPPINGS: dict[str, dict[str, dict[str, str]]] = {
    "galaxy": {
        # 一級標簽 — first-level label. Some values are direct H; "Comp" needs 二級.
        "一級標簽": {
            "staff cost": "H_LABOR",
            "Sponsorship": "H_SPONSORSHIP",
            "Professional Fee": "H_PROFESSIONAL",
            "Marketing": "H_ADVERTISING",
            "Transportation": "H_COMP_OTHER",
            # "Comp" / "comp" → handled by MULTI_COL_RULES below (depends on 二級標簽)
            # "Capex"/"Barter"/"Others"/"Allocation" → skip (no H signal)
        },
        # 二級標簽 — second-level. Mostly used with 一級=Comp (see MULTI_COL_RULES);
        # some values are H-direct regardless of 一級
        "二級標簽": {
            "Salaries & Leave": "H_LABOR",
            "Casual Labor": "H_LABOR",
            "AirTickets": "H_COMP_TICKET",
            "Utilities": "H_OTHER",
        },
        # Source — Galaxy's transaction source field; some are H-specific
        "Source": {
            "5.Staff Cost": "H_LABOR",
            "10.HK Limo": "H_COMP_OTHER",
            "6.FB": "H_FNB",
            "9.ADR": "H_HOTEL_ROOM",
            # "8.MICE" → ambiguous (could be H_VENUE or H_ADVERTISING)
            # "7.GICC&Arena" → ambiguous
        },
        # 调整 — adjustment column; only the room-comp ADR one is clean
        "调整": {
            "用於計算酒店贈房支出的贈房單價超過ADR部分的支出": "H_HOTEL_ROOM",
        },
    },
    "sjm": {
        "Admin comp (Spa 計入hotel，limo/Gift/ticket放其他）": {
            "Hotel Room": "H_HOTEL_ROOM",
            "Hotel room": "H_HOTEL_ROOM",
            "FnB": "H_FNB",
            "Hotel Misc": "H_COMP_OTHER",
            "Hotel-Spa": "H_COMP_OTHER",
            "其他": "H_COMP_OTHER",
        },
    },
    "wynn": {
        "comp费用大类": {
            "食品與飲料支出": "H_FNB",
            "會場支出": "H_VENUE",
            "其他": "H_COMP_OTHER",
            "房間支出": "H_HOTEL_ROOM",
            "門票支出": "H_COMP_TICKET",
        },
        "Breakdown on Comp Expenses in Kind": {
            "Comp Expense, Food": "H_FNB",
            "Comp Expense, Wine": "H_FNB",
            "Comp Expense, Liquor": "H_FNB",
            "Comp Expense, Beer": "H_FNB",
            "Comp Expense, Beverage": "H_FNB",
            "Comp Expense, Venue Rental": "H_VENUE",
            "Comp Expense, Room": "H_HOTEL_ROOM",
            "Comp Expense, Other": "H_COMP_OTHER",
            "Comp Expense, Admission": "H_COMP_TICKET",
            "Service Comp Exp - Food and Venue Rental": "H_FNB",
            "Service Comp Exp - Food": "H_FNB",
            "Service Comp Exp - Hotel": "H_HOTEL_ROOM",
            "Service Comp Exp - Venue": "H_VENUE",
            "Service Comp Exp - Beverage": "H_FNB",
            "Service Comp Exp - Other": "H_COMP_OTHER",
            # New variants from Wynn 25 'Opex&Capex匯總' sheet:
            "Comp Expense, Retail": "H_COMP_OTHER",
            "Comp Expense, CNY Moon Cake": "H_FNB",
            "Service Comp Exp - Spa": "H_COMP_OTHER",
            "Service Comp Exp, Hotel, Misc": "H_COMP_OTHER",
        },
        # MAJOR ADDITION (Wynn 25): 'Nature of Expenses' is the chart-of-accounts
        # H-classifier — 100% coverage, 73 unique values, 2.15B total impact.
        # Previously UNMAPPED. Direct mapping from output of deep_profile_entity_year.
        "Nature of Expenses": {
            "Building Improvements": "H_CONSTRUCTION",
            "Furniture Fixtures & Equipment": "H_EQUIP",
            "Promotional Expense": "H_ADVERTISING",
            "Staff Cost": "H_LABOR",
            "Donation & Sponsorship": "H_SPONSORSHIP",
            "Outside Services": "H_PROFESSIONAL",
            "CIP Renovation Projects": "H_CONSTRUCTION",
            "Advertising Expenses": "H_ADVERTISING",
            "WIN/WRM, Singapore": "H_LABOR",
            "WIN/WRM, Japan": "H_LABOR",
            "Macau Management Costs": "H_PROFESSIONAL",
            "Macau Contruction Costs": "H_CONSTRUCTION",  # typo'd in some rows
            "Operating Items and Equipment": "H_EQUIP",
            "G&A Alloc, WLV/WRM Services Al": "H_PROFESSIONAL",
            "Upgrades": "H_CONSTRUCTION",
            "Government Taxes - Non-Gaming": "H_LICENSE",
            "Equipment Service / Maint Cont": "H_MAINTENANCE",
            "CIP WRM Managed Projects": "H_CONSTRUCTION",
            "F&B Inventory": "H_FNB",
            "Cost of Sales": "H_FNB",  # default; F&B-heavy in Wynn
            # SKIP (context-dependent or already handled by other cols):
            # "Expense in Kind" → comp scope (Breakdown on Comp + comp费用大类 handle this)
            # "Refer WDD Job Detail" → too generic, needs job lookup
            # "Macau Construction Costs" (alt spelling) — duplicate of Building Improvements?
        },
        # NEW: Annex 2 categories (Construction Costs 360M, Staff Cost 4M)
        "Annex 2 Summary Cateogry": {  # note: typo'd in raw col name
            "Macau Construction Costs": "H_CONSTRUCTION",
            "Staff Cost": "H_LABOR",
        },
        "Annex 2 Summary Category": {  # alternate spelling
            "Macau Construction Costs": "H_CONSTRUCTION",
            "Macau Contruction Costs": "H_CONSTRUCTION",  # typo variant
            "Staff Cost": "H_LABOR",
        },
        "Category": {
            "Macau Construction costs": "H_CONSTRUCTION",
            "Staff Cost": "H_LABOR",
        },
        # 潜在调整事项-for database — values ending in "—設施建設" → construction
        "潜在调整事项-for database": {
            "—設施建設": "H_CONSTRUCTION",
            "不符合\"吸引外國客源\"定義的相關投資支出—設施建設": "H_CONSTRUCTION",
            "未在原投資計劃明確列示且其後不被認可的新增項目投資支出—設施建設": "H_CONSTRUCTION",
            "缺乏支持性文件的從非承批公司轉入承批公司的博彩項目投資支出—設施建設": "H_CONSTRUCTION",
            "計入2024年報告投資金額的採購預付款—設施建設": "H_CONSTRUCTION",
        },
    },
    "vml": {
        "會計科目分類": {
            "工程建設": "H_CONSTRUCTION",
            "外採服務成本": "H_PROFESSIONAL",
            "設施採購": "H_EQUIP",
            "內部附贈資源-食物": "H_FNB",
            "人工支出": "H_LABOR",
            "交通費用": "H_COMP_OTHER",
            "內部附贈資源-其他": "H_COMP_OTHER",
            "內部附贈資源-客房": "H_HOTEL_ROOM",
            "器具採購": "H_EQUIP",
            "廣告費": "H_ADVERTISING",
            "媒體費用": "H_ADVERTISING",
            "拆除支出": "H_CONSTRUCTION",
            "贊助費用": "H_SPONSORSHIP",
            "租賃費": "H_LEASE",
            "授權費用": "H_LICENSE",
            "娛樂合約成本": "H_PERFORMER",
            "租賃折扣": "H_LEASE",
        },
        "Comp類型": {
            "食品飲料支出": "H_FNB",
            "Comp其他": "H_COMP_OTHER",
            "客房支出": "H_HOTEL_ROOM",
            "贈票支出": "H_COMP_TICKET",
        },
        "comp支出類型": {
            "食品飲料支出": "H_FNB",
            "客房支出": "H_HOTEL_ROOM",
            "Comp其他": "H_COMP_OTHER",
            "會場支出": "H_VENUE",
            "贈票支出": "H_COMP_TICKET",
        },
        "進一步分類": {
            "工程建設": "H_CONSTRUCTION",
            "專業服務費": "H_PROFESSIONAL",
            "人工成本": "H_LABOR",
            "設施採購": "H_EQUIP",
            "食品飲料支出": "H_FNB",
            "客房支出": "H_HOTEL_ROOM",
            "交通費": "H_COMP_OTHER",
            "Comp其他": "H_COMP_OTHER",
            "廣告費": "H_ADVERTISING",
            "媒體費用": "H_ADVERTISING",
            "推廣費": "H_ADVERTISING",
            "會場支出": "H_VENUE",
            "拆除支出": "H_CONSTRUCTION",
            "器具採購": "H_EQUIP",
            "娛樂表演合約成本": "H_PERFORMER",
            "授權費": "H_LICENSE",
            "贊助費": "H_SPONSORSHIP",
            "租賃費": "H_LEASE",
            "租賃折扣": "H_LEASE",
        },
        # NEW: 分類1.1 — almost identical values to 進一步分類
        "分類1.1": {
            "工程建設": "H_CONSTRUCTION",
            "設施採購": "H_EQUIP",
            "娛樂表演合約成本": "H_PERFORMER",
            "專業服務費": "H_PROFESSIONAL",
            "人工成本": "H_LABOR",
            "贊助費": "H_SPONSORSHIP",
            "授權費": "H_LICENSE",
            "媒體費用": "H_ADVERTISING",
            "租賃費": "H_LEASE",
            "食品飲料支出": "H_FNB",
            "器具採購": "H_EQUIP",
            "會場支出": "H_VENUE",
            "客房支出": "H_HOTEL_ROOM",
            "租賃折扣": "H_LEASE",
            "交通費": "H_COMP_OTHER",
        },
        # NEW: Payroll col — all values are clearly H_LABOR
        "Payroll": {
            "Capitalized Payroll": "H_LABOR",
            "Capitalized payroll": "H_LABOR",
            "Payroll & Benefit": "H_LABOR",
            "Payroll Allocation": "H_LABOR",
        },
        # NEW: Team col — Construction Accounting → H_CONSTRUCTION (408M)
        "Team": {
            "Construction Accounting": "H_CONSTRUCTION",
            "FP&A/Construction Accounting": "H_CONSTRUCTION",
        },
    },
    "melco": {
        "支出性質": {
            "營銷費用": "H_ADVERTISING",
            "影視製作及租賃成本": "H_LICENSE",
            "管理費": "H_PROFESSIONAL",
            "職工薪酬": "H_LABOR",
            "維修與保養費用": "H_MAINTENANCE",
            "一般用品採購": "H_EQUIP",
            "差旅費": "H_COMP_OTHER",
            "資產：新工作范圍": "H_CONSTRUCTION",
            "表演者費用": "H_PERFORMER",
            "職工福利": "H_LABOR",
            "Comp-差旅費": "H_COMP_OTHER",
            "資產：IT設備和軟件": "H_EQUIP",
            "資產：家具，配件和設備": "H_EQUIP",
            "資產：安全和監控設備": "H_EQUIP",
            "資產：博彩設備和軟件": "H_EQUIP",
            "贊助費": "H_SPONSORSHIP",
            "贊助費用": "H_SPONSORSHIP",
            "媒體費用": "H_ADVERTISING",
            "資產：租賃物業裝修": "H_LEASE",
            "Interco-人工成本-費用": "H_LABOR",
            "Interco-人工成本-收入": "H_LABOR",
            "水電費": "H_OTHER",
            "專業服務費": "H_PROFESSIONAL",
        },
        "Comp性質-CN（N/A為Net off及不適用），待確認kp識別，客戶未識別部分）": {
            "客房支出": "H_HOTEL_ROOM",
            "贈票支出": "H_COMP_TICKET",
            "食品飲料支出": "H_FNB",
            "其他": "H_COMP_OTHER",
            "場地租借": "H_VENUE",
        },
        "Comp性質-CN": {
            "食品飲料支出": "H_FNB",
            "客房支出": "H_HOTEL_ROOM",
            "演唱會門票支出": "H_COMP_TICKET",
            "其他": "H_COMP_OTHER",
            "場地租借": "H_VENUE",
        },
        # NEW: Comp性質分類-EN — English equivalent of Comp性質-CN
        "Comp性質分類-EN": {
            "Rental": "H_VENUE",
            "F&B": "H_FNB",
            "Rooms": "H_HOTEL_ROOM",
            "House of Dancing Water Ticket": "H_COMP_TICKET",
            "Arena Tickets": "H_COMP_TICKET",
            "Water Park Ticket": "H_COMP_TICKET",
            "F&B - Non Melco Operate": "H_FNB",
            "Golden Reel / Super Fun Zone / Kids CityTicket": "H_COMP_TICKET",
            "Limo": "H_COMP_OTHER",
            "Movie Ticket": "H_COMP_TICKET",
            "Trip ticket - Melco Paid to Vendor": "H_COMP_TICKET",
        },
        # NEW: Comp性質分類 (Chinese-named with English values)
        "Comp性質分類": {
            "Rooms": "H_HOTEL_ROOM",
            "Rental": "H_VENUE",
            "F&B": "H_FNB",
            "Arena Tickets": "H_COMP_TICKET",
            "Trip ticket - Melco Paid to Vendor": "H_COMP_TICKET",
            "F&B - Non Melco Operate": "H_FNB",
            "Water Park Ticket": "H_COMP_TICKET",
            "Golden Reel / Super Fun Zone / Kids CityTicket": "H_COMP_TICKET",
            "Movie Ticket": "H_COMP_TICKET",
            "Limo": "H_COMP_OTHER",
        },
        "支出性质-mapping": {
            "營銷費用": "H_ADVERTISING",
            "職工薪酬": "H_LABOR",
            "F&B成本": "H_FNB",
            "管理費": "H_PROFESSIONAL",
            "維修與保養費用": "H_MAINTENANCE",
            "資產：IT設備和軟件": "H_EQUIP",
            "差旅費": "H_COMP_OTHER",
            "資產：家具，配件和設備": "H_EQUIP",
            "資產：安全和監控設備": "H_EQUIP",
            "影視製作及租賃成本": "H_LICENSE",
            "資產：博彩設備和軟件": "H_EQUIP",
            "表演者費用": "H_PERFORMER",
            "一般用品採購": "H_EQUIP",
            "內部資源（Comp）": "H_COMP_OTHER",
            "資產：租賃物業裝修": "H_LEASE",
            "Interco-人工成本-費用": "H_LABOR",
            "Interco-人工成本-收入": "H_LABOR",
            "水電費": "H_OTHER",
        },
        # NEW: JL source — Payroll-related → H_LABOR, Ticketing → H_COMP_TICKET
        "JL source": {
            "Payroll Journal": "H_LABOR",
            "Payroll General Journal": "H_LABOR",
            "Payroll Accrual Journal": "H_LABOR",
            "Ticketing": "H_COMP_TICKET",
        },
        # NEW: KP識別人工 — Y means staff-cost-related → H_LABOR
        "KP識別人工": {
            "Y": "H_LABOR",
            "Y,staff cost分攤": "H_LABOR",
            "Melco - Staff cost come from Project Team": "H_LABOR",
        },
        "KP識別人工（Staff cost come from Project Team需Capex提供）": {
            "Y": "H_LABOR",
            "Staff cost come from Project Team": "H_LABOR",
            "Y,staff cost分攤": "H_LABOR",
        },
        # NEW: 人工成本分類 — labor classification col
        "人工成本分類（專門做投資項目的部門人工/涉及投資執行項目與其他經營業務之間進行分攤的相關執行部門人工）": {
            "Staff cost come from Project Team": "H_LABOR",
            "Y": "H_LABOR",
        },
    },
    "mgm": {
        # nature — 人工成本/Payroll → H_LABOR; 設施 → H_EQUIP
        "nature": {
            "人工成本": "H_LABOR",
            "Payroll": "H_LABOR",
            "設施": "H_EQUIP",
        },
        # item_type — Payroll → H_LABOR (others are too generic)
        "item_type": {
            "Payroll": "H_LABOR",
        },
        # ledger_l4 — high confidence H mappings from MGM's chart of accounts
        "ledger_l4": {
            "Standard: Promotional Expenses(L4)": "H_ADVERTISING",
            "Standard: Property And Equipment(L4)": "H_EQUIP",
            "Standard: Construction In Progress(L4)": "H_CONSTRUCTION",
            "Standard: Payroll(L4)": "H_LABOR",
            "Standard: PPE under finance lease(L4)": "H_LEASE",
            "Standard: Show Production(L4)": "H_PERFORMER",
            "Standard: Taxes & Licenses(L4)": "H_LICENSE",
        },
        # sub_type — granular L5 mapping
        "sub_type": {
            "Standard: Property Promotional Expenses (L5)": "H_ADVERTISING",
            "Standard: Casino Promotional Expenses(L5)": "H_ADVERTISING",
            "Standard: Advertising(L5)": "H_ADVERTISING",
            "Standard: Professional Fees(L5)": "H_PROFESSIONAL",
            "Standard: Salaries & Wages(L5)": "H_LABOR",
            "Property Promotional, Sponsorship Fees": "H_SPONSORSHIP",
            "70017 - Audio and Video equipment - Assets": "H_EQUIP",
            "70005 - Slot Machine & Equipment/Slot Signage/Slot machines - Assets": "H_EQUIP",
            "70008 - IT Equipment - Assets": "H_EQUIP",
            "52030 - Security & Surveillance Eq. - Assets": "H_EQUIP",
            "70070 - Building & improvements after Nov2017 - Assets": "H_CONSTRUCTION",
            "36519 - Cash Contribution:Donation/Sponsorship/Community Support": "H_SPONSORSHIP",
        },
    },
}


# Multi-column AND-logic rules (applied BEFORE single-col MAPPINGS).
# Each rule: (col_a, val_a_set, col_b, val_b_set, target_H)
# Row matches when df[col_a] in val_a_set AND df[col_b] in val_b_set.
MULTI_COL_RULES: dict[str, list[tuple[str, set, str, set, str]]] = {
    "galaxy": [
        # 一級標簽=Comp AND 二級標簽 specifies H type
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"room", "Room"}, "H_HOTEL_ROOM"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"F&B"}, "H_FNB"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Venue"}, "H_VENUE"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Tickets", "AirTickets"}, "H_COMP_TICKET"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Service Charge"}, "H_COMP_OTHER"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Entertainment"}, "H_COMP_OTHER"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Casual Labor", "staff cost", "Salaries & Leave"}, "H_LABOR"),
        ("一級標簽", {"Comp", "comp"}, "二級標簽", {"Utilities"}, "H_OTHER"),
    ],
}


def apply_for_entity(ent: str, com: str, dry_run: bool = False) -> dict:
    interim = Path(f"data/{ent}/interim")
    parquet = interim / f"{com}_raw.parquet"
    if not parquet.exists():
        return {"entity": ent, "error": "raw parquet missing"}

    mappings = MAPPINGS.get(ent, {})
    if not mappings:
        return {"entity": ent, "skipped": "no mappings defined"}

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols_cfg = cfg.get("columns", {})
    ac_col = cols_cfg.get("account_code", "")
    ad_col = cols_cfg.get("account_desc", "")
    dn_col = cols_cfg.get("description", "")
    amt_col = cols_cfg.get("amount", "")
    jc_col = cols_cfg.get("job_code", "")

    print(f"\n[{ent}] reading raw parquet...", flush=True)
    df = pd.read_parquet(parquet)
    print(f"  rows={len(df):,}  cols={len(df.columns)}", flush=True)

    if not (ac_col in df.columns and ad_col in df.columns and dn_col in df.columns and amt_col in df.columns):
        return {"entity": ent, "error": "missing required cols for signature building"}

    # Pre-compute signature per row
    from kpi.lib.text import normalize_description
    df["_acct_code"] = df[ac_col].astype("string").fillna("").str.strip()
    df["_acct_desc"] = df[ad_col].astype("string").fillna("").str.strip()
    df["_desc_norm"] = df[dn_col].apply(normalize_description)
    if jc_col and jc_col in df.columns:
        df["_job_code"] = df[jc_col].astype("string").fillna("").str.strip()
        df["_sig"] = df["_acct_code"] + "|" + df["_acct_desc"] + "|" + df["_desc_norm"] + "|" + df["_job_code"]
    else:
        df["_sig"] = df["_acct_code"] + "|" + df["_acct_desc"] + "|" + df["_desc_norm"]
    df["_amt_abs"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0).abs()

    # For each row, derive target H from manual labels (precedence: multi-col then single-col).
    target_h = pd.Series([None] * len(df), index=df.index, dtype="object")
    mapping_hit_stats: dict[str, int] = defaultdict(int)

    # ── STEP 1: Multi-column AND-logic rules (highest priority — most specific) ──
    multi_rules = MULTI_COL_RULES.get(ent, [])
    for col_a, vals_a, col_b, vals_b, h in multi_rules:
        if col_a not in df.columns or col_b not in df.columns:
            print(f"  ⚠️  multi-col rule skipped: '{col_a}' or '{col_b}' not in parquet", flush=True)
            continue
        a_match = df[col_a].astype(str).fillna("").isin(vals_a)
        b_match = df[col_b].astype(str).fillna("").isin(vals_b)
        mask = (target_h.isna()) & a_match & b_match
        n = int(mask.sum())
        if n:
            target_h.loc[mask] = h
            mapping_hit_stats[f"[multi] {col_a}∈{vals_a} & {col_b}∈{vals_b} → {h}"] += n

    # ── STEP 2: Single-column rules (lower priority — only if multi-col didn't catch it) ──
    for col, value_map in mappings.items():
        if col not in df.columns:
            print(f"  ⚠️  col '{col}' not in parquet, skip", flush=True)
            continue
        col_vals = df[col].astype(str).fillna("")
        for val, h in value_map.items():
            mask = (target_h.isna()) & (col_vals == val)
            n = int(mask.sum())
            if n:
                target_h.loc[mask] = h
                mapping_hit_stats[f"{col}={val}→{h}"] += n

    df["_target_h"] = target_h
    labeled = df.dropna(subset=["_target_h"])
    print(f"  rows with manual label: {len(labeled):,}/{len(df):,} ({100*len(labeled)/len(df):.1f}%)")
    print(f"  label hit breakdown:")
    for k, n in sorted(mapping_hit_stats.items(), key=lambda x: -x[1])[:15]:
        print(f"    {k:<70} {n:,} rows")

    if len(labeled) == 0:
        return {"entity": ent, "labeled_rows": 0}

    # Per-sig majority H (amount-weighted)
    sig_h_amt: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for _, r in labeled.iterrows():
        sig_h_amt[r["_sig"]][r["_target_h"]] += r["_amt_abs"]

    sig_overrides = {}
    for sig, h_amt in sig_h_amt.items():
        # Pick H with max weighted amount
        best_h = max(h_amt, key=h_amt.get)
        sig_overrides[sig] = best_h

    print(f"  → {len(sig_overrides):,} unique sigs to override")

    if dry_run:
        # Show distribution
        from collections import Counter
        h_dist = Counter(sig_overrides.values())
        print(f"  H distribution preview:")
        for h, n in h_dist.most_common():
            print(f"    {h:<20} {n:,} sigs")
        return {"entity": ent, "sigs": len(sig_overrides), "dry_run": True}

    # Apply to feedback.xlsx
    fb_path = Path(f"data/{ent}/output/{com}_feedback.xlsx")
    if fb_path.exists():
        fb_df = pd.read_excel(fb_path)
    else:
        fb_df = pd.DataFrame(columns=["signature", "correct_horizontal", "notes"])
    for c in ("signature", "correct_horizontal", "notes"):
        if c not in fb_df.columns:
            fb_df[c] = ""
    fb_df["signature"] = fb_df["signature"].astype(str).fillna("")
    fb_df["correct_horizontal"] = fb_df["correct_horizontal"].astype(str).fillna("")
    fb_df["notes"] = fb_df["notes"].astype(str).fillna("")

    n_updated = 0
    n_added = 0
    new_rows = []
    for sig, h in sig_overrides.items():
        if sig in fb_df["signature"].values:
            fb_df.loc[fb_df["signature"] == sig, "correct_horizontal"] = h
            fb_df.loc[fb_df["signature"] == sig, "notes"] = "[manual-label-override] from project team raw label"
            n_updated += 1
        else:
            new_rows.append({
                "signature": sig,
                "correct_horizontal": h,
                "notes": "[manual-label-override] from project team raw label",
            })
            n_added += 1
    if new_rows:
        fb_df = pd.concat([fb_df, pd.DataFrame(new_rows)], ignore_index=True)

    fb_path.parent.mkdir(parents=True, exist_ok=True)
    fb_df.to_excel(fb_path, index=False)
    print(f"  ✅ wrote {fb_path.name}  (updated: {n_updated:,}, added: {n_added:,})")
    return {"entity": ent, "sigs": len(sig_overrides), "updated": n_updated, "added": n_added}


def main():
    parser = argparse.ArgumentParser()
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--entity", choices=list(ENTITIES))
    g.add_argument("--all", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    targets = list(ENTITIES.items()) if args.all else [(args.entity, ENTITIES[args.entity])]
    for ent, com in targets:
        r = apply_for_entity(ent, com, dry_run=args.dry_run)
        if "error" in r:
            print(f"  ❌ {r['error']}")


if __name__ == "__main__":
    main()
