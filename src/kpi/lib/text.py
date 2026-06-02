"""Text normalization for building row signatures."""
from __future__ import annotations

import re

# ── V_NOISE: bookkeeping/AP/control-char strippers ─────────────────────────
# These patterns collapse sigs that differ ONLY by accounting-plumbing prefix
# (A-ARIBA\, RCL-ARIPO\, VEN HOTEL AP TRADE LIABILITY, **IREF_, control-chars).
# They describe HOW the entry was posted (accrual / reclass / Ariba PO),
# not WHAT the transaction is — so stripping preserves classification info.

# N1: ASCII control bytes + Excel-escaped hex (_x001A_ etc). Excel exports
# from Chinese-encoded files leave SUB control chars as either literal 0x1A
# or as the text "_xNNNN_". Strip both.
_CONTROL_CHARS = re.compile(r"[\x00-\x1f]+|_x[0-9a-fA-F]{4}_")

# N2: Accrual prefix — A-ARIBA\, A-ARIBAOP\, A-CIS\, A-CISO\, A-OTHER\, A-VTL\, A-INV\
_ACCRUAL_PREFIX = re.compile(
    r"\b(?:A-(?:ARIBA(?:OP)?|CIS[O]?|OTHER|VTL|INV))\\\s*",
    re.IGNORECASE,
)

# N3: Reclass / allocation / FX / adjustment prefix — all bookkeeping operations
_RECLASS_PREFIX = re.compile(
    r"\b(?:RCL-(?:ARIPO|ARIINV|ARINV)"
    r"|ALLO-(?:ARIPO|ARIINV|ARINV)"
    r"|RECLASS|REALIZED|AMORTIZED|FOREX|REBATES|ADJUSTMENT"
    r"|ADJ-ALLOC|MALL\s+ALLOC|ALLOCATION|ARI-FA/EXP|PYRCLS|RCLS)\\\s*",
    re.IGNORECASE,
)

# N4: AP trade liability prefix variants (VEN HOTEL / VTL / VEN / HOTEL / bare)
_AP_PREFIX = re.compile(
    r"\b(?:VEN\s+HOTEL\s+|VTL\s+|VEN\s+|HOTEL\s+)?AP\s+TRADE\s+LIABILITY\b\s*",
    re.IGNORECASE,
)

# N5: IREF marker (the **IREF_ before invoice ref code)
_IREF_MARKER = re.compile(r"\*\*\s*IREF[_\s]+", re.IGNORECASE)

# N6: HOTEL REVENUE (MOP) prefix variants (with/without space)
_HOTEL_REV_PREFIX = re.compile(
    r"\bHOTEL\s+REVENUE\s*\(?\s*MOP\s*\)?\s*\d*\s*",
    re.IGNORECASE,
)

# N7: Accrual / Adj / Provision / 預提 prefix (SJM/Wynn common bookkeeping leaders)
_ACCR_LEAD = re.compile(
    r"^(?:Adj\.?\s+|Accr\.?\s+|Accrual\s+|Provision\s+of\s+|預提\s*|預付\s*|Reaccrual\s+|Reaccr\.?\s+|Rcls?\.?\s+|Reclass\s+)",
    re.IGNORECASE,
)

# N8: From <entity> intercompany prefix (e.g., "From GLP", "From GLPH", "From XXX -")
_FROM_PREFIX = re.compile(
    r"^(?:From\s+[A-Z]{2,6}\b\s*-?\s*)",
)

# N9: Netoff / Net-off (Wynn-style — paired transaction markers)
_NETOFF = re.compile(r"\bNet[-\s]?off\b", re.IGNORECASE)

# N10: Cost allocation prefix (e.g. "PCM Cost Allocation - " in SJM)
_COST_ALLOC = re.compile(
    r"\b(?:PCM\s+Cost\s+Allocation|Cost\s+Allocation|Cost\s+reallocation|IC[-_\s]?Recharge)\s*-?\s*",
    re.IGNORECASE,
)

# N11: PROJECT_XXX / PROJECT GL-XXX prefix (SJM SAP project codes)
_PROJECT_CODE = re.compile(
    r"\b(?:PROJECT_GL[-_][A-Z]+[-_]?[A-Z0-9]+|PROJECT\s+GL[-_][A-Z]+(?:\s+for)?)\b",
    re.IGNORECASE,
)


_DATE_PATTERNS = [
    re.compile(r"\b\d{1,2}[-/\.]\w{1,3}[-/\.]\d{2,4}\b"),
    re.compile(r"\b\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}\b"),
    re.compile(r"\b\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2,4}\b"),
    re.compile(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[-/\s]?\d{2,4}\b", re.IGNORECASE),
    re.compile(r"\b\d{4}\b"),
]
_PO_PATTERNS = [
    re.compile(r"#\s*\d{4,}"),
    re.compile(r"\b(?:PO|GR|FA|DD|SO)[#\-\s]*\d{4,}\b", re.IGNORECASE),
    re.compile(r"\(\d{6,}\)"),
]

# ── General time-token strippers ────────────────────────────────────────────
# Collapses recurring sigs that differ ONLY by time tokens — the canonical
# "Rent - Jan / Rent - Feb / ..." style. Additive vs V0: sigs that don't match
# any time-token are unchanged (V0 hash preserved → cache stays hit).
_MONTH_TOKEN = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec"
    r"|January|February|March|April|June|July|August|September|October|November|December)\b",
    re.IGNORECASE,
)
_CN_MONTH = re.compile(r"\d{0,4}年?\d{1,2}月份?")
_CN_QUARTER = re.compile(r"\d{0,4}年?第?[一二三四1234]季度?")
_EN_QUARTER = re.compile(r"\bQ[1-4](?:[-/\s]\d{2,4})?\b", re.IGNORECASE)
_FY_TOKEN = re.compile(r"\bFY[-/\s]?\d{2,4}\b", re.IGNORECASE)
_WEEK_DAY_PERIOD = re.compile(r"\b(?:Week|Wk|Day|Period|P)[-/\s]?\d{1,3}\b", re.IGNORECASE)
_YTD_TOKEN = re.compile(r"\b(?:YTD|MTD|QTD|YOY|MOM)\b", re.IGNORECASE)

# N12: Letter+digits project/zone codes (Galaxy/SJM E021, D306, J018, H001,
# P3D, GL-ADW-M02, etc). These differ between rows of the same actual transaction
# type. Collapse to <CODE>. Conservative: only 1-2 letters + 2-5 digits.
_LETTER_DIGIT_CODE = re.compile(r"\b[A-Z]{1,2}\d{2,5}(?:[A-Z]{1,2})?\b")

# N13: Floor designators (8/F, 7/F, 31F, 10F, UG/F, B1/F)
_FLOOR = re.compile(r"\b(?:UG|B\d)?/?\d{1,3}[/-]?F\b", re.IGNORECASE)

# N14: SAP/SJM/PR/PO/CEA doc numbers (Doc#SJM2025500001267, PR41002058, PO8500002196)
_DOC_NUM = re.compile(
    r"\b(?:SJM|PR|PO|CEA|GR|FA|DOC|INV|CMVPD|TPMCAR|VIK|MIN)#?[A-Z]*[-/_]?\d{4,}\b",
    re.IGNORECASE,
)

# N15: Year-range tokens (24&25, 2024&2025, YR1, YR1&2)
_YEAR_RANGE = re.compile(r"\b(?:\d{2,4}[&-]\d{2,4}|YR[1-9](?:&[1-9])?)\b", re.IGNORECASE)

# N16: Vendor V-codes (V#135269, V9876, V1449612)
_VENDOR_CODE = re.compile(r"\bV[#]?\d{3,}\b")

# N17: % deposit markers (50%Dep, 100%Bal, R50%Dep, st50%, R100%D)
_PCT_MARKER = re.compile(r"\b[Rr]?\d{1,3}%[A-Za-z]+\b")

_LARGE_NUM = re.compile(r"\b\d{2,}\b")
_SEGMENT_NUM = re.compile(r"([/\-])\d+(?=\b|$)")
_INWORD_NUM = re.compile(r"\d+(?=[A-Za-z])|(?<=[A-Za-z])\d+")
_WHITESPACE = re.compile(r"\s+")
_PUNCT = re.compile(r"[\*\-_]{2,}")


def normalize_description(s: object) -> str:
    if s is None:
        return ""
    txt = str(s).strip()
    if txt in ("", "-", "nan", "None", "NaN"):
        return ""
    # V_NOISE — strip bookkeeping plumbing FIRST so it doesn't pollute downstream sigs.
    # Each replacement is conditional (no match = no change), so sigs without noise
    # prefix hash identically to V0/V1 → cache stays hit for those.
    txt = _CONTROL_CHARS.sub("", txt)
    txt = _ACCRUAL_PREFIX.sub("<JE> ", txt)
    txt = _RECLASS_PREFIX.sub("<RCL> ", txt)
    txt = _AP_PREFIX.sub("<AP> ", txt)
    txt = _IREF_MARKER.sub("<IREF> ", txt)
    txt = _HOTEL_REV_PREFIX.sub("<HREV> ", txt)
    txt = _ACCR_LEAD.sub("<ACCR> ", txt)
    txt = _FROM_PREFIX.sub("<FROM> ", txt)
    txt = _NETOFF.sub("<NETOFF>", txt)
    txt = _COST_ALLOC.sub("<ALLOC> ", txt)
    txt = _PROJECT_CODE.sub("<PRJ>", txt)

    for p in _DATE_PATTERNS:
        txt = p.sub("<DATE>", txt)
    for p in _PO_PATTERNS:
        txt = p.sub("<ID>", txt)
    # General time-token collapse (post-DATE, since DATE catches Jan-25 first).
    # Standalone Jan/Feb/.../Dec and English/Chinese quarter + FY + Week tokens.
    txt = _MONTH_TOKEN.sub("<MONTH>", txt)
    txt = _CN_MONTH.sub("<MONTH>", txt)
    txt = _CN_QUARTER.sub("<QTR>", txt)
    txt = _EN_QUARTER.sub("<QTR>", txt)
    txt = _FY_TOKEN.sub("<FY>", txt)
    txt = _WEEK_DAY_PERIOD.sub("<PERIOD>", txt)
    txt = _YTD_TOKEN.sub("<TD>", txt)

    # Round 2 V_NOISE — apply BEFORE generic <NUM> normalizer so structured
    # codes (PR41002058, V#135269, 50%Dep) get specific tokens not generic <NUM>.
    txt = _DOC_NUM.sub("<DOC>", txt)
    txt = _VENDOR_CODE.sub("<VEND>", txt)
    txt = _PCT_MARKER.sub("<PCT>", txt)
    txt = _YEAR_RANGE.sub("<YR>", txt)
    txt = _FLOOR.sub("<FLR>", txt)
    txt = _LETTER_DIGIT_CODE.sub("<CODE>", txt)

    txt = _LARGE_NUM.sub("<NUM>", txt)
    txt = _SEGMENT_NUM.sub(r"\1<NUM>", txt)
    txt = _INWORD_NUM.sub("<NUM>", txt)
    txt = _PUNCT.sub(" ", txt)
    txt = _WHITESPACE.sub(" ", txt)
    return txt.strip().lower()


def signature_key(account_code: object, account_desc: object, description: object) -> str:
    a = "" if account_code is None else str(account_code).strip()
    ad = "" if account_desc is None else str(account_desc).strip()
    d = normalize_description(description)
    return f"{a}|{ad}|{d}"
