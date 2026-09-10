#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
layout.py — 報告 pptx 版式引擎（KPMG house style）。全部版式常數集中喺呢度，
make_report / render_review_table_pptx 只負責「派數字」，唔再各自砌 formatting。

版式對齊 MGM 2025 final 報告 pptx（10.83 x 7.5 in）。2026-09-07 由 `inspect_pptx --fmt`
逐 109 版實測校準，唔再係掃描件目測值。每版：頂 nav → 「章節 | 子題」(y=0.50, 12pt)
→ navy 導語 (y=0.70) → 內容【固定由 y=1.55 起】→ 資料來源 (y=6.72) → footer (y=6.85)。

★ 兩種表要分清楚：
  · 敘述框表（事項描述／附件工作範圍／現場走訪）＝ 原報告 native，表頭底 00338D、
    全表 9pt、框線 00338D 虛線。
  · 數字表 ＝ 原報告係 Tableau 截圖，冇 native 版可抄；表頭沿用項目組指定 1E49E2／098E7E。
    （user 2026-09-07 拍板：我哋出 native 表，唔複製截圖。）

顏色跟 KPMG Visual identity overview（品牌手冊）：
    Primary   KPMG Blue 00338D｜Medium Blue 005EB8｜Light Blue 0091DA
    Secondary Violet 483698｜Purple 470A68｜Light Purple 6D2077｜Green 00A3A1
表格 tint（由 KPMG Blue 派生）：section EEF1F8、小計 D9E1F2、總計 BDD7EE、格線 BFBFBF。
字體：PowerPoint 用 Arial（品牌手冊指定）行數字/英文，中文用微软雅黑。

★ 高度控制：PowerPoint 會自動長高 row 去就內容 → 純靠 row 數分頁一定爆版。
  呢度用 est_lines() 逐 cell 估 wrap 行數 → 逐 row 定實高度 → 按【累積高度】分頁，
  所以永遠唔會超出 slide 底（單項審查「超出 border」嘅正解）。
"""
import re

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ── KPMG 品牌色 ──────────────────────────────────────────────────────────
NAVY = RGBColor(0x00, 0x33, 0x8D)          # KPMG Blue（表頭 / 標題 / 導語）
MBLUE = RGBColor(0x00, 0x5E, 0xB8)         # Medium Blue
LBLUE = RGBColor(0x00, 0x91, 0xDA)         # Light Blue
VIOLET = RGBColor(0x48, 0x36, 0x98)
PURPLE = RGBColor(0x47, 0x0A, 0x68)
LPURPLE = RGBColor(0x6D, 0x20, 0x77)
GREEN = RGBColor(0x00, 0xA3, 0xA1)
# 表格 tint（KPMG Blue 派生；報告通篇用呢 3 級）
SECFILL = RGBColor(0xEE, 0xF1, 0xF8)       # 範疇 section 行
SUBTOT = RGBColor(0xD9, 0xE1, 0xF2)        # 小計
TOTAL = RGBColor(0xBD, 0xD7, 0xEE)         # 總計
BORDER = "BFBFBF"                          # 格線（srgbClr hex）
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x22, 0x22, 0x22)           # 內文黑
GREY = RGBColor(0x59, 0x59, 0x59)          # 註 / 資料來源
DARK = RGBColor(0x0C, 0x23, 0x3C)          # 封面 / 章節分隔深底

# ── 原報告主題色盤（2026-09-10 解 theme1.xml clrScheme + slideMaster clrMap 得出）──
#   之前逐版量度出嚟嘅色好多係「THEME:xxx」索引，換唔到 RGB；解開之後先知係咩。
#   三個 slideMaster 嘅色盤【完全一樣】，所以唔使分。
TH_BG1 = RGBColor(0xFF, 0xFF, 0xFF)        # bg1 → lt1
TH_TX1 = RGBColor(0x00, 0x00, 0x00)        # tx1 → dk1　正文黑（原報告內文係純黑，唔係 #222222）
TH_BG2 = RGBColor(0xE5, 0xE5, 0xE5)        # bg2 → lt2　breadcrumb 非當前章
TH_TX2 = NAVY                              # tx2 → dk2 = #00338D
CRUMB_OFF = TH_BG2                         # breadcrumb 非當前章：原報告實測 8pt #E5E5E5（好淺）
LGREY = RGBColor(0x8C, 0x8C, 0x8C)         # 頁腳（版權／文檔分類）
#   ⚠ 之前 breadcrumb 同頁腳共用 LGREY。原報告 breadcrumb 係 #E5E5E5，但頁腳未量到，
#     一齊改就會令版權行接近睇唔見 → 分開兩個常數，頁腳維持 #8C8C8C。

# 負數用括號表示（KPMG palette 冇紅色）→ 唔另外上色。想要紅色改呢個做 RGBColor(0xC0,0,0)。
NEG_COLOR = None

# 字體：由公司 template theme 實測（inspect_pptx --spec，2026-08-12）
#   majorFont latin=KPMG Bold  ea=Microsoft YaHei ｜ minorFont latin=Arial  ea=Microsoft YaHei
# ⚠ python-pptx 嘅 font.name 只寫 <a:latin>，中文字要寫 <a:ea>，否則 PowerPoint 會用 theme 預設。
FONT_CN = "Microsoft YaHei"                 # <a:ea>（中文）
FONT_NUM = "Arial"                          # <a:latin>（數字/英文）
FONT_HEAD = "KPMG Bold"                     # 標題 latin（中文照樣行 ea）

# ── 字號（集中一處；inspect_pptx --fonts 可印出嚟同真報告逐個位置對）────────
# ★ 由項目組真報告實測（inspect_pptx --fonts --range 10-63，2026-08-12）：
#   章節｜子題 12pt（202 runs 單一值）｜導語 13-14pt｜內文 body 9pt（575 runs）
#   資料來源/註 7pt｜native 表格 9pt（佢哋大表多數係 Tableau 截圖，所以表身抽到 0）
#   我哋原本細成 1.5 倍（8.5/8.5/8.0/6.0/6.5）→ 全部校準。
#   ⚠ 字大咗，同一版塞唔到咁多 → 自動分頁會出多幾版，呢個【正合】報告嘅版數
#     （報告 1.3 有 4 版、1.4 有 3 版，我哋之前一版塞晒）。
SZ_CRUMB = 8.0      # ① 頂 breadcrumb —— 原報告實測 8pt（非當前 #E5E5E5、當前 8pt 粗 #00338D）
SZ_TITLE = 12.0     # ② 章節｜子題
SZ_HEAD = 12.0      # ③ 導語 strapline（1.x／2.x／3.x 全部 12pt；4.x／附件見 HEAD_SIZE）
SZ_BODY = 9.0       # ④ 內文 body（prose 段落）
SZ_BODY_HEAD = 9.5  # ④ 內文小標題
SZ_TBL = 9.0        # 表身（final 報告 native 表【全部】9pt —— 2026-09-07 --fmt 實測）
SZ_TBL_HDR = 9.0    # 表頭（同樣 9pt，白色粗體、底 00338D）
SZ_TBL_WIDE = 6.0   # 表身（>16 欄嘅大表，9pt 塞唔落 18 欄）
SZ_TBL_MID = 7.5    # 表身（11-14 欄）
SZ_CAPTION = 7.5    # 表頂 navy caption bar
SZ_NOTE = 7.0       # ⑤ 資料來源 / 註
SZ_FOOT = 6.0       # ⑥ footer 版權
SZ_PAGE = 9.0       # ⑥ 頁碼

SLIDE_W = 10.83                             # 報告 slide 尺寸（scan 量度確認）
SLIDE_H = 7.5

# 版面錨點（吋）
MARGIN = 0.53          # template 實測：內容 x=0.53、闊 9.76（--spec）
COL_GAP = 0.21         # template 兩欄 gap 實測
# ★ 以下 y 全部由 final 報告 109 版逐版量（inspect_pptx --fmt，2026-09-07），唔再係 scan 估值。
CRUMB_Y = 0.10         # 頂 nav 條（原報告係 UpSlide GROUP y=0.10 h=0.21）
SUBTITLE_Y = 0.50      # 「章 | 節」 12pt（原報告 x=0.53 y=0.50 w=9.76 h=0.17）
HEAD_Y = 0.70          # 導語 strapline（原報告 y=0.62~0.80，h≈0.84）
CONTENT_Y = 1.55       # ★ 內容起始線 —— 原報告【每一版】表／文字都由 1.53~1.58 開始
FOOT_Y = 6.85          # 版權行（原報告 © y=6.85、頁碼 y=6.84、文檔分類 y=6.88）
CONTENT_BOTTOM = 6.72  # 資料來源 pin 底時嘅 y（原報告 6.72~6.74）

# 導語字號按章節（原報告實測）：概述／期後／主要發現 12pt，4.x 大細唔同，附件最大。
HEAD_SIZE = {0: 12.0, 1: 12.0, 2: 12.0, 3: 18.0, 4: 18.0, 5: 24.0}

# ── 表左＋敘述右嘅兩欄版 ────────────────────────────────────────────────
# 原報告實測（s10/s20/s24）：表 4.47~4.60in、gap 0.14、敘述 5.09~5.19in
#   → 表約佔 46%、敘述 52%。我哋之前 W*0.60 = 6.50 / 3.05，表霸咗成版，
#     敘述擠喺右邊窄條，係「似但唔一樣」最搶眼嗰項。
# ⚠ 但原報告嗰啲表係 Tableau 截圖（可任意縮細唔理可讀性），我哋係 native 表，
#   欄多就一定要闊啲先讀得到 → ≤6 欄跟足原報告，再多就按欄數放寬。
SPLIT_TBL_W = 4.55      # 表闊（≤6 欄）
SPLIT_GAP = 0.14        # 表同敘述之間


def tbl_font(ncol):
    """表身字號按欄數分三級。原本得 9pt／6pt 兩級，15 欄嘅表用緊 9pt，
    塞唔落 → 23 年單項審查表出到 12 版（原報告 7 版）。"""
    if ncol <= 10:
        return SZ_TBL
    return SZ_TBL_MID if ncol <= 14 else SZ_TBL_WIDE


def split_left(ncol):
    """兩欄版嘅表闊（吋）。"""
    return min(6.50, SPLIT_TBL_W + max(0, ncol - 6) * 0.35)


# 原報告【冇】表頂 navy 標題條，表直接由 CONTENT_Y 開始。想要返就改 True。
SHOW_TABLE_CAPTION = False

SECTIONS = ["2025年度投資計劃執行情況概述", "過往年度投資計劃在2025年繼續執行的審查跟進",
            "本年度審查工作的主要發現", "其他信息", "投資計劃執行報告的六項KPI分析", "附件"]

_CN_RE = None


def set_ea(run_or_font, ea=None):
    """寫 <a:ea>（中文字體）—— python-pptx 只寫 <a:latin>，唔寫 ea 中文會跌返 theme 預設。
    OOXML 次序：… latin, ea, cs …，所以要 insert 喺 latin 之後。"""
    f = getattr(run_or_font, "font", run_or_font)
    try:
        rPr = f._rPr
    except AttributeError:
        return
    if rPr is None:
        return
    el = rPr.find(qn("a:ea"))
    if el is None:
        el = rPr.makeelement(qn("a:ea"), {})
        lat = rPr.find(qn("a:latin"))
        (lat.addnext(el) if lat is not None else rPr.append(el))
    el.set("typeface", ea or FONT_CN)


# openpyxl 讀 Excel 有啲 cell 帶住 _x0000_ 呢類轉義（原檔有控制字元），照抄落 pptx 會見到
#   「在泰國曼_xFFFF_」咁嘅怪字。喺【所有文字最後出口】清一次，唔使逐個 loader 補。
_ESC = re.compile(r"_x[0-9A-Fa-f]{4}_")


def scrub(t):
    return _ESC.sub("", str(t))


def setfont(run, size, *, bold=False, italic=False, color=None, heading=False, latin=None):
    """一次過設 size/bold/color + <a:latin> + <a:ea>（跟 template theme）。
    順手清走 Excel 轉義殘留（_xFFFF_ 之類）——每個 run 一定會行過呢度。"""
    if "_x" in run.text:
        run.text = scrub(run.text)
    f = run.font
    f.size = Pt(size); f.bold = bold; f.italic = italic
    if color is not None:
        f.color.rgb = color
    f.name = latin or (FONT_HEAD if heading else FONT_NUM)
    set_ea(f)
    return run


def _is_cn(ch):
    return "⺀" <= ch <= "鿿" or "＀" <= ch <= "￯" or "　" <= ch <= "〿"


def has_cn(s):
    return any(_is_cn(c) for c in str(s))


def text_w(s, size):
    """估文字闊度（pt）：中文/全形 ≈ 1 em、英數 ≈ 0.52 em。"""
    w = 0.0
    for c in str(s):
        w += size * (1.0 if _is_cn(c) else 0.52)
    return w


def est_lines(s, col_w_in, size, margin_in=0.06):
    """估 wrap 行數（col_w_in = 欄闊吋）。認 \\n 明碼換行。"""
    avail = max((col_w_in - margin_in) * 72.0, 6.0)
    n = 0
    for seg in str(s).split("\n"):
        n += max(1, -(-text_w(seg, size) // avail))     # ceil
    return int(n)


def row_h(cells, widths, size, pad_in=0.045, min_h=0.155):
    """一行嘅需要高度（吋）＝ 最多 wrap 行數 × 行距 + 上下 padding。"""
    lines = 1
    for txt, w in zip(cells, widths):
        lines = max(lines, est_lines(txt, w, size))
    return max(min_h, lines * size * 1.24 / 72.0 + pad_in)


# ── 基本元件 ─────────────────────────────────────────────────────────────
def size_of(prs):
    return prs.slide_width / 914400.0, prs.slide_height / 914400.0


def blank(prs):
    lay = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    return prs.slides.add_slide(lay)


def _tb(slide, x, y, w, h, wrap=True):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.margin_left = tf.margin_right = Emu(0)
    tf.margin_top = tf.margin_bottom = Emu(0)
    return box


def put(slide, x, y, w, h, text, *, size=8, bold=False, color=INK, align=PP_ALIGN.LEFT,
        italic=False, font=None, wrap=True):
    """一行/一段文字框。wrap=False 用喺一定要一行嘅嘢（breadcrumb 頁籤）。"""
    box = _tb(slide, x, y, w, h, wrap)
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    p.font.size = Pt(size); p.font.name = FONT_NUM; set_ea(p.font)   # 空段落唔好跌返 theme 預設
    p._p.get_or_add_endParaRPr().set("sz", str(int(round(size * 100))))
    r = p.add_run(); r.text = str(text)
    setfont(r, size, bold=bold, italic=italic, color=color, latin=font)
    return box


BAND = RGBColor(0xF2, 0xF2, 0xF2)          # breadcrumb 淺灰底（scan 頂部 banner）


def _rect(slide, x, y, w, h, fill, line=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line; sh.line.width = Pt(0.5)
    sh.shadow.inherit = False
    return sh


def _name(shape, nm):
    try:
        shape.name = nm
    except Exception:      # noqa: BLE001 — 改唔到名只係少咗 hyperlink，唔好炸咗成個 build
        pass


def breadcrumb(slide, W, active=0, entity="MGM"):
    """頂 nav（對 scan p-23 放大）：白底、頁籤用「｜」分隔，當前頁籤 navy 粗體、其餘淺灰，
    右邊 entity + ◀ ⌂ ▶ 三粒圓掣。shape 改名做 nav:* ，wire_nav() 事後駁內部 hyperlink。"""
    # ★ 原報告冇 ◀⌂▶ 圓掣（實測：淨係一條 UpSlide 導航條 x=0.54 w=8.62 + 右邊 entity
    #   x=9.20）→ 圓掣已拎走，只保留頁籤本身嘅內部跳頁。
    x0 = MARGIN - 0.23
    ex = 9.20                                              # entity 位置對正原報告
    put(slide, ex, CRUMB_Y, 1.10, 0.18, entity, size=SZ_CRUMB, bold=True,
        color=INK, align=PP_ALIGN.LEFT)
    sep, avail = " ｜ ", (ex - 0.10) - x0
    # 鬆位【只畀粗體嗰一格】—— 之前成排 ×1.08，8pt 之下估到 9.26in > 可用 8.80in，
    # 於是自動縮成 7.6pt。但原報告個導航條實測 8.62in、字係 8pt，我哋唔加 fudge
    # 估出嚟係 8.66in（同佢一致）＝ 裝得落。粗體只有當前嗰格，單獨放鬆就夠。
    widths = [text_w(t, SZ_CRUMB) * (1.06 if i == active else 1.0) / 72.0
              for i, t in enumerate(SECTIONS)]
    sw = text_w(sep, SZ_CRUMB) / 72.0
    scale = min(1.0, avail / (sum(widths) + sw * (len(SECTIONS) - 1)))
    x = x0
    for i, t in enumerate(SECTIONS):
        if i:
            put(slide, x, CRUMB_Y, sw * scale + 0.03, 0.18, sep, size=SZ_CRUMB * scale,
                color=CRUMB_OFF, wrap=False)
            x += sw * scale
        w = widths[i] * scale
        _name(put(slide, x, CRUMB_Y, w + 0.05, 0.18, t, size=SZ_CRUMB * scale, wrap=False,
                  bold=(i == active), color=NAVY if i == active else CRUMB_OFF), f"nav:sec{i}")
        x += w


def _hlink(shape, rid):
    """畀 shape 內所有 run 加內部跳頁 hyperlink（a:hlinkClick + ppaction://hlinksldjump）。
    hlinkClick 喺 CT_TextCharacterProperties 排 latin/ea 之後 → append 就啱序。"""
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            rPr = r._r.get_or_add_rPr()
            h = rPr.makeelement(qn("a:hlinkClick"),
                                {qn("r:id"): rid, "action": "ppaction://hlinksldjump"})
            rPr.append(h)


def wire_nav(prs, sec_slide=None, home=0):
    """全部 slide 砌完（連目錄插咗、重排咗）之後至駁：◀/▶ = 上/下頁、⌂ = 目錄、
    頁籤 = 該章分隔頁。sec_slide = {章 index: slide index}。"""
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    slides = list(prs.slides)
    n = len(slides)
    for i, s in enumerate(slides):
        tgt = {"nav:prev": max(i - 1, 0), "nav:next": min(i + 1, n - 1), "nav:home": home}
        for k, v in (sec_slide or {}).items():
            tgt[f"nav:sec{k}"] = v
        for sh in s.shapes:
            j = tgt.get(sh.name)
            if j is None or j == i or not sh.has_text_frame:
                continue
            _hlink(sh, s.part.relate_to(slides[j].part, RT.SLIDE))


def footer(slide, W, H, page):
    """底：KPMG 字標 + 版權 + 文檔分類 + 頁碼。
    位置對 final 報告 master（© x=1.68 y=6.85、文檔分類 x=7.97 y=6.88、頁碼 x=9.73 y=6.84）。"""
    kb = _tb(slide, MARGIN, FOOT_Y, 0.7, 0.22)
    kr = kb.text_frame.paragraphs[0].add_run(); kr.text = "KPMG"
    setfont(kr, 11, bold=True, italic=True, color=NAVY)
    put(slide, 1.68, FOOT_Y, 5.74, 0.2,
        "© 2026畢馬威會計師事務所 — 澳門特別行政區合夥制事務所。版權所有，不得轉載。",
        size=SZ_FOOT, color=LGREY)
    put(slide, 7.97, FOOT_Y + 0.03, 1.66, 0.16, "文檔分類: 保密", size=SZ_FOOT, color=LGREY)
    if page is not None:
        put(slide, W - 1.15, FOOT_Y, 0.95, 0.2, str(page), size=SZ_PAGE, bold=True,
            color=NAVY, align=PP_ALIGN.RIGHT)


MAX_HEAD_H = 0.85      # 導語高度上限：HEAD_Y 0.70 + 0.85 = 1.55 ＝ 原報告內容起始線


def head_h(headline, W, hsize=SZ_HEAD):
    """導語需要嘅高度 + 實際字號（長就自動縮到 MAX_HEAD_H 為止）→ (h, size)。"""
    if not headline:
        return 0.06, hsize
    while hsize > 6.0:
        h = est_lines(headline, W - 2 * MARGIN, hsize) * hsize * 1.35 / 72.0
        if h <= MAX_HEAD_H:
            return h, hsize
        hsize -= 0.5
    return MAX_HEAD_H, hsize


def subsec_marker(slide, crumb):
    """畫布外嘅章節標記（原報告每版都有：x=0 y=-0.28 w=1.39 h=0.01，15pt）。
    UpSlide 讀呢個生成目錄；原報告主要發現／附件嗰批版根本冇麵包屑，節名淨係喺呢度。
    冇佢就砌唔到目錄，diff_report ① 亦對唔到版。"""
    sub = re.split(r"\s*[|｜]\s*", str(crumb or ""), maxsplit=1)
    sub = sub[1].strip() if len(sub) > 1 else str(crumb or "").strip()
    if not sub:
        return
    box = _tb(slide, 0.0, -0.28, 1.39, 0.01, wrap=False)
    box.left = Emu(0)
    box.top = Emu(int(-0.28 * 914400))
    r = box.text_frame.paragraphs[0].add_run(); r.text = sub
    setfont(r, 15.0, color=NAVY)
    _name(box, "upslide:subsection")


def content_top(headline, W, hsize=SZ_HEAD):
    """page_head() 將會回嘅內容起始 y —— 分頁前想預算可用高度就用呢個，
    唔好自己砌 HEAD_Y + head_h + 常數（會同 page_head 行開，高估可用高度而爆版）。"""
    if not headline:
        return CONTENT_Y
    return max(CONTENT_Y, HEAD_Y + head_h(headline, W, hsize)[0] + 0.06)


def page_head(slide, W, crumb, headline=None, *, hsize=SZ_HEAD, label_only=False):
    """灰色「章節 | 子題」+ navy 粗體導語 → 回內容起始 y。
    ★ 原報告【每版】內容都由 CONTENT_Y(1.55) 開始，唔跟導語浮動 —— 所以固定回 1.55，
      只有導語真係長過上限先順延（避免疊字）。"""
    if label_only:
        # ★ 原報告主要發現版（s30-44）：y=0.30 淨寫「主要發現」（冇「|」），
        #   導語落到 y=0.50、高 1.05。同其他章唔同，所以獨立一條路。
        put(slide, MARGIN, 0.30, W - 2 * MARGIN, 0.19, "主要發現",   # 原報告淨寫呢四個字
            size=SZ_TITLE, color=NAVY)
        subsec_marker(slide, crumb)
        if not headline:
            return CONTENT_Y
        box = _tb(slide, MARGIN, SUBTITLE_Y, W - 2 * MARGIN, 1.05)
        r = box.text_frame.paragraphs[0].add_run(); r.text = str(headline)
        setfont(r, hsize, bold=True, color=NAVY, heading=True)
        return CONTENT_Y
    put(slide, MARGIN, SUBTITLE_Y, W - 2 * MARGIN, 0.2, crumb, size=SZ_TITLE, bold=True, color=NAVY)
    subsec_marker(slide, crumb)
    if not headline:
        return CONTENT_Y
    h, hsize = head_h(headline, W, hsize)
    box = _tb(slide, MARGIN, HEAD_Y, W - 2 * MARGIN, h)
    p = box.text_frame.paragraphs[0]
    r = p.add_run(); r.text = str(headline)
    setfont(r, hsize, bold=True, color=NAVY, heading=True)
    return max(CONTENT_Y, HEAD_Y + h + 0.06)


def caption_bar(slide, x, y, w, text, *, size=SZ_CAPTION):
    """表頂 caption bar。★ 2026-09-08：原報告實測【冇】呢條 —— 表直接由 y=1.55 開始，
    表名喺 Tableau 截圖入面。SHOW_TABLE_CAPTION=False 就 no-op（回原 y，唔佔高度）。"""
    if not SHOW_TABLE_CAPTION:
        return y
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(0.17))
    bar.fill.solid(); bar.fill.fore_color.rgb = CAPTION_FILL
    bar.line.fill.background(); bar.shadow.inherit = False
    tf = bar.text_frame
    tf.margin_top = tf.margin_bottom = Emu(0)
    tf.margin_left = Emu(36000); tf.margin_right = Emu(18000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = str(text)
    setfont(r, size, bold=True, color=WHITE)
    return y + 0.17


SOURCE_LINE = "資料來源：管理層提供的項目投入明細表，管理層訪談；畢馬威分析"

# 調整表（1.4／2.2／2.4）表底嘅說明框 —— 原報告 s15／s22／s26 都有，
# 白底、8pt 粗體 navy、闊 7.3~7.6in。之前我哋完全冇呢一段。
ADJ_NOTE_OVERLAP = ("上表的調整金額已考慮不同調整項之間的重合部分。"
                    "若存在重合的金額，已在其中一項調整金額中列示，不會重複調整。")
ADJ_NOTE_1_4 = ADJ_NOTE_OVERLAP + "關於上表列示的各項潛在調整事項詳情，請見下頁。"
ADJ_NOTE_POST = (ADJ_NOTE_OVERLAP
                 + "對於同一類調整，上表的調整序號與「2025年度投資計劃報告投資金額的"
                   "潛在調整事項匯總」的調整序號一致。")


def adj_note(slide, y, text, *, w=7.4):
    """調整表下面嘅說明框（原報告 s15/s22/s26 實測：白底 8pt 粗體 navy）。回底部 y。"""
    if not text:
        return y
    h = min(est_lines(text, w, 8.0) * 8.0 * 1.35 / 72.0 + 0.06,
            max(0.18, CONTENT_BOTTOM + 0.30 - y))
    y = min(y, CONTENT_BOTTOM + 0.30 - h)
    box = _tb(slide, MARGIN, y, w, h)
    box.fill.solid(); box.fill.fore_color.rgb = WHITE
    p = box.text_frame.paragraphs[0]
    r = p.add_run(); r.text = text
    setfont(r, 8.0, bold=True, color=NAVY)
    return y + h


def table_footnote(slide, x, y, w, note=None, *, source=None):
    """表底下嘅「資料來源＋註釋」—— 原報告係【一個 7pt 文字框】，闊度＝表闊，
    緊貼表底（s10 y=5.94／s20 y=5.47／s24 y=5.37），唔係 pin 死版底、唔係兩個框。
    回文字框底部 y。"""
    txt = source or SOURCE_LINE
    if note:
        txt += "\n" + str(note).strip()
    lines = sum(max(1, est_lines(seg, w, SZ_NOTE)) for seg in txt.split("\n"))
    h = min(lines * SZ_NOTE * 1.3 / 72.0 + 0.04, max(0.16, CONTENT_BOTTOM + 0.16 - y))
    y = min(y, CONTENT_BOTTOM + 0.16 - h)
    put(slide, x, y, w, h, txt, size=SZ_NOTE, color=NOTE_FG)
    return y + h


def source_note(slide, W, y=None, *, note=None, more=False):
    """表下：資料來源（左）+（下頁待續）（右）。"""
    y = CONTENT_BOTTOM if y is None else y
    put(slide, MARGIN, y, W - 2.0, 0.16,
        note or "資料來源：管理層提供之項目投資計劃及執行報告資料，畢馬威分析",
        size=SZ_NOTE, color=NOTE_FG)
    if more:
        put(slide, W - MARGIN - 1.2, y, 1.2, 0.16, "（下頁待續）", size=SZ_NOTE, color=NOTE_FG,
            align=PP_ALIGN.RIGHT)


# ── 表格 ─────────────────────────────────────────────────────────────────
# ★ 真報告（IMG_0441 彩色版）嘅表：【冇逐格格線、冇 row 底色】。
#   得返：表頭 navy（設施建設/活動舉辦 嗰組 teal）＋ 小計/總計行上下幼橫線
#   ＋ 欄組之間虛線直線。之前全格線 + sec/小計/總計 3 級藍底 = 自己作，同報告唔同。
# ★ 表格配色（項目組 2026-08-17 逐項指定 hex，唔再靠影相估）：
RULE = "00338D"                             # 表格線（橫線 + 欄組虛線）＝ KPMG Blue
HDR_FILL = RGBColor(0x1E, 0x49, 0xE2)       # 表頭預設（accent1 亮藍）
#   ⚠ 唔可以叫 HDR —— make_report 已經有個 HDR(=NAVY)，bundle dedup 會靜靜丟咗佢
HDR_KEY = RGBColor(0x09, 0x8E, 0x7E)        # 綠：重點欄（獲批的計劃投資金額／潛在調整後／三年累計／比例）
HDR_SKY = RGBColor(0x00, 0xB8, 0xF5)        # 天藍：調整事項欄組（1-7+合計）／2025年度／潛在調整金額
HDR_PUR = RGBColor(0x48, 0x36, 0x98)        # 紫：2024年度（KPMG Violet；項目組未畀 hex，暫用品牌紫）
CAPTION_FILL = NAVY                         # caption 條 #00338D
SEC_FG = NAVY                               # 「博彩項目 / 非博彩項目」字色
NOTE_FG = NAVY                              # 註 / 資料來源
HDR1, HDR2, HDR3 = HDR_FILL, HDR_FILL, HDR_KEY        # 舊名保留（唔好散落 import error）
TEAL = HDR_KEY


def _edge(cell, side, *, w=9525, color=RULE, dash=None):
    """畫單一條邊（side ∈ T/B/L/R）。ln* 要插喺 tcPr 最前，否則 PowerPoint 會叫修復。
    同一邊重覆設就換走舊嗰條。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tag = qn(f"a:ln{side}")
    old = tcPr.find(tag)
    if old is not None:
        tcPr.remove(old)
    ln = tcPr.makeelement(tag, {"w": str(w), "cap": "flat"})
    fill = ln.makeelement(qn("a:solidFill"), {})
    clr = fill.makeelement(qn("a:srgbClr"), {"val": color})
    fill.append(clr); ln.append(fill)
    if dash:
        ln.append(ln.makeelement(qn("a:prstDash"), {"val": dash}))
    tcPr.insert(0, ln)


def set_cell(cell, text, *, size=SZ_TBL, bold=False, fill=None, align=PP_ALIGN.RIGHT,
             color=None, wrap=True, anchor=MSO_ANCHOR.MIDDLE, italic=False):
    cell.margin_left = cell.margin_right = Emu(18000)
    cell.margin_top = cell.margin_bottom = Emu(9000)
    cell.vertical_anchor = anchor
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill if fill is not None else WHITE
    tf = cell.text_frame; tf.word_wrap = wrap
    tf.clear()          # merge 會把被合併格嘅字搬入 origin → 唔清就會出「萬澳門元萬澳門元」
    p = tf.paragraphs[0]; p.alignment = align
    txt = "" if text is None else str(text)
    if color is None:
        color = NEG_COLOR if (NEG_COLOR is not None and txt.startswith("(")) else INK
    # ★ 空格一定要定死字號：PowerPoint 見到「冇 run／空 run」會跌返去 endParaRPr／預設 text
    #   style（python-pptx fresh deck ＝ Calibri 18pt）→ row 被撐到 ~0.3in，成張表爆版。
    #   所以：paragraph 層 defRPr + endParaRPr 都寫死，而且空字串索性唔加 run。
    p.font.size = Pt(size); p.font.bold = bold
    p.font.name = FONT_NUM; set_ea(p.font)
    epr = p._p.get_or_add_endParaRPr()
    epr.set("sz", str(int(round(size * 100))))
    if not txt:
        return
    # ⚠ DrawingML 入面 "\n" 唔係換行（會當空白）→ 一定要用 <a:br/>，否則表頭喺 PowerPoint 會擠成一行
    for i, seg in enumerate(txt.split("\n")):
        if i:
            p.add_line_break()
        if not seg:
            continue
        r = p.add_run(); r.text = seg
        setfont(r, size, bold=bold, italic=italic, color=color)


_NSA = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def cell_border(cell, spec):
    """寫格仔框線。spec = {"T": [hex, pt, dashed], …}（extract_canned 抽返嚟嗰個形狀）。
    ⚠ 只用喺【罐頭表】—— 我哋自己生成嗰啲表係跟原報告嘅 Tableau 截圖樣式（冇全框、
      只有指定橫線），照畫全框會走樣。"""
    if not spec:
        return
    from lxml import etree
    tc = cell._tc
    tcPr = tc.find(f"{_NSA}tcPr")
    if tcPr is None:
        tcPr = etree.SubElement(tc, f"{_NSA}tcPr")
    for side in "LRTB":                        # OOXML 次序：lnL lnR lnT lnB，錯序 PowerPoint 會當壞檔
        v = spec.get(side)
        if not v:
            continue
        hexv, pt, dashed = (list(v) + [1.0, False])[:3]
        old = tcPr.find(f"{_NSA}ln{side}")
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, f"{_NSA}ln{side}")
        ln.set("w", str(int(round(float(pt or 1.0) * 12700))))
        ln.set("cap", "flat"); ln.set("cmpd", "sng"); ln.set("algn", "ctr")
        fill = etree.SubElement(ln, f"{_NSA}solidFill")
        etree.SubElement(fill, f"{_NSA}srgbClr").set("val", str(hexv or "000000"))
        if dashed:
            etree.SubElement(ln, f"{_NSA}prstDash").set("val", "dash")
    # lnL/lnR/lnT/lnB 必須排喺 tcPr 最前（schema 次序），其餘子元素跟後
    order = {f"{_NSA}ln{s}": i for i, s in enumerate("LRTB")}
    for ch in sorted(list(tcPr), key=lambda e: order.get(e.tag, 99)):
        tcPr.append(ch)


ROW_FILL = {"sec": None, "subtot": None, "tot": None, "data": None, "formula": None}   # 報告：body 全白，靠橫線分層


def header_h(supers, subs, widths, hfont):
    """表頭需要高度（吋）。"""
    h = row_h(subs, widths, hfont, pad_in=0.05, min_h=0.20)
    return (0.17 + h) if supers else h


def fit_rows(rows, widths, font, avail_h, hh):
    """按【累積高度】切頁 → 保證唔會超出可用高度。rows = [(kind, cells)]。
    keep=True 嘅 row（範疇 block）盡量唔拆：見 fit_blocks。"""
    out, cur, used = [], [], 0.0
    cap = max(avail_h - hh, 0.6)          # guard：導語太長時唔好變 0/負數（會無限開版）
    for kind, cells in rows:
        h = row_h(cells, widths, font)
        if cur and used + h > cap:
            out.append(cur); cur, used = [], 0.0
        cur.append((kind, cells)); used += h
    if cur:
        out.append(cur)
    return out or [[]]


def fit_blocks(blocks, widths, font, avail_h, hh):
    """block = 一個範疇（section + data + 小計）。整個 block 唔拆頁（對 scan：全報告冇「續」）；
    單一 block 大過一版先逼住切。"""
    cap = max(avail_h - hh, 0.6)          # guard：同上
    pages, cur, used = [], [], 0.0
    for blk in blocks:
        bh = sum(row_h(c, widths, font) for _, c in blk)
        if cur and used + bh > cap:
            pages.append(cur); cur, used = [], 0.0
        if bh > cap:                                    # 單一範疇爆版 → 逐行切（安全網）
            for kind, cells in blk:
                h = row_h(cells, widths, font)
                if cur and used + h > cap:
                    pages.append(cur); cur, used = [], 0.0
                cur.append((kind, cells)); used += h
            continue
        cur.extend(blk); used += bh
    if cur:
        pages.append(cur)
    return pages or [[]]


def draw_table(slide, x, y, w, subs, rows, widths, *, supers=None, font=SZ_TBL, hfont=SZ_TBL_HDR,
               left_cols=1, fill_h=None, max_row_h=0.26, hdr_cols=None):
    """畫 navy 表。subs=欄名（可含 \\n）；rows=[(kind, cells)]；widths=相對闊度（會 scale 到 w）。
    supers=[(label, c0, c1_exclusive)] 兩層表頭。fill_h=想填滿嘅高度（行數少時撐開行高，
    唔好剩一大橛白位；每行最多 max_row_h）。回 (bottom_y, 實際高度)。"""
    ncol = len(subs)
    scale = w / sum(widths)
    wid = [v * scale for v in widths]
    nhdr = 2 if supers else 1
    heights = [row_h(cells, wid, font) for _, cells in rows]
    hsub = row_h(subs, wid, hfont, pad_in=0.05, min_h=0.20)
    if fill_h and heights:
        slack = fill_h - ((0.17 if supers else 0) + hsub + sum(heights))
        if slack > 0.05:
            add = min(slack / len(heights), max(0.0, max_row_h - max(heights)))
            heights = [h + add for h in heights]
    total_h = (0.17 if supers else 0) + hsub + sum(heights)
    tbl = slide.shapes.add_table(nhdr + len(rows), ncol, Inches(x), Inches(y),
                                 Inches(w), Inches(total_h)).table
    tbl.first_row = False; tbl.horz_banding = False
    for i, v in enumerate(wid):
        tbl.columns[i].width = Inches(v)
    # 三色欄組——【要 caller 明示】：4.2 表都有「設施建設/活動舉辦」欄但成排 navy（scan p.24），
    #   所以唔可以淨靠欄名估。hdr_cols = {欄 index: 顏色}，冇指定就 HDR1。
    hc = dict(hdr_cols or {})
    if supers:
        for c in range(ncol):
            set_cell(tbl.cell(0, c), "", size=hfont, fill=hc.get(c, HDR1), color=WHITE)
        # 一個欄組跨住兩隻表頭色（報告：潛在調整後 = 深藍嗰兩欄 + 綠嗰兩欄）→ 拆開兩格，
        #   個 label 兩邊都寫（同 IMG_0441 一樣，「潛在調整後投資金額」出現兩次）。
        for label, c0, c1 in supers:
            a = c0
            while a < c1:
                b = a + 1
                while b < c1 and hc.get(b, HDR1) == hc.get(a, HDR1):
                    b += 1
                if b - a > 1:
                    tbl.cell(0, a).merge(tbl.cell(0, b - 1))
                set_cell(tbl.cell(0, a), label or "", size=hfont + 0.5, bold=True,
                         fill=hc.get(a, HDR1), color=WHITE, align=PP_ALIGN.CENTER)
                a = b
        tbl.rows[0].height = Emu(int(0.17 * 914400))
    for c, s in enumerate(subs):
        # 報告嘅欄名喺表頭【貼底】（wrap 做兩行時尤其明顯）
        set_cell(tbl.cell(nhdr - 1, c), s, size=hfont, bold=True, anchor=MSO_ANCHOR.BOTTOM,
                 fill=hc.get(c, HDR1), color=WHITE,
                 align=PP_ALIGN.LEFT if c < left_cols else PP_ALIGN.CENTER)
    # 角位：報告係一格（序號欄冇字），單位「萬澳門元」貼住最左
    if left_cols >= 2 and not str(subs[0]).strip() and str(subs[1]).strip():
        tbl.cell(nhdr - 1, 0).merge(tbl.cell(nhdr - 1, 1))
        set_cell(tbl.cell(nhdr - 1, 0), subs[1], size=hfont, bold=True, anchor=MSO_ANCHOR.BOTTOM,
                 fill=hc.get(0, HDR1), color=WHITE, align=PP_ALIGN.LEFT)
    tbl.rows[nhdr - 1].height = Emu(int(hsub * 914400))
    for ri, (kind, cells) in enumerate(rows, start=nhdr):
        bold = kind in ("sec", "subtot", "tot")
        if kind == "formula":      # 報告表頭下面嗰行斜體公式（a｜1..7｜b｜c=a+b｜d=b/a）
            for c, v in enumerate(cells):
                set_cell(tbl.cell(ri, c), v, size=max(4.5, font - 1.0), italic=True,
                         color=GREY, align=PP_ALIGN.LEFT if c < left_cols else PP_ALIGN.RIGHT)
            tbl.rows[ri].height = Emu(int(max(0.14, (font - 1.0) * 1.24 / 72.0 + 0.03) * 914400))
            continue
        # 標籤（範疇/小計/總計/表尾說明行）喺報告係【由最左邊起】，唔係縮喺名稱欄：
        #   序號欄空 + 名稱欄有字 → merge 埋，個 label 先有位唔會 wrap
        k = 0
        if left_cols >= 2 and not str(cells[0]).strip() and str(cells[1]).strip():
            k = 1
            tbl.cell(ri, 0).merge(tbl.cell(ri, 1))
        for c, v in enumerate(cells):
            if 0 < c <= k:
                continue                      # 已 merge 入 col 0
            al = PP_ALIGN.LEFT if c < left_cols else PP_ALIGN.RIGHT
            set_cell(tbl.cell(ri, c), cells[k] if c == 0 and k else v,
                     size=font, bold=bold, fill=ROW_FILL.get(kind), align=al,
                     color=SEC_FG if kind == "sec" else None)
        tbl.rows[ri].height = Emu(int(heights[ri - nhdr] * 914400))
    # ── 線：只有小計/總計橫線 + 欄組虛線直線（報告冇逐格格線）────────────
    # 只喺【有名嘅欄組】邊界畫虛線；標籤欄自成一「組」（label=""）唔算
    gsep = {c0 for _l, c0, _c1 in (supers or []) if c0 > 0 and str(_l).strip()}
    last = nhdr + len(rows) - 1
    for ri, (kind, _c) in enumerate(rows, start=nhdr):
        if kind in ("subtot", "tot"):
            for c in range(ncol):
                _edge(tbl.cell(ri, c), "T")
                _edge(tbl.cell(ri, c), "B")      # 報告：小計同總計【上下都有】幼線
    for ri in range(nhdr, nhdr + len(rows)):
        for c in gsep:
            _edge(tbl.cell(ri, c), "L", w=6350, color=RULE, dash="sysDash")
    return y + total_h, total_h


# ── 敘述 ─────────────────────────────────────────────────────────────────
def prose(box, items, *, head_size=SZ_BODY_HEAD, body_size=SZ_BODY, gap=6):
    """scan 敘述格式：navy 粗體小標題一行 + 下面 body 段落（唔用 ■ bullet）。
    items = [(head, body)]；head 可為空。"""
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for head, body in items:
        if head:
            head = str(head).rstrip("：:")        # 項目組：小標題唔應該有冒號
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.space_before = Pt(0 if first else gap); p.space_after = Pt(1)
            p.font.size = Pt(head_size); p.font.name = FONT_NUM; set_ea(p.font)
            p._p.get_or_add_endParaRPr().set("sz", str(int(round(head_size * 100))))
            r = p.add_run(); r.text = str(head)
            setfont(r, head_size, bold=True, color=NAVY)
            first = False
        if body:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            p.space_before = Pt(0 if first else (0 if head else gap)); p.space_after = Pt(1)
            p.font.size = Pt(body_size); p.font.name = FONT_NUM; set_ea(p.font)
            p._p.get_or_add_endParaRPr().set("sz", str(int(round(body_size * 100))))
            r = p.add_run(); r.text = str(body)
            setfont(r, body_size, color=RGBColor(0x33, 0x33, 0x33))
            first = False


def prose_numbered(box, items, *, size=SZ_BODY, gap=7, indent=0.24, title=None, tsize=SZ_BODY_HEAD):
    """scan 表旁格式（p-11/p-13 右欄）：navy 粗體小標題 + 編號清單
        1.  {粗體類型}（{金額}）：{內文…}       ← hanging indent，內文對齊類型名
    items = [(編號, 粗體引子, 內文)]；編號跟七大類 canonical 序（會跳號）。"""
    tf = box.text_frame; tf.word_wrap = True
    first = True
    if title:
        p = tf.paragraphs[0]; p.space_after = Pt(4)
        p.font.size = Pt(tsize)
        p._p.get_or_add_endParaRPr().set("sz", str(int(round(tsize * 100))))
        r = p.add_run(); r.text = str(title)
        setfont(r, tsize, bold=True, color=NAVY)
        first = False
    emu = int(indent * 914400)
    for no, head, body in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(gap); p.space_after = Pt(0)
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(emu)); pPr.set("indent", str(-emu))     # hanging indent
        p.font.size = Pt(size); p.font.name = FONT_NUM; set_ea(p.font)
        p._p.get_or_add_endParaRPr().set("sz", str(int(round(size * 100))))
        rn = p.add_run(); rn.text = f"{no}.\t"
        setfont(rn, size, bold=True, color=NAVY)
        rh = p.add_run(); rh.text = str(head)
        setfont(rh, size, bold=True, color=NAVY)
        rb = p.add_run(); rb.text = str(body)
        setfont(rb, size, color=RGBColor(0x33, 0x33, 0x33))


def est_numbered_h(items, w, size=SZ_BODY, gap=7, title=None, tsize=SZ_BODY_HEAD, indent=0.24):
    h = (est_lines(title, w, tsize) * tsize * 1.3 / 72.0 + 4 / 72.0) if title else 0.0
    for _no, head, body in items:
        h += est_lines(f"　{head}{body}", w - indent, size) * size * 1.35 / 72.0 + gap / 72.0
    return h


def prose_box(slide, x, y, w, h, items, **kw):
    box = _tb(slide, x, y, w, h)
    prose(box, items, **kw)
    return box


def est_prose_h(items, w, head_size=SZ_BODY_HEAD, body_size=SZ_BODY, gap=6):
    """估敘述高度（吋）→ 用嚟分頁，唔會爆版。"""
    h = 0.0
    for head, body in items:
        if head:
            h += est_lines(head, w, head_size) * head_size * 1.3 / 72.0 + gap / 72.0
        if body:
            h += est_lines(body, w, body_size) * body_size * 1.35 / 72.0 + 2 / 72.0
    return h


def fit_prose(items, w, avail_h, **kw):
    """按估算高度切頁 → [[items]]。"""
    pages, cur, used = [], [], 0.0
    avail_h = max(avail_h, 0.6)           # guard
    for it in items:
        ih = est_prose_h([it], w, **kw)
        if cur and used + ih > avail_h:
            pages.append(cur); cur, used = [], 0.0
        cur.append(it); used += ih
    if cur:
        pages.append(cur)
    return pages or [[]]


# ── 深色版（封面 / 章節分隔）─────────────────────────────────────────────
def apply_theme_fonts(prs):
    """把生成 deck 嘅 theme 字體改成公司 template 嗰套（major KPMG Bold / minor Arial，
    ea 兩者都 Microsoft YaHei）。python-pptx 開新檔用 Office 預設 theme（Calibri），
    凡係我哋冇明寫字體嘅地方（placeholder、空段落、表格預設）都會跌返 Calibri。"""
    import re as _re
    try:
        for m in prs.slide_masters:
            part = m.part.part_related_by(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
            xml = part.blob.decode("utf-8")
            for tag, latin in (("majorFont", FONT_HEAD), ("minorFont", FONT_NUM)):
                def _fix(mo, latin=latin):
                    seg = mo.group(0)
                    seg = _re.sub(r'<a:latin typeface="[^"]*"', f'<a:latin typeface="{latin}"', seg, count=1)
                    seg = _re.sub(r'<a:ea typeface="[^"]*"', f'<a:ea typeface="{FONT_CN}"', seg, count=1)
                    return seg
                xml = _re.sub(r"<a:" + tag + r">.*?</a:" + tag + r">", _fix, xml, flags=_re.S)
            part._blob = xml.encode("utf-8")
    except Exception:
        pass            # theme 改唔到唔應該搞冧成個 build（每個 run 都已經明寫咗字體）


def dark_slide(prs):
    slide = blank(prs)
    W, H = size_of(prs)
    rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid(); rect.fill.fore_color.rgb = DARK
    rect.line.fill.background(); rect.shadow.inherit = False
    kb = _tb(slide, 0.55, 0.35, 2.2, 0.4)
    kr = kb.text_frame.paragraphs[0].add_run(); kr.text = "KPMG"
    setfont(kr, 20, bold=True, italic=True, color=WHITE)
    return slide, W, H
