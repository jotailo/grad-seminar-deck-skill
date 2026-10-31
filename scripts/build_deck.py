#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""build_deck.py — 研究生组会 PPT 结构化生成器（16:9，白底，单一强调色，无渐变无剪贴画）。

用法：
    python build_deck.py deck.json -o 组会汇报.pptx
    python build_deck.py deck.json -o 组会汇报.pptx --accent 1F6F5C
    python build_deck.py deck.json -o 组会汇报.pptx --skip-budget-check

JSON 规格见 references/deck-schema.md；完整样例见 assets/deck.example.json。

结构：deck.json 用「part（汇报种类）→ unit（一篇论文 / 一组实验 / 一个项目）→ page」三级描述，
生成时自动插入种类分区页、自动生成目录、自动编页码与页脚。每个 unit 的内容页必须落在 2–4 页，
不符合直接报错退出（--skip-budget-check 可临时放过，但交付前必须补正）。

依赖：python-pptx, Pillow, lxml
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

try:
    from lxml import etree
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Emu, Inches, Pt
except ImportError as e:  # pragma: no cover
    sys.exit(f"缺少依赖：{e}；请 pip install python-pptx pillow lxml")

try:
    from PIL import Image, ImageFont
except ImportError:  # pragma: no cover
    Image = None
    ImageFont = None

# ---------------------------------------------------------------- 设计常量
SW, SH = 13.3333, 7.5          # 16:9，英寸
MX = 0.85                      # 左右安全边距
CW = SW - 2 * MX               # 内容宽度 ≈ 11.63in
GUT = 0.34                     # 列间距

INK = "1A1A1A"
BODY = "333333"
GRAY = "5F6672"
LIGHT = "969CA6"
RULE = "E3E6EB"
PANEL = "F6F7F9"
ACCENT = "2F5597"              # 单一强调色
ACCENT_SOFT = "EDF1F8"
WHITE = "FFFFFF"

# 字号档位（投影可读是底线：正文 ≥14pt，任何文字 ≥10pt）
SZ_TITLE = 27
SZ_TITLE_MIN = 20
SZ_KICKER = 11
SZ_BODY = 18
SZ_BODY_MIN = 14
SZ_SUB = 15
SZ_NOTE = 13.5
SZ_TAKE = 14
SZ_CAP = 12
SZ_FOOT = 10.5
SZ_TABLE_MAX = 15
SZ_TABLE_MIN = 12

KIND_LABEL = {
    "paper": "论文阅读汇报",
    "experiment": "实验报告",
    "project": "项目进展汇报",
}
KIND_ROMAN = {"paper": "01", "experiment": "02", "project": "03"}  # 分区页大号数字，与目录编号一致

BUDGET_MIN, BUDGET_MAX = 2, 4   # 每个 unit 的内容页数区间


FONT_PATHS = [
    (r"C://Windows//Fonts//msyh.ttc", "微软雅黑"),
    (r"C://Windows//Fonts//msyhbd.ttc", "微软雅黑"),
    (r"C://Windows//Fonts//simhei.ttf", "黑体"),
    ("/System/Library/Fonts/PingFang.ttc", "PingFang SC"),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "Noto Sans CJK SC"),
    ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", "WenQuanYi Micro Hei"),
]
FONT_PATH = next((Path(p) for p, _ in FONT_PATHS if Path(p).exists()), None)


def pick_cn_font() -> str:
    """选一个本机真实存在的中文字体，避免回退成宋体发虚。"""
    import platform
    cands = [
        (r"C:\Windows\Fonts\msyh.ttc", "微软雅黑"),
        (r"C:\Windows\Fonts\msyhbd.ttc", "微软雅黑"),
        (r"C:\Windows\Fonts\simhei.ttf", "黑体"),
        ("/System/Library/Fonts/PingFang.ttc", "PingFang SC"),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "Noto Sans CJK SC"),
        ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", "WenQuanYi Micro Hei"),
    ]
    for path, name in cands:
        if Path(path).exists():
            return name
    return "微软雅黑" if platform.system() == "Windows" else "PingFang SC"


FONT_CN = pick_cn_font()
FONT_EN = "Segoe UI" if FONT_CN == "微软雅黑" else "Helvetica Neue"


# ---------------------------------------------------------------- 基础工具
def rgb(h: str) -> RGBColor:
    return RGBColor.from_string((h or "").lstrip("#").upper())


def set_run_font(run, size=None, bold=None, color=None, font=FONT_CN, italic=None):
    """同时写 latin / 东亚 / 复杂文种，避免中文回退成宋体。"""
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = rgb(color)
    run.font.name = font
    rPr = run.font._rPr
    # 标记语言，让 PowerPoint 按中文避头尾规则断行（否则句号会被挤到下一行）
    rPr.set("lang", "zh-CN")
    rPr.set("altLang", "en-US")
    for tag in ("a:ea", "a:cs"):
        for el in rPr.findall(qn(tag)):
            rPr.remove(el)
    for tag in ("a:ea", "a:cs"):
        etree.SubElement(rPr, qn(tag)).set("typeface", font)


def add_rich_text(p, text, size, color, bold=False, font=FONT_CN):
    """支持 **加粗**、__强调色__、`等宽` 三种轻量标记。"""
    import re
    tokens = re.split(r"(\*\*.+?\*\*|__.+?__)", str(text))
    for tk in tokens:
        if not tk:
            continue
        b, c = bold, color
        if tk.startswith("**") and tk.endswith("**") and len(tk) > 4:
            tk, b = tk[2:-2], True
        elif tk.startswith("__") and tk.endswith("__") and len(tk) > 4:
            tk, b, c = tk[2:-2], True, ACCENT
        r = p.add_run()
        r.text = tk
        set_run_font(r, size=size, bold=b, color=c, font=font)


def allow_hang_punct(p):
    """允许行末标点悬挂在版心外：否则句号会被挤到下一行，出现只有一两个标点的孤行。"""
    p._p.get_or_add_pPr().set("hangingPunc", "1")


def set_bullet(p, char="•", color=ACCENT, mar_l=0.30, ind=0.20):
    allow_hang_punct(p)
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(mar_l * 914400)))
    pPr.set("indent", str(-int(ind * 914400)))
    for tag in ("a:buClr", "a:buFont", "a:buChar", "a:buNone", "a:buSzPct"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    etree.SubElement(buClr, qn("a:srgbClr")).set("val", color)
    etree.SubElement(pPr, qn("a:buFont")).set("typeface", "Arial")
    etree.SubElement(pPr, qn("a:buChar")).set("char", char)
    tabLst = etree.SubElement(pPr, qn("a:tabLst"))
    etree.SubElement(tabLst, qn("a:tab")).set("pos", str(int(mar_l * 914400)))


def clear_bullet(p):
    allow_hang_punct(p)
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    for tag in ("a:buChar", "a:buClr", "a:buFont", "a:buSzPct"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    etree.SubElement(pPr, qn("a:buNone"))


def textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    return tb


def para(tf, first=False):
    return tf.paragraphs[0] if first else tf.add_paragraph()


def line_rule(slide, l, t, w, color=RULE, thick=1.0):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Pt(thick))
    ln.fill.solid()
    ln.fill.fore_color.rgb = rgb(color)
    ln.line.fill.background()
    ln.shadow.inherit = False
    return ln


def accent_rule(slide, l, t, w=0.55, color=ACCENT, thick=3.0):
    return line_rule(slide, l, t, w, color=color, thick=thick)


def panel(slide, l, t, w, h, fill=PANEL, bar=None):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    box.fill.solid()
    box.fill.fore_color.rgb = rgb(fill)
    box.line.fill.background()
    box.shadow.inherit = False
    if bar:
        b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(0.055), Inches(h))
        b.fill.solid()
        b.fill.fore_color.rgb = rgb(bar)
        b.line.fill.background()
        b.shadow.inherit = False
    return box


def add_footer(slide, left_text, page_no, total):
    tb = textbox(slide, MX, SH - 0.52, CW - 1.0, 0.28)
    add_rich_text(tb.text_frame.paragraphs[0], left_text, SZ_FOOT, LIGHT)
    tb2 = textbox(slide, SW - MX - 0.9, SH - 0.52, 0.9, 0.28)
    p2 = tb2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.RIGHT
    add_rich_text(p2, f"{page_no} / {total}", SZ_FOOT, LIGHT)


def title_size(text: str) -> float:
    n = len(text)
    if n <= 26:
        return SZ_TITLE
    if n <= 38:
        return SZ_TITLE - 3
    if n <= 54:
        return SZ_TITLE - 5
    return SZ_TITLE_MIN - 1


def slide_title(slide, title, kicker=None, sub=None):
    """页头：小标签 + 标题 + 强调短线。返回正文起点 y。"""
    y = 0.78
    if kicker:
        tb = textbox(slide, MX, 0.46, CW, 0.26)
        add_rich_text(tb.text_frame.paragraphs[0], kicker, SZ_KICKER, ACCENT, bold=True)
        y = 0.80
    tb = textbox(slide, MX, y, CW, 0.86)
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.16
    add_rich_text(p, title, title_size(title), INK, bold=True)
    accent_rule(slide, MX, y + 0.92, 0.5, ACCENT, 2.5)
    body_top = y + 1.18
    if sub:
        tb = textbox(slide, MX, body_top - 0.06, CW, 0.32)
        add_rich_text(tb.text_frame.paragraphs[0], sub, SZ_NOTE, GRAY)
        body_top += 0.40
    return body_top


def add_takeaway(slide, text, top=None):
    """页脚上方的一条结论条：这一页到底说明什么。"""
    top = SH - 1.42 if top is None else top
    h = 0.62
    box = panel(slide, MX, top, CW, h, fill=ACCENT_SOFT, bar=ACCENT)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.22)
    tf.margin_right = Inches(0.18)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    if len(str(text)) > 78:
        print(f"[warn] 结论条过长（{len(str(text))} 字），请压到 45 字以内：{str(text)[:18]}…")
    add_rich_text(p, text, SZ_TAKE, BODY)


def fit_picture(slide, path, l, t, w, h, frame=True, valign="top"):
    if Image is None:
        iw, ih = 1600, 900
    else:
        with Image.open(path) as im:
            iw, ih = im.size
    img_ar, box_ar = iw / ih, w / h
    if img_ar >= box_ar:
        pw, ph = w, w / img_ar
    else:
        ph, pw = h, h * img_ar
    px = l + (w - pw) / 2
    py = t + (h - ph) / 2 if valign == "middle" else t
    if frame:
        pad = 0.05
        fr = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(px - pad), Inches(py - pad),
                                    Inches(pw + 2 * pad), Inches(ph + 2 * pad))
        fr.fill.background()
        fr.line.color.rgb = rgb(RULE)
        fr.line.width = Pt(0.75)
        fr.shadow.inherit = False
    slide.shapes.add_picture(str(path), Inches(px), Inches(py), Inches(pw), Inches(ph))
    return px, py, pw, ph


def add_caption(slide, caption, l, t, w):
    tb = textbox(slide, l, t, w, 0.52)
    p = tb.text_frame.paragraphs[0]
    p.line_spacing = 1.15
    add_rich_text(p, caption, SZ_CAP, GRAY)


# ---------------------------------------------------------------- 表格（自动适配字号）
_FONT_CACHE: dict[int, object] = {}


def _char_w(ch: str, size: float) -> float:
    """单字宽度（pt）：优先用本机真实字体度量，与 check_deck.py 同一口径，避免「生成时算得下、
    体检时算溢出」；没有字体文件时退回比例估算（中文 1.02em、西文 0.60em）。"""
    if ImageFont is not None and FONT_PATH is not None:
        key = int(round(size))
        f = _FONT_CACHE.get(key)
        if f is None:
            f = ImageFont.truetype(str(FONT_PATH), max(6, key))
            _FONT_CACHE[key] = f
        return f.getlength(ch)
    return size * (1.02 if ord(ch) > 0x2E80 else 0.60)


def _wrap_lines(text: str, size: float, box_pt: float) -> int:
    lines, cur = 1, 0.0
    for ch in str(text):
        if ch == "\n":
            lines += 1
            cur = 0.0
            continue
        w = _char_w(ch, size)
        if cur + w > box_pt:
            lines += 1
            cur = w
        else:
            cur += w
    return lines


def _clean(text: str) -> str:
    import re
    return re.sub(r"\*\*|__", "", str(text))


# ---------------------------------------------------------------- 文字自适应
def est_lines(text: str, size: float, box_in: float) -> int:
    return _wrap_lines(_clean(text), size, max(box_in * 72, 20))


def _wrap_widths(text: str, size: float, box_pt: float):
    """返回每行的实际占用宽度（pt），用于判断末行是不是只剩一两个字。"""
    widths, cur = [], 0.0
    for ch in _clean(text):
        if ch == "\n":
            widths.append(cur)
            cur = 0.0
            continue
        w = _char_w(ch, size)
        if cur + w > box_pt:
            widths.append(cur)
            cur = w
        else:
            cur += w
    widths.append(cur)
    return widths


def widow_hits(bullets, size, box_w, ratio=0.10):
    """找出「末行只剩很短一截」的要点：这类孤字/孤标点投出来很难看，最好改写。"""
    hits = []
    for it in bullets:
        text = it if isinstance(it, str) else it.get("text", "")
        lvl = 0 if isinstance(it, str) else int(it.get("level", 0))
        s_ = size if lvl == 0 else max(size - 2, SZ_BODY_MIN)
        box_pt = max((box_w - (0.28 if lvl == 0 else 0.58)) * 72, 20)
        widths = _wrap_widths(text, s_, box_pt)
        if len(widths) <= 1:
            continue
        # 扣掉末尾可以悬挂的标点：这些标点会被挂在版心外，不算孤字
        tail = widths[-1]
        for ch in reversed(_clean(text).rstrip()):
            if ch in "。，、；：！？）」』”’),.;:!?)]}":
                w = _char_w(ch, s_)
                if tail - w > 0:
                    tail -= w
                else:
                    tail = 0
                    break
            else:
                break
        if 0 < tail < box_pt * ratio:
            hits.append(_clean(text)[:14])
    return hits


def est_bullets_h(bullets, size: float, box_w: float, gap0=8.0, gap1=5.0) -> float:
    """估算要点块的渲染高度（英寸），把缩进和段间距都算进去。"""
    h = 0.0
    for it in bullets:
        text = it if isinstance(it, str) else it.get("text", "")
        lvl = 0 if isinstance(it, str) else int(it.get("level", 0))
        s = size if lvl == 0 else max(size - 2, SZ_BODY_MIN)
        indent = 0.28 if lvl == 0 else 0.58
        n = est_lines(text, s, box_w - indent)
        h += n * s * 1.32 * 1.16 / 72 + (gap0 if lvl == 0 else gap1) / 72
    return h


def fit_bullet_size(bullets, box_w, box_h, max_size=SZ_BODY, min_size=SZ_BODY_MIN) -> float:
    """选一个能塞进 box_h 的最大字号；塞不下就用最小字号（交给 check_deck.py 报溢出）。"""
    size = max_size
    while size > min_size:
        if est_bullets_h(bullets, size, box_w) <= box_h * 0.93:
            return size
        size -= 1
    return min_size


def spread_gap(bullets, size, box_w, box_h, base=8.0, cap=16.0) -> float:
    """内容偏少时把段间距拉开，让版面填满而不是全挤在上半页。"""
    used = est_bullets_h(bullets, size, box_w, gap0=base)
    free_pt = (box_h - used) * 72
    n = max(1, len(bullets))
    return max(base, min(cap, base + free_pt * 0.55 / n))


def plan_table(cols, rows, widths, total_w, avail_h, lh=1.30, pad=0.13):
    """为表格挑一个能塞进 avail_h 的最大字号，并给出每行高度（英寸）。"""
    n_c = len(cols)
    for size in (SZ_TABLE_MAX, 14, 13, 12, SZ_TABLE_MIN):
        heights, ok = [], True
        for ri in range(len(rows) + 1):
            lines = 1
            for ci in range(n_c):
                cell_w_pt = (widths[ci] * total_w - 0.24) * 72
                val = cols[ci] if ri == 0 else rows[ri - 1][ci]
                lines = max(lines, _wrap_lines(_clean(val), size, max(cell_w_pt, 20)))
            h = max(0.40, lines * size * lh / 72 + pad)
            heights.append(h)
        if sum(heights) <= avail_h:
            return size, heights, True
        last = (size, heights)
    return last[0], last[1], False


def set_cell_borders(cell, bottom=None, top=None, color=RULE, width=0.75):
    """只画需要的横线：简洁表格靠留白和细线，不靠框。"""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    specs = []
    for tag, on, col, wd in (("a:lnL", False, None, 0), ("a:lnR", False, None, 0),
                             ("a:lnT", top, color if top else color, width if top else 0),
                             ("a:lnB", bottom, color if bottom else color, width if bottom else 0)):
        ln = etree.Element(qn(tag))
        ln.set("w", str(int((wd or 0.75) * 12700)))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        if on:
            etree.SubElement(ln, qn("a:solidFill")).append(
                etree.Element(qn("a:srgbClr"), val=col))
        else:
            etree.SubElement(ln, qn("a:noFill"))
        specs.append(ln)
    for i, ln in enumerate(specs):      # 顺序固定为 L,R,T,B，且必须排在填充之前
        tcPr.insert(i, ln)


def draw_table(slide, cols, rows, top, bottom, widths=None, caption=None, total_ratio=1.0,
               highlight_col=None):
    n_c = len(cols)
    widths = widths or [1 / n_c] * n_c
    total_w = CW * total_ratio
    cap_h = 0.42 if caption else 0.0
    size, heights, fitted = plan_table(cols, rows, widths, total_w, bottom - top - cap_h)
    if not fitted:
        print(f"[warn] 表格偏密：「{str(cols[0])[:12]}…」{len(rows)} 行按最小字号 {size}pt 仍可能塞不下"
              f"（需要 {sum(heights):.2f}in，只有 {bottom - top - cap_h:.2f}in），"
              f"请精简单元格文字或拆成两页")
    # 行数少时把行高撑开并略微垂直居中，避免「表在上半页、下半页空一大片」
    avail = bottom - top - cap_h
    total = sum(heights)
    if total < avail:
        grow = min(avail - total, total * 0.42 + 0.20 * len(heights))
        scale = (total + grow) / total if total else 1.0
        heights = [min(h * scale, 1.15) for h in heights]
        total = sum(heights)
        top += min(0.15 if cap_h else 0.40, max(0.0, (avail - total) * 0.45))

    frame = slide.shapes.add_table(len(rows) + 1, n_c, Inches(MX), Inches(top),
                                   Inches(total_w), Inches(sum(heights)))
    tbl = frame.table
    tbl.first_row = False
    tbl.horz_banding = False
    try:
        tbl._tbl.tblPr.set("firstRow", "0")
        tbl._tbl.tblPr.set("bandRow", "0")
    except Exception:
        pass
    for i, w in enumerate(widths):
        tbl.columns[i].width = Emu(int(Inches(total_w) * w))
    for i, h in enumerate(heights):
        tbl.rows[i].height = Inches(h)

    for ri in range(len(rows) + 1):
        for ci in range(n_c):
            cell = tbl.cell(ri, ci)
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            cell.margin_top = Inches(0.05)
            cell.margin_bottom = Inches(0.05)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            cell.fill.fore_color.rgb = rgb(ACCENT_SOFT if ri == 0 else
                                           (PANEL if ri % 2 == 0 else WHITE))
            if ri == 0:
                set_cell_borders(cell, bottom=ACCENT, color=ACCENT, width=1.0)
            else:
                set_cell_borders(cell, bottom=(RULE if ri < len(rows) else None))
            tf = cell.text_frame
            tf.word_wrap = True
            val = cols[ci] if ri == 0 else rows[ri - 1][ci]
            p = tf.paragraphs[0]
            p.line_spacing = 1.22
            allow_hang_punct(p)
            color = ACCENT if ri == 0 else (ACCENT if ci == highlight_col else BODY)
            add_rich_text(p, str(val), size, color, bold=(ri == 0))
    if caption:
        add_caption(slide, caption, MX, top + sum(heights) + 0.10, total_w)
    return sum(heights)


# ---------------------------------------------------------------- 版式
def slide_cover(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    tb = textbox(sl, 0.95, 1.40, 9.5, 0.28)
    add_rich_text(tb.text_frame.paragraphs[0], s.get("kicker", "研究生组会汇报"), SZ_KICKER, GRAY, bold=True)
    accent_rule(sl, 0.95, 1.90, 0.62, ACCENT, 3.0)

    title = s.get("title", "")
    size = 32 if len(title) <= 40 else (27 if len(title) <= 62 else 23)
    tb = textbox(sl, 0.95, 2.14, 11.3, 2.2)
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.line_spacing = 1.24
    add_rich_text(p, title, size, INK, bold=True)
    if s.get("subtitle"):
        p2 = para(tf)
        p2.space_before = Pt(12)
        p2.line_spacing = 1.3
        add_rich_text(p2, s["subtitle"], 15, GRAY)
    line_rule(sl, 0.95, 5.50, 4.2, RULE, 1.0)
    if s.get("meta"):
        tb = textbox(sl, 0.95, 5.66, 9.5, 1.0)
        tf = tb.text_frame
        for i, m in enumerate(s["meta"]):
            p = para(tf, first=(i == 0))
            p.space_after = Pt(4)
            add_rich_text(p, m, 13, GRAY)
    return sl


def slide_part(prs, s):
    """汇报种类分区页。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    kind = s.get("kind", "")
    num = s.get("roman") or KIND_ROMAN.get(kind, s.get("number", ""))
    if num:
        tb = textbox(sl, 0.95, 1.92, 3.2, 1.5)
        add_rich_text(tb.text_frame.paragraphs[0], str(num), 60, ACCENT, bold=True, font=FONT_EN)
    accent_rule(sl, 0.95, 3.46, 0.62, ACCENT, 3.0)
    tb = textbox(sl, 0.95, 3.68, 10.8, 0.9)
    add_rich_text(tb.text_frame.paragraphs[0], s.get("title", KIND_LABEL.get(kind, "")), 30, INK, bold=True)
    y = 4.62
    if s.get("subtitle"):
        tb = textbox(sl, 0.95, y, 10.8, 0.7)
        p = tb.text_frame.paragraphs[0]
        p.line_spacing = 1.3
        add_rich_text(p, s["subtitle"], 15, GRAY)
        y += 0.46
    items = s.get("items") or []
    if items:
        tb = textbox(sl, 0.95, y + 0.04, 10.8, min(2.1, 0.38 * len(items) + 0.24))
        tf = tb.text_frame
        for i, it in enumerate(items):
            p = para(tf, first=(i == 0))
            p.line_spacing = 1.25
            p.space_after = Pt(7)
            add_rich_text(p, it, 15, GRAY)
    return sl


def slide_agenda(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", "汇报内容"), s.get("kicker"))
    items = s.get("items", [])
    n = max(len(items), 1)
    row_h = min(1.05, (SH - 0.95 - top) / n)
    for i, it in enumerate(items):
        y = top + i * row_h
        tb = textbox(sl, MX, y + 0.12, 0.9, 0.5)
        add_rich_text(tb.text_frame.paragraphs[0], it.get("num", f"0{i+1}"), 22,
                      ACCENT if it.get("on", True) else LIGHT, bold=True, font=FONT_EN)
        tb = textbox(sl, MX + 1.0, y + 0.14, CW - 1.1, 0.36)
        add_rich_text(tb.text_frame.paragraphs[0], it.get("text", ""), 17, INK, bold=True)
        if it.get("note"):
            tb = textbox(sl, MX + 1.0, y + 0.52, CW - 1.1, 0.34)
            add_rich_text(tb.text_frame.paragraphs[0], it["note"], 12.5, GRAY)
        if i < n - 1:
            line_rule(sl, MX, y + row_h - 0.04, CW, RULE, 0.75)
    return sl


def _render_bullets(tf, bullets, size=SZ_BODY, first=True, gap0=8.0, gap1=5.0):
    idx = 0
    for item in bullets:
        if isinstance(item, str):
            text, level, bold = item, 0, False
        else:
            text = item.get("text", "")
            level = int(item.get("level", 0))
            bold = bool(item.get("bold", False))
        p = para(tf, first=(first and idx == 0))
        idx += 1
        p.line_spacing = 1.32
        p.space_after = Pt(gap0 if level == 0 else gap1)
        p.space_before = Pt(0)
        if level == 0:
            add_rich_text(p, text, size, BODY, bold=bold)
            set_bullet(p, "•", ACCENT, 0.28, 0.22)
        else:
            add_rich_text(p, text, max(size - 2, SZ_BODY_MIN), GRAY, bold=bold)
            set_bullet(p, "–", LIGHT, 0.58, 0.20)


def slide_bullets(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", ""), s.get("kicker"), s.get("sub"))
    takeaway = s.get("takeaway")
    bottom = SH - 1.58 if takeaway else SH - 0.82
    if takeaway:
        add_takeaway(sl, takeaway)
    bullets = s.get("bullets", [])
    box_h = bottom - top
    size = fit_bullet_size(bullets, CW, box_h, max_size=SZ_BODY + 4)
    gap = spread_gap(bullets, size, CW, box_h, cap=24.0)
    for w in widow_hits(bullets, size, CW):
        print(f"[hint] 「{s.get('title', '')[:14]}」末行只剩一两个字，建议改写措辞：{w}…")
    fill = est_bullets_h(bullets, size, CW, gap0=gap) / box_h
    if fill < 0.45:
        print(f"[warn] 「{s.get('title', '')[:16]}」内容偏少（约占版面 {fill:.0%}），"
              f"建议补充细节或与相邻页合并，避免半页空白")
    tb = textbox(sl, MX, top, CW, box_h)
    _render_bullets(tb.text_frame, bullets, size=size, gap0=gap)
    return sl


def slide_image(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", ""), s.get("kicker"))
    takeaway = s.get("takeaway")
    bottom = SH - 1.58 if takeaway else SH - 0.80
    if takeaway:
        add_takeaway(sl, takeaway)

    bullets = s.get("bullets") or []
    cap = s.get("caption", "")
    ratio = float(s.get("image_ratio", 0.60))
    if bullets:
        if s.get("image_side", "left") == "left":
            img_l, img_w = MX, CW * ratio
            txt_l, txt_w = MX + img_w + 0.42, CW - img_w - 0.42
        else:
            txt_l, txt_w = MX, CW * (1 - ratio) - 0.42
            img_l = MX + CW - CW * ratio
            img_w = CW * ratio
        box_h = bottom - top
        size = fit_bullet_size(bullets, txt_w, box_h)
        gap = spread_gap(bullets, size, txt_w, box_h, base=7.0, cap=13.0)
        for w in widow_hits(bullets, size, txt_w):
            print(f"[hint] 「{s.get('title', '')[:14]}」末行只剩一两个字，建议改写措辞：{w}…")
        tb = textbox(sl, txt_l, top - 0.02, txt_w, box_h)
        _render_bullets(tb.text_frame, bullets, size=size, gap0=gap, gap1=4.0)
    else:
        img_l, img_w = MX, CW

    cap_h = 0.50 if cap else 0.0
    px, py, pw, ph = fit_picture(sl, s["image"], img_l, top, img_w, bottom - top - cap_h,
                                 frame=s.get("frame", True),
                                 valign="top" if cap else "middle")
    if cap:
        add_caption(sl, cap, img_l, min(py + ph + 0.10, bottom - 0.40), img_w)
    return sl


def slide_table(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", ""), s.get("kicker"), s.get("sub"))
    takeaway = s.get("takeaway")
    bottom = SH - 1.58 if takeaway else SH - 0.82
    if takeaway:
        add_takeaway(sl, takeaway)
    draw_table(sl, s["columns"], s["rows"], top, bottom,
               widths=s.get("col_widths"), caption=s.get("caption"),
               total_ratio=float(s.get("width_ratio", 1.0)),
               highlight_col=s.get("highlight_col"))
    return sl


def slide_cards(prs, s):
    """事实卡片：左标签右内容，2–3 列。用于出处/实验目的/项目概况这类信息块。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", ""), s.get("kicker"), s.get("sub"))
    takeaway = s.get("takeaway")
    bottom = SH - 1.58 if takeaway else SH - 0.82
    if takeaway:
        add_takeaway(sl, takeaway)

    entries = s.get("entries", [])
    n = len(entries)
    cols = int(s.get("cols") or (2 if n > 3 else 1))
    cols = max(1, min(cols, 3))
    rows = math.ceil(n / cols)
    gap = 0.22
    card_w = (CW - gap * (cols - 1)) / cols
    avail = bottom - top
    card_h = min(2.30, (avail - gap * (rows - 1)) / rows)

    for i, e in enumerate(entries):
        r, c = divmod(i, cols)
        x = MX + c * (card_w + gap)
        y = top + r * (card_h + gap)
        wide = bool(e.get("wide"))
        w = card_w if not wide else CW
        panel(sl, x, y, w, card_h, fill=PANEL, bar=ACCENT if e.get("strong") else None)
        label = e.get("label", "")
        tb = textbox(sl, x + 0.26, y + 0.16, w - 0.46, 0.26)
        add_rich_text(tb.text_frame.paragraphs[0], label, SZ_KICKER, ACCENT, bold=True)
        val = e.get("value", "")
        vlen = len(_clean(val))
        size = 15 if vlen <= 60 else 14
        tb = textbox(sl, x + 0.26, y + 0.50, w - 0.46, card_h - 0.62)
        p = tb.text_frame.paragraphs[0]
        p.line_spacing = 1.30
        add_rich_text(p, val, size, INK if e.get("strong") else BODY, bold=bool(e.get("bold")))
    return sl


def slide_columns(prs, s):
    """2–4 栏对照：每栏一个小标题 + 要点。用于当前进度/问题/下一步这类并列信息。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    top = slide_title(sl, s.get("title", ""), s.get("kicker"), s.get("sub"))
    takeaway = s.get("takeaway")
    bottom = SH - 1.58 if takeaway else SH - 0.82
    if takeaway:
        add_takeaway(sl, takeaway)

    cols = s.get("columns", [])
    n = max(1, min(len(cols), 4))
    col_w = (CW - GUT * (n - 1)) / n
    fills = []
    for i, col in enumerate(cols[:n]):
        x = MX + i * (col_w + GUT)
        # 栏顶用细灰线分隔：若用强调色短线，会和页标题的强调线在左边缘叠成「双线」
        line_rule(sl, x, top, col_w - 0.06, RULE, 1.5)
        tb = textbox(sl, x, top + 0.16, col_w, 0.40)
        p = tb.text_frame.paragraphs[0]
        p.line_spacing = 1.15
        add_rich_text(p, col.get("heading", ""), 15, INK, bold=True)
        box_h = bottom - top - 0.66
        bl = col.get("bullets", [])
        size = fit_bullet_size(bl, col_w, box_h, max_size=(SZ_BODY + 1) if n <= 3 else SZ_BODY)
        gap = spread_gap(bl, size, col_w, box_h, base=7.0, cap=14.0)
        for w in widow_hits(bl, size, col_w):
            print(f"[hint] 「{s.get('title', '')[:14]}」第{i + 1}栏末行只剩一两个字：{w}…")
        fills.append(est_bullets_h(bl, size, col_w, gap0=gap) / box_h if box_h else 0)
        tb = textbox(sl, x, top + 0.66, col_w, box_h)
        _render_bullets(tb.text_frame, bl, size=size, gap0=gap, gap1=4.0)
    if fills and sum(fills) / len(fills) < 0.42:
        print(f"[warn] 「{s.get('title', '')[:16]}」各栏内容偏少（约占 {sum(fills) / len(fills):.0%}），"
              f"建议每栏补一条具体结果或数据")
    return sl


def slide_statement(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    accent_rule(sl, 1.55, 2.18, 0.60, ACCENT, 3.0)
    tb = textbox(sl, 1.55, 2.42, SW - 3.0, 2.8, anchor=MSO_ANCHOR.TOP)
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.line_spacing = 1.38
    add_rich_text(p, s.get("text", ""), 25, INK, bold=True)
    if s.get("attribution"):
        p2 = para(tf)
        p2.space_before = Pt(20)
        add_rich_text(p2, s["attribution"], SZ_NOTE, GRAY)
    return sl


def slide_end(prs, s):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    accent_rule(sl, 1.0, 3.12, 0.62, ACCENT, 3.0)
    tb = textbox(sl, 1.0, 3.36, SW - 2.0, 1.0)
    add_rich_text(tb.text_frame.paragraphs[0], s.get("title", "谢谢，请老师和同学批评指正"), 28, INK, bold=True)
    if s.get("subtitle"):
        tb = textbox(sl, 1.0, 4.46, SW - 2.0, 0.9)
        p = tb.text_frame.paragraphs[0]
        p.line_spacing = 1.4
        add_rich_text(p, s["subtitle"], SZ_NOTE, GRAY)
    return sl


BUILDERS = {
    "cover": slide_cover,
    "part": slide_part,
    "agenda": slide_agenda,
    "bullets": slide_bullets,
    "image": slide_image,
    "image_only": slide_image,
    "table": slide_table,
    "cards": slide_cards,
    "columns": slide_columns,
    "statement": slide_statement,
    "end": slide_end,
}


# ---------------------------------------------------------------- 结构展开
def expand(deck: dict) -> list[dict]:
    """把 part/unit 三级结构摊平成带上下文的页面列表。"""
    plan: list[dict] = []

    cover = dict(deck.get("cover") or {})
    cover.setdefault("type", "cover")
    cover.setdefault("title", deck.get("meta", {}).get("title", "组会汇报"))
    if "meta" not in cover:
        m = deck.get("meta") or {}
        if m.get("cover_meta"):
            cover["meta"] = list(m["cover_meta"])
        elif m.get("presenter") or m.get("date"):
            cover["meta"] = [x for x in (m.get("presenter", ""), m.get("date", "")) if x]
    plan.append({"kind": "cover", "page": cover, "ctx": {}})

    parts = deck.get("parts") or []
    if deck.get("agenda"):
        items = []
        for i, p in enumerate(parts, 1):
            units = p.get("units") or []
            note = p.get("agenda_note")
            if not note and units:
                names = [u.get("name", "") for u in units]
                note = "；".join(names[:3]) + ("…" if len(names) > 3 else "")
            items.append({"num": f"0{i}", "text": KIND_LABEL.get(p.get("kind"), p.get("title", "")),
                          "note": note, "on": True})
        plan.append({"kind": "agenda",
                     "page": {"type": "agenda", "title": "汇报内容", "items": items},
                     "ctx": {}})

    for pi, part in enumerate(parts, 1):
        kind = part.get("kind", "")
        label = KIND_LABEL.get(kind, part.get("title", ""))
        units = part.get("units") or []
        plan.append({
            "kind": "part",
            "page": {
                "type": "part",
                "kind": kind,
                "title": part.get("title") or label,
                "subtitle": part.get("subtitle"),
                "items": part.get("items") or [
                    f"{j:02d}　{u.get('name', '')}" for j, u in enumerate(units, 1)
                ] or None,
            },
            "ctx": {"part": pi, "part_label": label},
        })
        for page in part.get("pages") or []:
            plan.append({"kind": "content", "page": page,
                         "ctx": {"part": pi, "part_label": label, "part_page": True}})

        n_units = len(units)
        for ui, unit in enumerate(units, 1):
            ctx = {"part": pi, "part_label": label,
                   "unit": ui, "unit_total": n_units, "unit_name": unit.get("name", "")}
            if unit.get("divider"):
                plan.append({"kind": "part", "page": {
                    "type": "part", "kind": kind,
                    "title": unit.get("name", ""),
                    "subtitle": unit.get("subtitle") or f"{label} · 第 {ui} / {n_units} 部分",
                    "roman": f"{ui:02d}",
                    "items": None,
                }, "ctx": ctx})
            for page in unit.get("pages") or []:
                page = dict(page)
                if not page.get("kicker"):
                    page["kicker"] = f"{label} · {ui}/{n_units}"
                plan.append({"kind": "content", "page": page, "ctx": ctx})
    end = dict(deck.get("end") or {})
    end.setdefault("type", "end")
    plan.append({"kind": "end", "page": end, "ctx": {}})
    return plan


def check_budget(deck: dict) -> list[str]:
    errs = []
    for pi, part in enumerate(deck.get("parts") or [], 1):
        kind = part.get("kind", "")
        label = KIND_LABEL.get(kind, f"第{pi}部分")
        units = part.get("units") or []
        if not units and not part.get("pages"):
            errs.append(f"{label}：既没有 units 也没有 pages")
        for ui, unit in enumerate(units, 1):
            n = len(unit.get("pages") or [])
            if not (BUDGET_MIN <= n <= BUDGET_MAX):
                errs.append(f"{label} 第 {ui} 个单元「{unit.get('name', '')[:18]}」内容页 {n} 页，"
                            f"要求 {BUDGET_MIN}–{BUDGET_MAX} 页")
    return errs


def count_chars(page: dict) -> int:
    n = len(_clean(page.get("title", "")))
    for b in page.get("bullets") or []:
        n += len(_clean(b if isinstance(b, str) else b.get("text", "")))
    for e in page.get("entries") or []:
        n += len(_clean(e.get("value", "")))
    for col in page.get("columns") or []:
        n += len(_clean(col.get("heading", "")))
        n += sum(len(_clean(b if isinstance(b, str) else b.get("text", ""))) for b in col.get("bullets", []))
    for r in page.get("rows") or []:
        n += sum(len(_clean(c)) for c in r)
    return n


def main() -> int:
    ap = argparse.ArgumentParser(description="按 JSON 规格生成研究生组会 PPT")
    ap.add_argument("spec", help="deck.json 路径")
    ap.add_argument("-o", "--out", required=True, help="输出 .pptx 路径")
    ap.add_argument("--accent", help="覆盖强调色，如 1F6F5C")
    ap.add_argument("--skip-budget-check", action="store_true", help="允许单元页数越界（不推荐）")
    args = ap.parse_args()

    spec = json.loads(Path(args.spec).read_text(encoding="utf-8"))
    global ACCENT, ACCENT_SOFT
    if args.accent:
        ACCENT = args.accent.lstrip("#")

    if not spec.get("parts"):
        print("[error] parts 为空：没有任何汇报内容", file=sys.stderr)
        return 2

    budget_errs = check_budget(spec)
    if budget_errs and not args.skip_budget_check:
        print("[error] 每个单元必须 2–4 页内容，以下不合规：", file=sys.stderr)
        for e in budget_errs:
            print(f"  - {e}", file=sys.stderr)
        print("处理：合并/拆分页面后再生成；临时放过加 --skip-budget-check。", file=sys.stderr)
        return 5
    for e in budget_errs:
        print(f"[warn] {e}")

    plan = expand(spec)
    slides = [p for p in plan if p["kind"] != "end"] + [p for p in plan if p["kind"] == "end"]

    # 图片路径：绝对路径直接用；相对路径先按 deck.json 所在目录解析，再退回当前目录
    base = Path(args.spec).resolve().parent

    def resolve(p):
        raw = Path(str(p))
        if raw.is_absolute() and raw.exists():
            return raw
        cand = (base / raw)
        return cand if cand.exists() else raw

    for item in slides:
        page = item["page"]
        if page.get("image"):
            page["image"] = str(resolve(page["image"]))

    missing = [p["page"].get("image") for p in slides
               if p["kind"] == "content" and p["page"].get("type") in ("image", "image_only")
               and not Path(str(p["page"].get("image", ""))).exists()]
    bad_types = [(i, p["page"].get("type")) for i, p in enumerate(slides, 1)
                 if p["page"].get("type") not in BUILDERS]
    if bad_types:
        print(f"[error] 未知版式：{bad_types}", file=sys.stderr)
        return 2
    if missing:
        print("[error] 这些图片不存在，先把原文图表裁好再来：", file=sys.stderr)
        for m in missing:
            print(f"  - {m}", file=sys.stderr)
        print("提示：截图用 crop_figure.py；路径写绝对路径或相对 deck.json 的路径。", file=sys.stderr)
        return 3

    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    total = len(slides)
    meta = spec.get("meta") or {}
    warn_count = 0

    for i, item in enumerate(slides, 1):
        page, ctx = item["page"], item["ctx"]
        fn = BUILDERS[page["type"]]
        sl = fn(prs, page)
        if item["kind"] == "content":
            foot = meta.get("footer", "") or meta.get("title", "")
            seg = [foot]
            if page.get("source"):
                seg.append(page["source"])
            add_footer(sl, "　·　".join(s for s in seg if s), i, total)
            n = count_chars(page) if page["type"] == "bullets" else 0
            if n > 260:
                print(f"[warn] P{i} 正文 {n} 字，偏密，建议拆页或删句")
                warn_count += 1
        if page.get("notes"):
            sl.notes_slide.notes_text_frame.text = str(page["notes"])

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))

    # 结构摘要
    n_units = sum(len(p.get("units") or []) for p in spec["parts"])
    print(f"[ok] 已生成 {out}（{total} 页，{out.stat().st_size // 1024} KB）")
    print(f"[结构] {len(spec['parts'])} 个汇报种类 / {n_units} 个单元")
    for pi, part in enumerate(spec["parts"], 1):
        label = KIND_LABEL.get(part.get("kind"), part.get("title", ""))
        units = part.get("units") or []
        pages = [len(u.get("pages") or []) for u in units]
        print(f"  {pi}. {label}：{len(units)} 个单元，内容页 {pages or '-'}"
              f"，附part页 {len(part.get('pages') or [])}")
    print(f"[下一步] python check_structure.py {args.spec} → python check_deck.py {out.name} → 渲染目检")
    if warn_count:
        print(f"[注意] {warn_count} 处偏密提示，逐条处理后重新生成")
    return 0


if __name__ == "__main__":
    sys.exit(main())
