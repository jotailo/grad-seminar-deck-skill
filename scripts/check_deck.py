#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""check_deck.py — 对已生成的 PPTX 做静态体检：溢出、字号、遮挡、越界、插图清晰度。

用法：
    python check_deck.py 组会汇报.pptx
    python check_deck.py 组会汇报.pptx --json report.json
    python check_deck.py 组会汇报.pptx --min-body 14 --min-any 10

这是「用眼睛看」之前的机器关口。它抓不住审美问题，但能可靠抓住这几类硬伤：
字被框吃掉、字号太小、图和图注压在一起、元素出页面、截图分辨率不够发虚、空白页。
换行估算用本机真实字体度量（微软雅黑），中文按逐字折行。

退出码：0 通过；4 有 FAIL；3 文件打不开。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.oxml.ns import qn
except ImportError:  # pragma: no cover
    sys.exit("需要 python-pptx：pip install python-pptx")

try:
    from PIL import ImageFont
except ImportError:  # pragma: no cover
    ImageFont = None

EMU_PER_IN = 914400
FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
]
MIN_IMG_DPI = 110        # 低于此值投影会发虚
GOOD_IMG_DPI = 150


def pick_font():
    for p in FONT_CANDIDATES:
        if Path(p).exists():
            return p
    return None


class Measurer:
    def __init__(self):
        self.font_path = pick_font()
        self.cache = {}

    def width_pt(self, text: str, size_pt: float) -> float:
        if ImageFont is None or not self.font_path:
            return sum(1.0 if ord(c) > 0x2E80 else 0.54 for c in text) * size_pt
        key = int(round(size_pt * 4))
        f = self.cache.get(key)
        if f is None:
            f = ImageFont.truetype(self.font_path, max(6, int(size_pt)))
            self.cache[key] = f
        return f.getlength(text)

    def wrap_lines(self, text: str, size_pt: float, box_pt: float) -> int:
        if not text:
            return 1
        box_pt = max(box_pt, 8)
        lines, cur = 1, 0.0
        for ch in text:
            if ch == "\n":
                lines += 1
                cur = 0.0
                continue
            w = self.width_pt(ch, size_pt)
            if cur + w > box_pt:
                lines += 1
                cur = w
            else:
                cur += w
        return lines


def para_spec(p, default_size=18.0):
    text = "".join(r.text or "" for r in p.runs)
    sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
    size = max(sizes) if sizes else default_size
    ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.2
    return text, size, ls


def para_indent_pt(p) -> float:
    """段落左缩进（marL），换行后实际可用宽度要扣掉它，否则会低估行数。"""
    pPr = p._p.find(qn("a:pPr"))
    if pPr is None:
        return 0.0
    try:
        return int(pPr.get("marL", 0)) / 12700.0
    except (TypeError, ValueError):
        return 0.0


def tf_height_pt(tf, box_width_pt, m: Measurer):
    total, max_size = 0.0, 0.0
    for p in tf.paragraphs:
        text, size, ls = para_spec(p)
        max_size = max(max_size, size)
        n = m.wrap_lines(text, size, max(box_width_pt - para_indent_pt(p), 12))
        h = n * size * ls * 1.16
        if p.space_before is not None:
            h += p.space_before.pt
        if p.space_after is not None:
            h += p.space_after.pt
        total += h
    return total, max_size


def bbox_in(shape):
    return (shape.left / EMU_PER_IN, shape.top / EMU_PER_IN,
            (shape.left + shape.width) / EMU_PER_IN, (shape.top + shape.height) / EMU_PER_IN)


def overlap_area(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(0.0, w) * max(0.0, h)


def contains(outer, inner, tol=0.02):
    return (outer[0] - tol <= inner[0] and outer[1] - tol <= inner[1]
            and outer[2] + tol >= inner[2] and outer[3] + tol >= inner[3])


def main() -> int:
    ap = argparse.ArgumentParser(description="PPTX 静态体检")
    ap.add_argument("pptx")
    ap.add_argument("--json")
    ap.add_argument("--quiet-ok", action="store_true")
    ap.add_argument("--min-body", type=float, default=14.0, help="正文字号下限（pt）")
    ap.add_argument("--min-any", type=float, default=10.0, help="任何文字的字号下限（pt）")
    args = ap.parse_args()

    path = Path(args.pptx)
    if not path.exists():
        print(f"[error] 找不到文件：{path}", file=sys.stderr)
        return 3
    prs = Presentation(str(path))
    W = prs.slide_width / EMU_PER_IN
    H = prs.slide_height / EMU_PER_IN
    m = Measurer()

    fails, warns, info = [], [], []
    sizes = {}

    for si, slide in enumerate(prs.slides, 1):
        boxes, has_content = [], False
        for shape in slide.shapes:
            bb = bbox_in(shape)
            if bb[0] < -0.05 or bb[1] < -0.05 or bb[2] > W + 0.05 or bb[3] > H + 0.05:
                fails.append(f"P{si} 元素越出页面：{shape.shape_type} @ {tuple(round(v, 2) for v in bb)}")

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                has_content = True
                try:
                    w_in = shape.width / EMU_PER_IN
                    px, py_ = shape.image.size[0], shape.image.size[1]
                    dpi = px / w_in if w_in else 999
                    if dpi < MIN_IMG_DPI:
                        fails.append(f"P{si} 插图分辨率不足：{px}px 放到 {w_in:.1f}in ≈ {dpi:.0f}dpi"
                                     f"（< {MIN_IMG_DPI}）会发虚，请用更高 dpi 重渲染页面后重裁")
                    elif dpi < GOOD_IMG_DPI:
                        warns.append(f"P{si} 插图 {dpi:.0f}dpi 偏紧：投影尚可，打印会糊")
                    else:
                        info.append(f"P{si} 插图 {px}×{py_}px → {dpi:.0f}dpi 良好")
                    if w_in < 1.4:
                        warns.append(f"P{si} 插图仅 {w_in:.1f}in 宽，观众看不清细节，考虑单独占一页")
                except Exception:
                    pass

            if shape.has_text_frame and shape.text_frame.text.strip():
                has_content = True
                tf = shape.text_frame
                mw = (shape.width - tf.margin_left - tf.margin_right) / 12700
                mh = (shape.height - tf.margin_top - tf.margin_bottom) / 12700
                need, _ = tf_height_pt(tf, mw, m)
                if need > mh * 1.04 and need - mh > 4:
                    fails.append(f"P{si} 文字可能溢出：「{tf.text.strip()[:22]}…」"
                                 f"需 {need:.0f}pt 高，框仅 {mh:.0f}pt")
                run_sizes = [r.font.size.pt for p in tf.paragraphs for r in p.runs if r.font.size]
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.size:
                            k = round(r.font.size.pt, 1)
                            sizes[k] = sizes.get(k, 0) + len(r.text or "")
                if (run_sizes and shape.height / EMU_PER_IN >= 0.6
                        and len(tf.text.strip()) >= 40 and max(run_sizes) < args.min_body):
                    fails.append(f"P{si} 正文块字号仅 {max(run_sizes)}pt（{len(tf.text.strip())} 字），投影偏小")
                boxes.append((bb, "text", tf.text.strip()[:18]))

            if shape.has_table:
                has_content = True
                tbl = shape.table
                for ri, row in enumerate(tbl.rows):
                    for ci, cell in enumerate(row.cells):
                        tf = cell.text_frame
                        cw = (tbl.columns[ci].width - cell.margin_left - cell.margin_right) / 12700
                        chh = (row.height - cell.margin_top - cell.margin_bottom) / 12700
                        need, _ = tf_height_pt(tf, cw, m)
                        if need > chh * 1.18 + 4:
                            fails.append(f"P{si} 表格第{ri + 1}行第{ci + 1}列文字溢出："
                                         f"需 {need:.0f}pt，行高 {chh:.0f}pt")
                        for p in tf.paragraphs:
                            for r in p.runs:
                                if r.font.size:
                                    k = round(r.font.size.pt, 1)
                                    sizes[k] = sizes.get(k, 0) + len(r.text or "")
                boxes.append((bb, "media", ""))

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                boxes.append((bb, "media", ""))

        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                if a[1] == "text" and b[1] == "text":
                    continue
                if contains(a[0], b[0]) or contains(b[0], a[0]):
                    continue
                ov = overlap_area(a[0], b[0])
                if ov > 0.06:
                    fails.append(f"P{si} 疑似遮挡：{a[1]} 与 {b[1]} 重叠 {ov:.2f} in²"
                                 f"（图/表被文字压住或反之）")
        if not has_content and si > 1:
            warns.append(f"P{si} 这一页没有任何文字或图片，确认不是空页")

    for size, cnt in sorted(sizes.items()):
        if size < args.min_any:
            fails.append(f"存在 {size}pt 文字（{cnt} 字），过小，投影必看不清")
        else:
            info.insert(0, f"{size}pt × {cnt} 字")

    print(f"=== 体检报告：{path.name}（{len(prs.slides)} 页，{W:.2f}×{H:.2f}in）===")
    if not args.quiet_ok or fails or warns:
        print("[字号分布] " + " | ".join(i for i in info if "pt ×" in i))
        for line in [i for i in info if "插图" in i]:
            print("  " + line)
    for w in warns:
        print(f"[WARN] {w}")
    for f in fails:
        print(f"[FAIL] {f}")
    print(f"结论：{'通过' if not fails else f'{len(fails)} 项需修复'}"
          f"{'' if not warns else f'，{len(warns)} 项提醒'}")
    print("[下一步] 渲染成图片逐页目检：内容是否被遮挡、图注是否完整、字是否清晰"
          "（scripts/render_deck.ps1）")

    if args.json:
        Path(args.json).write_text(json.dumps(
            {"file": str(path), "slides": len(prs.slides), "fails": fails, "warns": warns},
            ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if not fails else 4


if __name__ == "__main__":
    sys.exit(main())
