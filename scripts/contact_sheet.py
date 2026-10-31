#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""contact_sheet.py — 把成套截图拼成一张审片图，一次看完所有图表的裁切质量。

用法：
    python contact_sheet.py figs/*.png -o sheet.png
    python contact_sheet.py figs/*.png -o sheet.png --cell 520

每格标注：序号、文件名、像素尺寸、估算 dpi；四条边若有墨迹跨过边界（疑似裁断）会把格子
边框染红并标 "EDGE"。看这一张图就能发现「表头被切掉一半」「坐标轴断掉」「图注缺一行」。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import numpy as np
    from PIL import Image, ImageDraw, ImageFont
except ImportError:  # pragma: no cover
    sys.exit("需要 Pillow 与 numpy：pip install pillow numpy")

INK_THRESHOLD = 245
HIGH_INK = 20
BG = (246, 247, 249)
INK = (26, 26, 26)
GRAY = (110, 118, 128)
RED = (198, 63, 63)


def edge_touched(gray: np.ndarray) -> bool:
    m = gray < INK_THRESHOLD
    hit = [int(m[:, 0].sum()), int(m[:, -1].sum()), int(m[0, :].sum()), int(m[-1, :].sum())]
    return any(v >= HIGH_INK for v in hit)


def load_font(size: int):
    for p in (r"C:\Windows\Fonts\msyh.ttc", "/System/Library/Fonts/PingFang.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"):
        if Path(p).exists():
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def main() -> int:
    ap = argparse.ArgumentParser(description="把多张截图拼成审片图")
    ap.add_argument("images", nargs="+")
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("--cell", type=int, default=520, help="每格图片最大宽度（像素），默认 520")
    ap.add_argument("--cols", type=int, default=2, help="列数，默认 2")
    ap.add_argument("--target-in", type=float, default=8.0, help="估算 dpi 的假定放置宽度（英寸）")
    args = ap.parse_args()

    files = [Path(p) for p in args.images if Path(p).is_file()]
    if not files:
        print("[error] 没有可读的图片", file=sys.stderr)
        return 3

    f_label = load_font(17)
    f_meta = load_font(14)
    pad, head = 16, 46
    cols = max(1, args.cols)
    rows = (len(files) + cols - 1) // cols
    cw, ch = args.cell + pad * 2, args.cell * 3 // 4 + head + pad * 2
    sheet = Image.new("RGB", (cols * cw, rows * ch), BG)
    draw = ImageDraw.Draw(sheet)

    for i, path in enumerate(files):
        r, c = divmod(i, cols)
        ox, oy = c * cw, r * ch
        img = Image.open(path).convert("RGB")
        w0, h0 = img.size
        touched = edge_touched(np.asarray(img.convert("L"), dtype=np.uint8))
        thumb = img.copy()
        thumb.thumbnail((args.cell, args.cell * 3 // 4), Image.LANCZOS)
        x = ox + pad + (args.cell - thumb.width) // 2
        y = oy + head + (args.cell * 3 // 4 - thumb.height) // 2
        dpi = w0 / args.target_in
        draw.rectangle([x - 1, y - 1, x + thumb.width, y + thumb.height],
                       outline=RED if touched else (209, 214, 221),
                       width=2 if touched else 1)
        sheet.paste(thumb, (x, y))
        title = f"{i + 1:02d}  {path.name}"
        if touched:
            title += "   [EDGE 疑似裁断]"
        draw.text((ox + pad, oy + 10), title, fill=RED if touched else INK, font=f_label)
        draw.text((ox + pad, oy + 30),
                  f"{w0}×{h0}px   ≈{dpi:.0f}dpi @{args.target_in:g}in   "
                  f"宽高比 {w0 / h0:.2f}",
                  fill=RED if dpi < 110 else GRAY, font=f_meta)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    bad = sum(1 for p in files if edge_touched(
        np.asarray(Image.open(p).convert("L"), dtype=np.uint8)))
    print(f"[ok] 审片图 {out}（{len(files)} 张，{cols}×{rows} 格）")
    if bad:
        print(f"[warn] {bad} 张疑似裁断（红框标注），逐张放大确认后按 crop_figure.py 的建议坐标重裁")
    else:
        print("[ok] 全部图片四边无墨迹跨过边界")
    return 0


if __name__ == "__main__":
    sys.exit(main())
