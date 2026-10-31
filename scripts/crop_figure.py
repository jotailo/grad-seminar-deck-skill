#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""crop_figure.py — 从整页渲染图裁切原文图表，并判断「有没有被裁断」。

用法：
    # 1) 先让脚本给出候选框（整页图上的内容块），再据此填 --bbox
    python crop_figure.py --page p05.png --propose
    # 2) 按框裁切（坐标是 0~1 相对比例，左上角为 0,0）
    python crop_figure.py --page p05.png --bbox 0.06,0.10,0.48,0.36 --out fig1.png
    # 3) 单独复查一张已裁好的图
    python crop_figure.py --check fig1.png

坐标一律用相对比例，因为看图估位置的模型给的就是比例；也接受像素值（>1 即视为像素）。

设计取向（别改成"聪明"的自动裁切）：
  裁切只做「按给定框 + 固定留白」，绝不自动向内收。因为向内收一定丢内容——
  图题在表格下方、坐标轴标签在图左侧，任何"找整齐缝隙"的启发式都会把它们切掉。
  若某条边上有墨迹跨过边界，脚本把「该边最近的空白位置」换算成建议比例打印出来，
  让修正一次到位，而不是盲试。

--propose 只给候选框，必须看图确认后再裁：它会按空白带切块，遇到跨栏大图或多子图会切错。

退出码：0 正常；2 参数错误；3 输入不可读；4 边缘有墨迹（疑似裁断，需按建议调整或看图确认）。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

try:
    import numpy as np
    from PIL import Image
except ImportError:  # pragma: no cover
    sys.exit("需要 Pillow 与 numpy：pip install pillow numpy")

INK_THRESHOLD = 245   # 灰度低于此值算墨迹
EDGE_BAND = 2         # 判定「贴近边界」的带宽（像素）
HINT_RANGE = 200      # 给建议坐标时的搜索半径（像素）
MIN_GAP = 6           # 认为「已经是空白」的最小空白宽度（像素）
HIGH_INK = 20         # 边界墨迹超过此像素数才判「真被切开」；小于它多半是文字正好写到版心


def load_gray(path: Path):
    img = Image.open(path).convert("RGB")
    return img, np.asarray(img.convert("L"), dtype=np.uint8)


def edge_report(gray: np.ndarray, box) -> dict:
    x0, y0, x1, y1 = box
    sub = gray[y0:y1, x0:x1]
    if sub.size == 0:
        return {"ok": False, "touched": {}, "empty": True}
    m = sub < INK_THRESHOLD
    # 「被裁断」的判据是墨迹跨过边界（最外一行/列有墨迹），而不是「靠近边界」。
    # 靠近边界通常是内容正好压边，正常；跨过边界才会出现半个字或断掉的坐标轴。
    hit = {"left": int(m[:, 0].sum()), "right": int(m[:, -1].sum()),
           "top": int(m[0, :].sum()), "bottom": int(m[-1, :].sum())}
    near = {"left": int(m[:, :EDGE_BAND].sum()), "right": int(m[:, -EDGE_BAND:].sum()),
            "top": int(m[:EDGE_BAND, :].sum()), "bottom": int(m[-EDGE_BAND:, :].sum())}
    hard = {k: v for k, v in hit.items() if v >= HIGH_INK}
    soft = {k: v for k, v in hit.items() if 0 < v < HIGH_INK}
    return {"ok": not hard, "touched": hard, "soft": soft,
            "near": {k: v for k, v in near.items() if v and not hit[k]}}


def blank_runs(profile: np.ndarray):
    runs, start = [], None
    for i, has in enumerate(profile):
        if not has:
            if start is None:
                start = i
        elif start is not None:
            runs.append((start, i))
            start = None
    if start is not None:
        runs.append((start, profile.size))
    return runs


def suggest(gray: np.ndarray, box, side: str):
    """某条边有墨迹时，向外找最近空白带，换算成建议的相对比例。"""
    x0, y0, x1, y1 = box
    H, W = gray.shape
    if side in ("left", "right"):
        prof = (gray[y0:y1] < INK_THRESHOLD).any(axis=0)
        pos = x0 if side == "left" else x1
        outward = -1 if side == "left" else 1
    else:
        prof = (gray[:, x0:x1] < INK_THRESHOLD).any(axis=1)
        pos = y0 if side == "top" else y1
        outward = -1 if side == "top" else 1
    best = None
    for s, e in blank_runs(prof):
        cand = e if outward < 0 else s
        d = (pos - cand) * outward
        if d < 0 or d > HINT_RANGE or e - s < MIN_GAP:
            continue
        if best is None or d < best[0]:
            best = (d, cand)
    if best is None:
        return None
    total = W if side in ("left", "right") else H
    return round(best[1] / total, 3)


def propose(gray: np.ndarray) -> list[dict]:
    """按空白带把整页切成内容块，给出候选相对框（需看图确认）。"""
    H, W = gray.shape
    ink = gray < INK_THRESHOLD
    rows = ink.any(axis=1)
    bands, start = [], None
    for i, has in enumerate(rows):
        if has and start is None:
            start = i
        elif not has and start is not None:
            bands.append((start, i))
            start = None
    if start is not None:
        bands.append((start, H))

    merged = []
    for b in bands:
        if merged and b[0] - merged[-1][1] < int(0.012 * H):
            merged[-1] = (merged[-1][0], b[1])
        else:
            merged.append(list(b))

    out = []
    for (y0, y1) in merged:
        if y1 - y0 < int(0.018 * H):
            continue
        sub = ink[y0:y1]
        cols = sub.any(axis=0)
        gaps = [(s, e) for s, e in blank_runs(cols) if e - s > int(0.05 * W) and s > int(0.06 * W)]
        mid = W / 2
        split = min(gaps, key=lambda g: abs((g[0] + g[1]) / 2 - mid)) if gaps else None
        segs = [(0, W)] if not split else [(0, split[0]), (split[1], W)]
        for (x0, x1) in segs:
            if x1 - x0 < int(0.06 * W):
                continue
            block = ink[y0:y1, x0:x1]
            ys, xs = np.where(block.any(axis=1))[0], np.where(block.any(axis=0))[0]
            if ys.size == 0 or xs.size == 0:
                continue
            bx0, bx1 = x0 + xs[0], x0 + xs[-1] + 1
            by0, by1 = y0 + ys[0], y0 + ys[-1] + 1
            if (bx1 - bx0) * (by1 - by0) < 0.012 * W * H:
                continue
            out.append({"bbox": [round(bx0 / W, 3), round(by0 / H, 3),
                                 round(bx1 / W, 3), round(by1 / H, 3)],
                        "w": bx1 - bx0, "h": by1 - by0, "ink": int(block.sum())})
    out.sort(key=lambda d: (d["bbox"][1], d["bbox"][0]))
    return out


def check_file(path: Path, target_in: float = 8.0) -> int:
    img, gray = load_gray(path)
    w, h = img.size
    rep = edge_report(gray, [0, 0, w, h])
    print(f"[check] {path.name}: {w}×{h}px；按放到 {target_in:g}in 宽 ≈ {w / target_in:.0f}dpi")
    if w / target_in < 110:
        print("  [WARN] 分辨率偏低，投到大屏会发虚：用 --dpi 400 重新渲染页面后重裁")
    if w / h > 4.2:
        print("  [WARN] 宽高比 > 4:1，疑似只截到一行/一个表头，确认上方或下方没有漏掉内容")
    if not rep["ok"]:
        print(f"  [FAIL] 墨迹触边，疑似被裁断：{rep['touched']}")
        print("  处理：把该边坐标向外放 0.02~0.04 重裁；若内容本身就到版面边缘，属正常")
        return 4
    print("  [OK] 四边均无墨迹跨过边界，边缘完整")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="从整页渲染图裁切原文图表（含完整性判定）")
    ap.add_argument("--page", help="整页渲染图（PNG/JPG）")
    ap.add_argument("--bbox", help="裁切框：相对比例 x0,y0,x1,y1（0~1）；也可给像素值")
    ap.add_argument("--out", help="输出 PNG 路径")
    ap.add_argument("--pad", type=int, default=14, help="四边统一留白（像素），默认 14")
    ap.add_argument("--trim", action="store_true", help="先把框贴紧内容（去掉框内多余空白）")
    ap.add_argument("--max-width", type=int, default=2400, help="输出最宽像素，超出等比缩小")
    ap.add_argument("--propose", action="store_true", help="只打印候选框，不裁切")
    ap.add_argument("--check", help="仅对已有图片做完整性体检")
    ap.add_argument("--target-in", type=float, default=8.0, help="估算 dpi 时假定的放置宽度（英寸）")
    args = ap.parse_args()

    if args.check:
        return check_file(Path(args.check), args.target_in)
    if not args.page:
        ap.error("需要 --page（或用 --check）")
    page = Path(args.page)
    if not page.exists():
        print(f"[error] 找不到页面图：{page}", file=sys.stderr)
        return 3
    img, gray = load_gray(page)
    W, H = img.size

    if args.propose:
        cands = propose(gray)
        print(f"[propose] {page.name} {W}×{H}px，候选内容块 {len(cands)} 个（从上到下）：")
        print("  注意：这是按空白带切的粗框，跨栏大图/多子图会被切错，必须看图确认后再用。")
        for i, c in enumerate(cands, 1):
            x0, y0, x1, y1 = c["bbox"]
            print(f"  {i:02d}  --bbox {x0},{y0},{x1},{y1}   "
                  f"（{c['w']}×{c['h']}px，占比 {(x1 - x0):.0%}×{(y1 - y0):.0%}）")
        print("  用法：选一个候选框，先用 --pad 16 裁出来看，再按需微调四边。")
        return 0

    if not (args.bbox and args.out):
        ap.error("需要 --bbox 与 --out（或用 --propose / --check）")

    try:
        vals = [float(v) for v in args.bbox.replace(" ", "").split(",")]
        if len(vals) != 4:
            raise ValueError
    except ValueError:
        print("[error] --bbox 需为 x0,y0,x1,y1 四个数字", file=sys.stderr)
        return 2

    if max(vals) <= 1.0:
        box = [int(round(vals[0] * W)), int(round(vals[1] * H)),
               int(round(vals[2] * W)), int(round(vals[3] * H))]
    else:
        box = [int(v) for v in vals]
    box = [max(0, min(box[0], W - 2)), max(0, min(box[1], H - 2)),
           max(box[0] + 2, min(box[2], W)), max(box[1] + 2, min(box[3], H))]

    if args.trim:
        m = gray[box[1]:box[3], box[0]:box[2]] < INK_THRESHOLD
        rows, cols = np.where(m.any(axis=1))[0], np.where(m.any(axis=0))[0]
        if rows.size and cols.size:
            box = [box[0] + int(cols[0]), box[1] + int(rows[0]),
                   box[0] + int(cols[-1]) + 1, box[1] + int(rows[-1]) + 1]

    if args.pad:
        box = [max(0, box[0] - args.pad), max(0, box[1] - args.pad),
               min(W, box[2] + args.pad), min(H, box[3] + args.pad)]

    crop = img.crop(tuple(box))
    if crop.width > args.max_width:
        ratio = args.max_width / crop.width
        crop = crop.resize((args.max_width, max(1, int(crop.height * ratio))), Image.LANCZOS)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    crop.save(out)

    print(f"[crop] 页面 {W}×{H} → 框 {box}（{box[2] - box[0]}×{box[3] - box[1]}px）")
    print(f"[crop] 相对比例 x0={box[0] / W:.3f} y0={box[1] / H:.3f} "
          f"x1={box[2] / W:.3f} y1={box[3] / H:.3f}")
    print(f"[crop] 输出 {out}（{crop.width}×{crop.height}px，按 {args.target_in:g}in 宽放置 "
          f"≈ {crop.width / args.target_in:.0f}dpi）")

    rep = edge_report(gray, box)
    if rep["ok"]:
        if rep.get("soft"):
            print(f"[OK] 无内容被切开（{list(rep['soft'])} 边有少量墨迹贴边，多为文字正好写到版心，看图确认即可）")
        elif rep.get("near"):
            print(f"[OK] 无墨迹跨过边界（{list(rep['near'])} 边内容贴近边界，属正常压边）")
        else:
            print("[OK] 四边无墨迹跨过边界，边缘完整")
        return 0

    print(f"[FAIL] 这些边的墨迹跨过了边界，内容被切开：{rep['touched']}")
    fixes = []
    for side in rep["touched"]:
        s = suggest(gray, box, side)
        if s is not None:
            fixes.append(f"{side} → {s}")
    if fixes:
        print("       建议坐标（已换算为相对比例）：" + "；".join(fixes))
        print("       换成上面的值重跑，通常一次到位。")
    else:
        print("       附近找不到空白位置：该边内容可能一直延伸到版面边缘，或这是满版图；改用整页裁切更稳妥。")
    print("       注意：若原图内容本身顶到该边（满版图/无边框图），触边属正常，看图确认即可。")
    return 4


if __name__ == "__main__":
    sys.exit(main())
