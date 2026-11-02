#!/usr/bin/env python3
"""把一篇论文 / 一组实验 / 一个项目追加进「当前这一份」组会 PPT，然后重新生成。

核心约定：
  · 只维护一个 deck.json，它代表「当前这一份 PPT」；
  · 每追加一个单元，整份文件的分区号、页码、目录、汇总表全部按当下内容重算；
  · 这一份结束后用 finish_deck.py 封存，下一份从空白开始，编号与页码重头来。

用法：
    python append_unit.py 组会汇报.json 新论文.json
    python append_unit.py 组会汇报.json 新论文.json --session 2026-11-02
    python append_unit.py 组会汇报.json 新论文.json -o 组会汇报.pptx --no-build

片段（新论文.json）两种写法都收：
    {"kind": "paper", "unit": {"name": "...", "brief": {...}, "pages": [...]}}
    {"kind": "paper", "name": "...", "brief": {...}, "pages": [...]}
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from build_deck import KIND_LABEL, KIND_ORDER  # noqa: E402  （只取常量，避免两处写法漂移）

DUP_EXIT = 6
BAD_MASTER = 7


def skeleton(title: str, subtitle: str = "") -> dict:
    return {
        "meta": {"title": title, "footer": ""},
        "cover": {"title": title, "subtitle": subtitle or title},
        "agenda": True,
        "parts": [],
        "end": {"title": "谢谢，请老师和同学批评指正"},
    }


def load_fragment(path: Path, kind_arg: str | None) -> tuple[str, dict, dict]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        sys.exit(f"[error] 找不到片段文件：{path}")
    except json.JSONDecodeError as e:
        sys.exit(f"[error] 片段不是合法 JSON：{e}")

    kind = kind_arg or raw.get("kind")
    if kind not in KIND_LABEL:
        sys.exit(f"[error] 片段必须写明 kind（paper / experiment / project），当前是 {kind!r}")

    unit = raw.get("unit") if isinstance(raw.get("unit"), dict) else None
    if unit is None:
        unit = {k: v for k, v in raw.items() if k not in ("kind", "part_title", "part_subtitle")}

    if not str(unit.get("name", "")).strip():
        sys.exit("[error] 片段缺 name（文献标题 / 实验名称 / 项目名称）")
    pages = unit.get("pages") or []
    if not pages:
        sys.exit("[error] 片段缺 pages：一个单元至少 2 页")
    if not (2 <= len(pages) <= 4):
        sys.exit(f"[error] 这个单元写了 {len(pages)} 页，要求 2–4 页；"
                 f"页数不合规先别追加，改好再来")
    return kind, unit, raw


def ensure_part(deck: dict, kind: str, frag: dict) -> dict:
    """取该种类的分区；没有就按规范顺序新建一个（保证分区号与目录一致）。"""
    for part in deck.setdefault("parts", []):
        if part.get("kind") == kind:
            return part
    order = {k: i for i, k in enumerate(KIND_ORDER)}
    rank = order.get(kind, len(order))
    pos = len(deck["parts"])
    for i, part in enumerate(deck["parts"]):
        if order.get(part.get("kind"), len(order)) > rank:
            pos = i
            break
    part = {
        "kind": kind,
        "title": frag.get("part_title") or KIND_LABEL[kind],
        "subtitle": frag.get("part_subtitle", ""),
        "units": [],
    }
    deck["parts"].insert(pos, part)
    print(f"[新建分区] {part['title']}（第 {pos + 1} 部分）")
    return part


def run(script: str, *args: str) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(HERE / script), *args],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")


def rebase_images(unit: dict, frag_dir: Path, master_dir: Path) -> list[str]:
    """片段里的相对图片路径若在 master 旁边找不到、但在片段旁边找得到，就改成绝对路径。

    生成时的图片路径是相对 deck.json 解析的；片段常放在别处（甚至别的盘），
    不修这一步就会莫名报「图片不存在」。
    """
    fixed = []
    for page in unit.get("pages") or []:
        img = page.get("image")
        if not img or Path(img).is_absolute():
            continue
        if (master_dir / img).exists() or not (frag_dir / img).exists():
            continue
        page["image"] = str((frag_dir / img).resolve())
        fixed.append(img)
    return fixed


def main() -> int:
    ap = argparse.ArgumentParser(description="把单元追加进当前这一份组会 PPT")
    ap.add_argument("master", help="当前这一份的 deck.json（不存在则新建）")
    ap.add_argument("fragment", help="要追加的那一篇/那一组/那一个项目的 JSON")
    ap.add_argument("--kind", help="片段里没写 kind 时用它指定")
    ap.add_argument("--session", help="给这个单元打上本次汇报的标签，如 2026-11-02")
    ap.add_argument("-o", "--out", help="输出 pptx（默认与 deck.json 同名）")
    ap.add_argument("--title", default="组会汇报", help="新建 deck 时的封面标题")
    ap.add_argument("--allow-dup", action="store_true", help="允许同名单元重复追加")
    ap.add_argument("--no-build", action="store_true", help="只改 JSON，不生成 pptx")
    args = ap.parse_args()

    master = Path(args.master).expanduser()
    frag_path = Path(args.fragment).expanduser()
    kind, unit, raw = load_fragment(frag_path, args.kind)

    if master.exists():
        try:
            deck = json.loads(master.read_text(encoding="utf-8"))
        except json.JSONDecodeError as e:
            print(f"[error] 现有 deck 解析失败：{e}", file=sys.stderr)
            return BAD_MASTER
        print(f"[当前这一份] {master}：已有 {sum(len(p.get('units') or []) for p in deck.get('parts') or [])} 个单元")
    else:
        deck = skeleton(args.title)
        print(f"[新一份] {master} 不存在，按空白新建（编号从 01、页码从 1 开始）")

    part = ensure_part(deck, kind, raw)

    if not args.allow_dup and any(u.get("name") == unit["name"] for u in part.get("units") or []):
        print(f"[error] 「{unit['name'][:24]}」已经在 {KIND_LABEL[kind]} 里了，"
              f"看起来是重复追加。确认要再放一份就加 --allow-dup", file=sys.stderr)
        return DUP_EXIT

    if args.session:
        unit["session"] = args.session
    unit.setdefault("added", date.today().isoformat())
    fixed = rebase_images(unit, frag_path.parent, master.parent)
    if fixed:
        print(f"[图片] {len(fixed)} 张图按片段所在目录定位：{fixed[:3]}"
              f"{'…' if len(fixed) > 3 else ''}")
    part.setdefault("units", []).append(unit)

    # 先写临时文件跑检查，通过才覆盖正式文件——检查不过时原文件保持不动
    master.parent.mkdir(parents=True, exist_ok=True)
    tmp = master.with_suffix(".appending.json")
    tmp.write_text(json.dumps(deck, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    chk = run("check_structure.py", str(tmp))
    body = (chk.stdout or "") + (chk.stderr or "")
    if chk.returncode != 0:
        tmp.unlink(missing_ok=True)
        print(body.rstrip(), file=sys.stderr)
        print(f"[中止] 内容检查没过，{master.name} 没有被改动。改好片段再追加。", file=sys.stderr)
        return chk.returncode

    if master.exists():
        backup = master.with_suffix(".bak.json")
        shutil.copyfile(master, backup)
        print(f"[备份] 上一版 → {backup.name}")
    os.replace(tmp, master)

    n_units = sum(len(p.get("units") or []) for p in deck["parts"])
    print(f"[追加] {KIND_LABEL[kind]} · 「{unit['name'][:30]}」{len(unit.get('pages') or [])} 页　"
          f"（这一份现在共 {n_units} 个单元）")
    for p in deck["parts"]:
        names = [u.get("name", "")[:14] for u in p.get("units") or []]
        print(f"  {KIND_LABEL.get(p.get('kind'), p.get('title', ''))}：{len(names)} 个 → {names}")

    if args.no_build:
        print(f"[下一步] python build_deck.py {master} -o <文件名>.pptx")
        return 0

    out = Path(args.out) if args.out else master.with_suffix(".pptx")
    b = run("build_deck.py", str(master), "-o", str(out))
    print((b.stdout or "").rstrip())
    if b.returncode != 0:
        print((b.stderr or "").rstrip(), file=sys.stderr)
        print("[中止] 生成失败，但 deck.json 已保存；修好后再 build 一次即可", file=sys.stderr)
        return b.returncode

    d = run("check_deck.py", str(out))
    dout = (d.stdout or "").rstrip().splitlines()
    bad = [ln for ln in dout if "FAIL" in ln]
    if bad:
        print("\n".join(bad[:12]))
        print(f"[注意] 版面体检有 {len(bad)} 条不通过，处理完重新 build")
    else:
        print("[体检] 版面通过（无 FAIL）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
