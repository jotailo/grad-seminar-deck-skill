#!/usr/bin/env python3
"""封存「当前这一份」组会 PPT：归档 deck.json 与 pptx，让下一份从空白开始。

规则（对应「这份完成了」之后的约定）：
  · 归档是**移动到子目录**，不删除任何东西；
  · 归档后当前目录不再有 deck.json，下次 append_unit.py 会新建一份空白 deck，
    分区从 01、页码从 1、汇总表从第一行重新开始；
  · 归档目录里留一份说明，写清这一份讲了几篇、共几页、哪天封的。

用法：
    python finish_deck.py 组会汇报.json
    python finish_deck.py 组会汇报.json --dry-run
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
import sys
from datetime import date
from pathlib import Path


def safe(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("_")[:40] or "组会汇报"


def main() -> int:
    ap = argparse.ArgumentParser(description="封存当前这一份组会 PPT")
    ap.add_argument("master", help="当前这一份的 deck.json")
    ap.add_argument("--out", help="对应的 pptx（默认与 deck.json 同名）")
    ap.add_argument("--dry-run", action="store_true", help="只打印要做什么，不动文件")
    args = ap.parse_args()

    master = Path(args.master).expanduser()
    if not master.exists():
        sys.exit(f"[error] 没有 {master}，当前目录没有在做的这一份")
    deck = json.loads(master.read_text(encoding="utf-8"))
    pptx = Path(args.out).expanduser() if args.out else master.with_suffix(".pptx")

    parts = deck.get("parts") or []
    units = [u for p in parts for u in (p.get("units") or [])]
    if not units:
        sys.exit("[error] 这一份里还没有任何单元，没必要封存")

    title = (deck.get("meta") or {}).get("title") or master.stem
    stamp = date.today().isoformat()
    folder = master.parent / "已结束" / f"{stamp}_{safe(title)}"
    n = 1
    base = folder
    while folder.exists():
        n += 1
        folder = Path(f"{base}_{n}")

    kinds = [p.get("kind") for p in parts]
    lines = [
        f"# {title}（已结束）",
        "",
        f"- 结束日期：{stamp}",
        f"- 汇报种类：{' / '.join(kinds)}",
        f"- 单元数：{len(units)}",
    ]
    for p in parts:
        for u in p.get("units") or []:
            lines.append(f"  - [{p.get('kind')}] {u.get('name', '')}"
                         f"（{len(u.get('pages') or [])} 页，加入于 {u.get('added', '未记')}）")
    if pptx.exists():
        lines.append(f"- 成品：{pptx.name}")

    print(f"[封存目标] {folder}")
    print("\n".join(lines))
    if args.dry_run:
        print("[dry-run] 没有移动任何文件")
        return 0

    folder.mkdir(parents=True, exist_ok=True)
    (folder / "_说明.md").write_text("\n".join(lines) + "\n", encoding="utf-8")

    shutil.move(str(master), str(folder / master.name))
    moved = [master.name]
    for extra in (pptx, master.with_suffix(".bak.json")):
        if extra.exists():
            shutil.move(str(extra), str(folder / extra.name))
            moved.append(extra.name)

    print(f"[完成] 已归档 {moved} → {folder}")
    print(f"[下一份] 当前目录已无 {master.name}；下次 append_unit.py 会自动新建空白 deck，"
          f"分区从 01、页码从 1 重新开始")
    return 0


if __name__ == "__main__":
    sys.exit(main())
