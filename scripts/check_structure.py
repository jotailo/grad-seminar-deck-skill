#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""check_structure.py — 在生成 PPT 之前，先检查 deck.json 的内容结构是否齐、是否守预算。

用法：
    python check_structure.py deck.json
    python check_structure.py deck.json --json report.json

它查的是「内容有没有讲全」，与版式无关（版式由 check_deck.py 查）。逐项对应：
  1. 三种汇报各自的必备内容点是否都在（用每页的 roles 字段登记）
  2. 每个单元是否 2–4 页
  3. 文献汇总表是否存在、列是否齐、是否每篇论文一行
  4. 证据（原文图表）是否至少有一张，且带图注 + 出处
  5. 期刊等级是否给了可核查口径

退出码：0 通过；4 有 FAIL；3 文件不可读。
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

KIND_LABEL = {"paper": "论文阅读汇报", "experiment": "实验报告", "project": "项目进展汇报"}

# 每种汇报的必备内容点（roles）。左边是键，右边是给人看的名字。
REQUIRED = {
    "paper": [
        ("name", "文献标题"),
        ("source", "文献出处（期刊/会议全称、等级）"),
        ("thesis", "作者观点与理论"),
        ("argument", "推理与论证过程"),
        ("problem", "所解决的问题"),
        ("contribution", "文章贡献"),
        ("experiment", "实验设计"),
        ("mapping", "实验 ↔ 观点逐项对应"),
        ("insight", "个人观点 / 对自身的启发"),
    ],
    "experiment": [
        ("name", "实验名称"),
        ("purpose", "实验目的"),
        ("setup", "实验设置"),
        ("finding", "核心发现"),
        ("limit", "局限"),
        ("next", "后续可做的事项"),
        ("conclusion", "结论"),
    ],
    "project": [
        ("name", "项目名称"),
        ("progress", "当前进度"),
        ("problem", "当前遇到的问题"),
        ("plan", "下一步展望与规划"),
    ],
}

# 文献汇总表列检查：每组里命中任意一个词即算有
SUMMARY_COLS = {
    "文献标题": ("标题", "题目"),
    "出处": ("出处", "期刊", "会议", "venue"),
    "发布时间": ("时间", "年份", "发表", "日期"),
    "一句话核心内容": ("核心", "内容", "一句话"),
}
GRADE_HINTS = ("中科院", "CCF", "JCR", "一区", "二区", "三区", "四区", "Q1", "Q2",
               "A类", "B类", "顶会", "SCI", "EI", "未核实", "分区", "预印本")


def pages_of(unit):
    return unit.get("pages") or []


def roles_of(page):
    r = page.get("roles") or page.get("role")
    if not r:
        return []
    return [r] if isinstance(r, str) else list(r)


def all_text(page) -> str:
    """把一页里所有会显示的文字抓出来，用于关键词类检查。"""
    buf = [str(page.get("title", ""))]
    for b in page.get("bullets") or []:
        buf.append(b if isinstance(b, str) else str(b.get("text", "")))
    for e in page.get("entries") or []:
        buf.append(f"{e.get('label', '')}{e.get('value', '')}")
    for col in page.get("columns") or []:
        buf.append(str(col.get("heading", "")))
        for b in col.get("bullets") or []:
            buf.append(b if isinstance(b, str) else str(b.get("text", "")))
    for row in page.get("rows") or []:
        buf.extend(str(c) for c in row)
    buf.extend(str(c) for c in page.get("columns", []) if isinstance(c, str))
    if page.get("takeaway"):
        buf.append(str(page["takeaway"]))
    return " ".join(buf)


def check(spec: dict):
    fails, warns, oks = [], [], []
    parts = spec.get("parts") or []
    if not parts:
        return ["parts 为空：没有任何汇报内容"], warns, oks

    seen_kinds = []
    for pi, part in enumerate(parts, 1):
        kind = part.get("kind", "")
        label = KIND_LABEL.get(kind, f"第 {pi} 部分（kind={kind}）")
        if kind not in KIND_LABEL:
            fails.append(f"{label}：kind 必须是 paper / experiment / project 之一")
            continue
        seen_kinds.append(kind)
        units = part.get("units") or []
        part_pages = part.get("pages") or []
        if not units:
            fails.append(f"{label}：没有 units（每个论文/实验/项目都应是一个 unit）")

        # —— 文献汇总表（仅论文类、且不止一篇时必须有）
        if kind == "paper" and (len(units) > 1 or any(p.get("require_summary") for p in part_pages)):
            tbls = [p for p in part_pages if "summary_table" in roles_of(p)]
            if not tbls:
                # 没手写表：只要每篇都写了 brief，生成时会自动出表
                missing = [u.get("name", "")[:16] or f"第{i}篇"
                           for i, u in enumerate(units, 1)
                           if not (u.get("brief") or {}).get("source")]
                if missing:
                    fails.append(f"{label}：缺文献汇总表——要么放一个 part 级 table 页"
                                 f"（roles 写 [\"summary_table\"]），要么给每篇补 brief"
                                 f"（这几篇还没写 brief.source：{missing}）")
                else:
                    oks.append(f"{label}：文献汇总表由 {len(units)} 篇的 brief 自动生成")
            else:
                t = tbls[0]
                cols = [str(c) for c in (t.get("columns") or [])]
                joined = " ".join(cols)
                for need, hints in SUMMARY_COLS.items():
                    if not any(h.lower() in joined.lower() for h in hints):
                        fails.append(f"{label}：文献汇总表缺「{need}」列（现有：{cols}）")
                rows = t.get("rows") or []
                if len(rows) != len(units):
                    fails.append(f"{label}：汇总表 {len(rows)} 行，但论文有 {len(units)} 篇，必须一一对应")
                if rows and len(rows[0]) != len(cols):
                    fails.append(f"{label}：汇总表每行列数（{len(rows[0])}）与表头（{len(cols)}）不一致")
                oks.append(f"{label}：文献汇总表 {len(rows)} 行 × {len(cols)} 列")

        # —— 追加模式最容易犯的两个错：同一篇被追加两次、brief 只写一半
        names = [str(u.get("name", "")).strip() for u in units]
        dup = sorted({n for n in names if n and names.count(n) > 1})
        if dup:
            # 只提醒不判失败：同名可能有正当理由（重读、同名子任务），
            # 真正的重复拦截放在 append_unit.py（可用 --allow-dup 放行）。
            warns.append(f"{label}：单元名重复，确认不是同一篇被追加了两次：{dup}")
        for unit in units:
            b = unit.get("brief") or {}
            if not b:
                continue
            lack = [k for k in ("source", "date", "gist") if not b.get(k)]
            if lack:
                warns.append(f"{label} · 「{str(unit.get('name', ''))[:16]}」brief 缺 {lack}，"
                             f"汇总表对应格子会空着")

        # —— 逐单元检查
        for ui, unit in enumerate(units, 1):
            name = unit.get("name", "") or (f"第{ui}个单元")
            short = name if len(name) <= 24 else name[:24] + "…"
            pages = pages_of(unit)
            n = len(pages)

            if not (2 <= n <= 4):
                fails.append(f"{label} · 「{short}」内容页 {n} 页，要求 2–4 页")
            else:
                oks.append(f"{label} · 「{short}」{n} 页")

            covered = set()
            for p in pages:
                covered.update(roles_of(p))
                if not str(p.get("title", "")).strip():
                    fails.append(f"{label} · 「{short}」有页面没写 title")
                if len(str(p.get("title", ""))) > 46:
                    warns.append(f"{label} · 「{short}」标题过长（{len(p['title'])} 字），会换行挤占正文：{p['title'][:20]}…")
                if p.get("type") in ("image", "image_only"):
                    if not str(p.get("caption", "")).strip():
                        fails.append(f"{label} · 「{short}」插页缺 caption（图/表编号 + 说明）")
                    if not str(p.get("source", "")).strip():
                        warns.append(f"{label} · 「{short}」插页缺 source（原文第几页哪张图表），被追问时翻不回去")
            slim = bool(unit.get("slim")) and n <= 2
            for key, human in REQUIRED[kind]:
                if key in covered:
                    continue
                if slim:
                    warns.append(f"{label} · 「{short}」精简版缺内容点：{human}"
                                 f"（确认是有意省略，否则补上）")
                else:
                    fails.append(f"{label} · 「{short}」缺内容点：{human}"
                                 f"（在对应页的 roles 里登记 \"{key}\"）")

            # 证据：论文必须有原文图表；实验必须有自己的结果图；项目至少提示一次
            evid = [p for p in pages if p.get("evidence") or p.get("type") in ("image", "image_only")]
            if not evid:
                reason = unit.get("no_evidence_reason")
                if reason:
                    warns.append(f"{label} · 「{short}」没有插图，理由：{reason}")
                elif kind in ("paper", "experiment"):
                    fails.append(f"{label} · 「{short}」没有任何图表作证据"
                                 f"（论文截原文图，实验放结果图；确实无法出图就写 no_evidence_reason）")
                else:
                    warns.append(f"{label} · 「{short}」没有图：进度汇报建议附一张进度图或结果图，比纯文字有说服力")
            else:
                oks.append(f"{label} · 「{short}」证据图 {len(evid)} 页")

            # 论文类：出处页必须给期刊等级口径
            if kind == "paper":
                src_pages = [p for p in pages if "source" in roles_of(p)]
                if src_pages:
                    txt = all_text(src_pages[0])
                    if not any(h.lower() in txt.lower() for h in GRADE_HINTS):
                        warns.append(f"{label} · 「{short}」出处页没写期刊/会议等级（CCF/中科院分区/JCR，"
                                     f"或「未核实」），组会上最容易被追问这一条")
            # 实验/论文：对应关系必须是表
            if kind in ("paper", "experiment"):
                mp = [p for p in pages if "mapping" in roles_of(p)]
                if mp and mp[0].get("type") != "table":
                    warns.append(f"{label} · 「{short}」mapping 建议用 table 版式（实验｜印证的观点｜证据位置），"
                                 f"当前是 {mp[0].get('type')}")

    if "paper" in seen_kinds and "experiment" in seen_kinds:
        pass  # 混排允许，分区由 build_deck.py 自动插入
    return fails, warns, oks


def main() -> int:
    ap = argparse.ArgumentParser(description="组会 PPT 内容结构检查")
    ap.add_argument("spec")
    ap.add_argument("--json", help="把报告写成 JSON")
    ap.add_argument("--quiet-ok", action="store_true")
    args = ap.parse_args()

    path = Path(args.spec)
    if not path.exists():
        print(f"[error] 找不到文件：{path}", file=sys.stderr)
        return 3
    spec = json.loads(path.read_text(encoding="utf-8"))
    fails, warns, oks = check(spec)

    print(f"=== 结构检查：{path.name} ===")
    if not args.quiet_ok:
        for o in oks:
            print(f"[OK] {o}")
    for w in warns:
        print(f"[WARN] {w}")
    for f in fails:
        print(f"[FAIL] {f}")
    print(f"结论：{'通过' if not fails else f'{len(fails)} 项需修复'}"
          f"{'' if not warns else f'，{len(warns)} 项提醒'}")

    if args.json:
        Path(args.json).write_text(json.dumps(
            {"file": str(path), "fails": fails, "warns": warns, "oks": oks},
            ensure_ascii=False, indent=2), encoding="utf-8")
    return 0 if not fails else 4


if __name__ == "__main__":
    sys.exit(main())
