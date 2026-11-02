# deck.json 规格

`build_deck.py` 只认这一份结构。字段名不要改，改了就报「未知版式」。

## 三级结构

```
deck.json
├── meta / cover / agenda / end      ← 全局
└── parts[]                          ← 种类（论文 / 实验 / 项目），一份 PPT 可有 1–3 个
    ├── pages[]                      ← 种类级页面（如「文献汇总表」）
    └── units[]                      ← 单元（一篇论文 / 一组实验 / 一个项目）
        └── pages[]                  ← 单元内容页，必须 2–4 页
```

生成时自动插入：封面 → 目录（`agenda: true`）→ 每个种类的分区页 → 种类级页面 → 各单元内容页 → 结束页。页码与页脚自动加，不用手写。

## 全局字段

```json
{
  "meta": {
    "title": "智能体方向组会汇报",
    "footer": "智能体方向 · 第 12 周组会",
    "cover_meta": ["汇报人：刘同学　　2026-10-31"]
  },
  "cover": {"kicker": "研究生组会汇报", "title": "…", "subtitle": "…"},
  "agenda": true,
  "end": {"title": "谢谢，请老师和同学批评指正", "subtitle": "…"}
}
```
- `meta.footer`：每页页脚左侧文字（加 `source` 时自动拼在后面）。
- `agenda`：`true` 时按 `parts` 自动生成目录，不用手写条目。

## part（种类）

| 字段 | 必填 | 说明 |
|---|---|---|
| `kind` | ✔ | `"paper"` / `"experiment"` / `"project"`，决定分区页标题与必备内容点 |
| `title` | | 分区页标题，缺省用种类名 |
| `subtitle` | | 分区页副标题（一句定位，如「两篇编排范式代表作」） |
| `items` | | 分区页的单元清单，缺省自动用单元名生成 |
| `pages` | | 种类级页面（论文类的汇总表放这里） |
| `units` | ✔ | 单元数组 |

## unit（单元）

| 字段 | 必填 | 说明 |
|---|---|---|
| `name` | ✔ | 论文名 / 实验名 / 项目名，用于分区页清单与报错定位 |
| `subtitle` | | 单元分隔页副标题（仅在 `divider: true` 时用） |
| `divider` | | `true` 时给这个单元单独加一页大编号分隔页（不计入 2–4 页预算，默认关） |
| `slim` | | `true` 时允许 2 页精简版，缺失内容点降级为提醒而非失败 |
| `no_evidence_reason` | | 确实无法出图时写原因，避免被判定成漏图 |
| `brief` | | 文献汇总表那一行，见下方；论文类多篇时写它就不用手写汇总表 |
| `session` | | 场次标签（如 `"2026-11-02"`），`build_deck.py --session` 按它挑单元 |
| `added` | | 加入日期，`append_unit.py` 自动写，只在归档清单里用 |
| `pages` | ✔ | 内容页，**2–4 页**，越界直接报错退出 |

`brief` 子字段（三个都建议写，缺了会在检查里提醒）：

| 子字段 | 对应汇总表列 |
|---|---|
| `source` | 出处（`ICLR 2023（CCF-A）`） |
| `date` | 发布时间（`2022-10`） |
| `gist` | 一句话核心内容（≤30 字，别写两句） |

标题目默认取 `name`，要不一样就写 `brief.title`。

**汇总表是自动的**：论文类有 ≥2 个单元、且每篇都写了 `brief.source` 时，`build_deck.py` 在分区页之后自动插一页汇总表，行数=篇数。
若你在 `part.pages` 里手写了 `roles: ["summary_table"]` 的表，则用手写版、不再自动生成（不会出现两张）。

## page（页面）

所有版式通用字段：

| 字段 | 说明 |
|---|---|
| `type` | 必填，见下表；缺省 `bullets` |
| `roles` | 内容点标签，供 `check_structure.py` 核对（见文末），可写多个 |
| `title` | 页标题，判断句，建议 ≤30 字（>46 字会换行挤占正文并告警） |
| `kicker` | 标题上方小标签；单元内页面不写会自动生成 `论文阅读汇报 · 2/3` |
| `sub` | 标题下的一句补充说明（可选） |
| `takeaway` | 页脚上方结论条，≤45 字，写「这一页到底说明什么」 |
| `source` | 出处，渲染到页脚（`原文 p.5 Table 1` / `实验记录 E-2310`） |
| `notes` | 演讲备注，写进备注栏，不显示在页面 |
| `evidence` | `true` 标记这页是证据页（插图/表格），供检查统计 |

### 版式一览

| `type` | 专有字段 | 用途 |
|---|---|---|
| `cover` / `end` / `agenda` | 自动生成 | 不用手写 |
| `bullets` | `bullets: [str \| {text, level, bold}]` | 观点、启发、问题、计划 |
| `image` / `image_only` | `image`(必填), `caption`, `bullets`, `image_ratio`(默认 0.60), `image_side`(`left`/`right`), `frame` | 原文图表作证据 |
| `table` | `columns[]`, `rows[[...]]`, `col_widths[]`(和为 1), `width_ratio`, `highlight_col`, `caption` | 汇总表、实验↔观点对应、指标对照 |
| `cards` | `entries: [{label, value, strong, bold, wide}]`, `cols`(1–3) | 标题/出处/目的/设置等并列事实 |
| `columns` | `columns: [{heading, bullets[]}]`，2–4 栏 | 进度/问题/下一步、局限/后续/结论 |
| `statement` | `text`, `attribution` | 一句必须让全场记住的话 |

### 写法示例

```json
{
  "type": "image",
  "roles": ["thesis", "argument"],
  "title": "作者主张：思考与行动必须交替发生",
  "bullets": ["总观点：__只推理不行动__会让模型自圆其说", {"text": "细节", "level": 1}],
  "image": "figs/react_fig1.png",
  "caption": "图 1　ReAct 的 Thought–Act–Observation 循环（原文 p.2 Figure 1）",
  "source": "原文 p.2 Figure 1",
  "evidence": true,
  "takeaway": "它不是在提示里加一句「请先思考」，而是把观测变成推理的输入。"
}
```

文本里可用三种轻量标记：`**加粗**`、`__强调色加粗__`（用于关键词）。
图片路径写相对 `deck.json` 的路径（推荐 `figs/xxx.png`）或绝对路径；不存在会在生成前报错。

### 表格的硬规则

- `col_widths` 之和为 1，建议列宽差不要超过 2.5 倍。
- 单元格文字 ≤ 28 字；超过会自动降字号，降到 12pt 仍塞不下会 `[warn] 表格偏密`，此时**精简文字或拆页**，不要加宽表格。
- 不要在一格里写「2022-10 / 2023-05」这类双值——列窄时会断在括号或斜杠上，很难看。

## roles：内容点标签

`check_structure.py` 按 `kind` 核对下面这些键是否都在（同一页可登记多个）：

| kind | 必备 roles |
|---|---|
| `paper` | `name` `source` `thesis` `argument` `problem` `contribution` `experiment` `mapping` `insight` |
| `experiment` | `name` `purpose` `setup` `finding` `limit` `next` `conclusion` |
| `project` | `name` `progress` `problem` `plan` |

另有两个种类级标签：
- `summary_table`：文献汇总表（放在 `part.pages`；不手写则由 `brief` 自动生成）。
- `mapping`：必须出现在 `table` 版式上，用来放「实验 ↔ 观点 ↔ 证据位置」。

## 追加片段（fragment）规格

`append_unit.py` 吃的是**一个单元**，不是整份 deck。两种写法等价：

```json
{ "kind": "paper", "unit": { "name": "...", "brief": {...}, "pages": [...] } }
{ "kind": "paper", "name": "...", "brief": {...}, "pages": [...] }
```

- `kind` 必填（也可用 `--kind` 传），决定追加进哪个种类分区；没有该分区就按 `paper → experiment → project` 的顺序新建。
- `part_title` / `part_subtitle` 可选，只在新建分区时用。
- `pages` 必须 2–4 页，否则**在写文件前**就被拒（不会留下半截状态）。
- 图片相对路径优先按 master 所在目录找；找不到而片段旁边有，就自动改写成绝对路径。
- 追加时自动补 `added`（今天）；`--session` 会写成 `session`。

## 会报错 / 会提示的情况

| 现象 | 级别 | 处理 |
|---|---|---|
| 单元内容页不是 2–4 页 | 错误（退出码 5） | 合并/拆分页面；临时放过 `--skip-budget-check` |
| `image` 指向的文件不存在 | 错误（退出码 3） | 先按 `figure-capture.md` 裁图 |
| `type` 未知 | 错误（退出码 2） | 只能用上表列出的版式 |
| 追加同名单元 | 错误（退出码 6） | 改标题，或 `--allow-dup` |
| 追加后内容检查不过 | 中止（检查自己的退出码） | `deck.json` 保持原样，改片段再追加 |
| `--session` 没有任何单元匹配 | 错误（退出码 4） | 先给单元打标签，或去掉该参数 |
| 正文 >260 字、结论条 >78 字 | 提示 | 拆页或删句 |
| 末行只剩一两个字 | 提示 `[hint]` | 改写措辞，避免孤字 |
| 内容占版面 <45% | 提示 | 补细节或与相邻页合并 |
| `brief` 缺 source/date/gist | 提示 | 汇总表对应格子会空着 |
| 图片有效 dpi <110 | 体检 FAIL | 用 300/400dpi 重渲染页面再裁 |
