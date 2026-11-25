# doc-to-ppt

WorkBuddy / Claude Skill：把文档（Word、PDF、Markdown、纯文本、截图文字、试卷、讲义、笔记）自动整理成 **16:9 中文讲解 PPT**，并套用一套固化的版式规范——字号行距、答案标红并嵌入题干括号、答案浅青色块、飞入动画、全篇统一纵向基线。

版式规则全部固化在 `scripts/deck_builder.py` 里，调用方只需要写一份内容 JSON，不用每次重新描述排版。

## 安装

克隆到 WorkBuddy 的 skills 目录（用户级，所有项目可用）：

```bash
# macOS / Linux
git clone https://github.com/TuTu8383/doc-to-ppt.git ~/.workbuddy/skills/doc-to-ppt

# Windows (PowerShell)
git clone https://github.com/TuTu8383/doc-to-ppt.git "$env:USERPROFILE\.workbuddy\skills\doc-to-ppt"
```

只想更新已有安装，进目录 `git pull` 即可。

也可以放到项目级目录 `<项目>/.workbuddy/skills/doc-to-ppt/`，只对该项目生效。

安装后重启会话，在对话里发文档并说「做成 PPT」「转成课件」「按上次那个格式再来一份」即可触发。

## 依赖

```bash
pip install python-pptx pillow pymupdf lxml
```

`preview.py` 做自检时需要本机装有 LibreOffice 或 WPS 之一（脚本会自动查找）。

字体：华文中宋 `STZHONGS.TTF`、楷体 `simkai.ttf`、黑体 `simhei.ttf`。换机器前先确认存在，否则行宽测量会偏，自动断行不准（引擎有粗略兜底）。

## 目录结构

| 路径 | 用途 |
| --- | --- |
| `SKILL.md` | 技能入口：触发条件、处理流程、版式铁律、可调参数 |
| `scripts/deck_builder.py` | 版式引擎。JSON → PPTX，所有排版规则在此 |
| `scripts/preview.py` | 转 PDF、逐页量测、出预览图，生成自检报告 |
| `references/json-schema.md` | spec 结构与各页面类型示例 |
| `references/exam-deck-spec.md` | 完整字号/行距/锚点/断行/分页规格与踩坑记录 |
| `references/qa-checklist.md` | 交付前硬指标与视觉抽查项 |

## 快速用法

1. 按 `references/json-schema.md` 写一份 `spec.json`（题干里用 `（　）` 留出答案落点）。
2. 渲染：`python scripts/deck_builder.py spec.json out.pptx`
3. 自检：`python scripts/preview.py out.pptx`，对照 `references/qa-checklist.md` 核对硬指标。
4. 需要改版式，改 JSON 重新渲染，不要手改 pptx 里的形状。

## 支持的内容形态

- **试卷**：选择题（题干 + 小项 + ABCD + 答案）、材料题（材料 + 问题 + 参考答案）、答案解析
- **讲义 / 笔记**：章节标题 + 条目要点
- **报告 / 汇报**：多级标题 + 段落 + 数据结论

页面类型：`cover` 封面、`answer_table` 答案表、`choice` 选择题、`reading` 材料题、`bullets` 要点、`section` 章节页。

## 常用参数

写在 spec 的 `params` 里，不传用默认值。

| 参数 | 取值 | 默认 | 效果 |
| --- | --- | --- | --- |
| `theme` | `quiet` / `exam` / `brand` | `quiet` | 浅青 / 浅灰 / 浅蓝答案块 |
| `language` | `zh` / `en` | `zh` | 内置文案语言 |
| `answer_style` | `inline` / `right_big` / `none` | `inline` | 答案嵌括号 / 右侧大字 / 不显示 |
| `animate` | true / false | true | 答案飞入 |
| `bullets_per_page` | 1—8 | 6 | 要点页密度 |

## 许可

自用技能，未附许可协议。如需转载或商用请先联系作者。
