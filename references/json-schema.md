# JSON 输入结构

`scripts/deck_builder.py` 只认这一份结构。把文档解析结果写进 `slides` 即可，
版式不需要在 JSON 里描述——引擎已固化（见 `exam-deck-spec.md`）。

## 顶层

```json
{
  "title": "九年级道法期中考前练习卷（二）",
  "subtitle": "课堂讲解 · 试题与参考答案",
  "footer": "姓名：____________　　班级：____________",
  "params": {
    "theme": "quiet",
    "language": "zh",
    "answer_style": "inline",
    "animate": true
  },
  "slides": [ ... ]
}
```

`title` / `subtitle` / `footer` 会被封面页默认引用，页面里可覆盖。

## params

| 键 | 取值 | 默认 | 说明 |
|---|---|---|---|
| `theme` | `quiet` / `exam` / `brand` | `quiet` | 纯白+浅青答案块 / 纯白+浅灰答案块 / 纯白+浅蓝答案块 |
| `theme_overrides` | `{"answer_fill": "#E9FFFF", "accent": "#E10000", ...}` | 无 | 覆盖主题里的任一颜色键：`bg` `answer_fill` `accent` `ink` `muted` `grid` `info` |
| `language` | `zh` / `en` | `zh` | 影响「参考答案」等内置文案 |
| `answer_style` | `inline` / `right_big` / `none` | `inline` | 答案嵌题干括号并标红 / 右侧 96 号大字 / 不显示答案 |
| `animate` | true / false | true | 答案（字母、色块）是否添加飞入动画 |
| `bullets_per_page` | 整数 | 6 | 要点页每页条数 |

## slides 类型

### `cover` 封面
```json
{"type": "cover"}
{"type": "cover", "title": "自定义标题", "subtitle": "副标题", "footer": "姓名/班级", "title_pt": 46}
```

### `section` 章节页
```json
{"type": "section", "title": "第二部分　非选择题", "subtitle": "21—23 题"}
```

### `answer_table` 选择题答案表
```json
{"type": "answer_table", "section": "选择题", "answers": ["D","D","C"], "per_row": 10, "start": 1}
```
`answers` 按题号顺序给出；每 `per_row` 题一行，自动折成「题号行 / 答案行」两组。

### `choice` 选择题（一题一页）
```json
{
  "type": "choice",
  "no": 1,
  "stem": "习近平总书记指出：“时代是思想之母……”这与下列古语蕴含的哲理一致的是（　）",
  "items": ["①一寸光阴一寸金", "②横看成岭侧成峰"],
  "options": ["A．①②", "B．①③", "C．②④", "D．③④"],
  "answer": "D"
}
```
- `stem` 里必须留 `（　）`（全角括号 + 一个全角空格）作为答案落点，引擎据此定位并填入红色字母。
- 没有 `items` 就省略该键。
- 字号从 28 号起自动下探到放得下为止（最低 20 号）。

### `reading` 材料题 / 解题页
```json
{
  "type": "reading",
  "label": "21（1）",
  "blocks": [
    {"kind": "material", "text": ["“江山就是人民……”某校九（2）班开展项目式学习。", "【任务一  寻身边之变】"]},
    {"kind": "question", "text": ["结合上述材料，谈谈这些“民生温度”是怎样贯彻党的群众路线的。（4分）"]},
    {"kind": "answer", "text": ["原理：党的群众路线是……", "①……", "②……"]}
  ]
}
```
- `label` 会自动并入首段开头（有材料并入材料首段，无材料并入问题），形如「21（1）．……」。
- 放不下时引擎自动分页：材料逐行续页 → 问题跟随 → 答案放不下时独占一页并**重印问题**。
- `blocks` 可只用 `answer`，用于「只有答案」的独立页；也可只用 `material`+`question`。

### `bullets` 要点页（普通文档 / 讲义 / 报告）
```json
{"type": "bullets", "title": "一、坚持党的领导", "items": ["要点一", "要点二"], "per_page": 6, "columns": 1}
```
超出 `per_page` 自动续页，标题加「（续）」。

### `blank` 空白页
```json
{"type": "blank"}
```

## 一条最小可用示例

```json
{
  "title": "第三课　认清基本国情",
  "subtitle": "课堂讲解",
  "params": {"theme": "quiet", "language": "zh"},
  "slides": [
    {"type": "cover"},
    {"type": "bullets", "title": "一、我国社会的主要矛盾", "items": ["人民日益增长的美好生活需要", "和不平衡不充分的发展之间的矛盾"]},
    {"type": "reading", "label": "例1", "blocks": [
      {"kind": "question", "text": ["简述我国社会主要矛盾的转化。（4分）"]},
      {"kind": "answer", "text": ["①……", "②……"]}
    ]}
  ]
}
```

渲染：
```bash
python scripts/deck_builder.py spec.json out.pptx
```
