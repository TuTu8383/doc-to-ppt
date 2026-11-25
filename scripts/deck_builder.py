# -*- coding: utf-8 -*-
"""doc-to-ppt 版式引擎

把「结构化 JSON」渲染成课堂讲解 / 汇报用 PPTX。所有排版规则（字号、行距、
避头尾断行、统一纵向基线、答案嵌括号并标红、答案色块、飞入动画、自动分页）
都固化在本文件中，调用方只需描述内容，不必重复描述版式。

用法：
    python deck_builder.py spec.json out.pptx

JSON 结构见 references/json-schema.md。
"""
from __future__ import annotations

import json
import os
import sys

from PIL import ImageFont
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- 基本几何
SW, SH = 13.333, 7.5                 # 16:9
ML = 0.55                            # 左右页边距（英寸）
CW = SW - 2 * ML
TOP = 0.42                           # 选择题/正文统一顶端基准
BOT = 0.28
AVAIL_PT = (SH - TOP - BOT) * 72

# 应用题文本框顶边：补偿「固定行距」与「1.5 倍行距」的行框差异，
# 使两类页面的首行字顶落在同一条水平线上（实测 42pt）。
APP_TOP = 0.55
APP_AVAIL_PT = (SH - APP_TOP - BOT) * 72

# ---------------------------------------------------------------- 字号规格
MAT_PT, MAT_SP = 32, 38              # 材料：楷体 32 / 固定 38
Q_PT, Q_SP = 28, 39                  # 问题：华文中宋加粗 28 / 固定 39
ANS_PT, ANS_SP = 24, 38              # 答案：黑体加粗 24 / 固定 38
ANS_PAD = 8                          # 色块上下内边距（磅）
ANS_MX = 0.20                        # 色块左右内边距（英寸）
Q_GAP_BEFORE_ANS = 10 / 72.0         # 问题 → 答案色块
CHOICE_LF = 1.8                      # 实测：1.5 倍行距的实际行高 = 1.8 × 字号

W_MAT = CW
W_ANS = CW - 2 * ANS_MX

# ---------------------------------------------------------------- 字体
F_SONG = "华文中宋"
F_KAI = "楷体"
F_HEI = "黑体"

_FONT_CANDIDATES = {
    F_SONG: [r"C:\Windows\Fonts\STZHONGS.TTF", r"C:\Windows\Fonts\simsun.ttc"],
    F_KAI: [r"C:\Windows\Fonts\simkai.ttf", r"C:\Windows\Fonts\simsun.ttc"],
    F_HEI: [r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\msyh.ttc"],
}

# ---------------------------------------------------------------- 主题
THEMES = {
    "quiet": {                       # 默认：纯白底、浅青答案块、红色答案
        "bg": "#FFFFFF",
        "answer_fill": "#E9FFFF",
        "accent": "#E10000",
        "ink": "#000000",
        "muted": "#595959",
        "grid": "#BFBFBF",
        "info": "#E10000",
    },
    "exam": {                        # 更素的考场版：淡灰答案块、深红答案
        "bg": "#FFFFFF",
        "answer_fill": "#F2F2F2",
        "accent": "#C00000",
        "ink": "#000000",
        "muted": "#595959",
        "grid": "#BFBFBF",
        "info": "#C00000",
    },
    "brand": {                       # 汇报版：浅蓝答案块、品牌蓝点缀
        "bg": "#FFFFFF",
        "answer_fill": "#EAF3FF",
        "accent": "#1A56DB",
        "ink": "#111827",
        "muted": "#6B7280",
        "grid": "#D1D5DB",
        "info": "#1A56DB",
    },
}

STRINGS = {
    "zh": {"answer_key": "参考答案", "no_content": "（待补充）"},
    "en": {"answer_key": "Answers", "no_content": "(to be filled)"},
}


# ================================================================ 字体测量
_FC: dict = {}


def _font_file(name: str) -> str | None:
    for p in _FONT_CANDIDATES.get(name, []):
        if os.path.exists(p):
            return p
    return None


def _font(name: str, size: float):
    """按 pt 取度量字体；字号放大 4 倍以提升 getlength 精度。"""
    key = (name, round(size, 2))
    if key in _FC:
        return _FC[key]
    fp = _font_file(name)
    if fp:
        try:
            _FC[key] = ImageFont.truetype(fp, max(4, int(round(size * 4))))
            return _FC[key]
        except Exception:
            pass
    _FC[key] = None
    return None


def text_w(text: str, name: str, size: float) -> float:
    """文本宽度（磅）。无可用字体文件时按 CJK 1em / ASCII 0.5em 估算。"""
    f = _font(name, size)
    if f is not None:
        return f.getlength(text) / 4.0
    w = 0.0
    for ch in text:
        w += size * (1.0 if ord(ch) > 0x2E7F else 0.5)
    return w


NON_HEAD = "，。、；：？！）》”’%,.;:?!)]}"


def tokenize(text: str):
    """把「（　）」「（4分）」「＿＿＿」等不可拆片段当作一个整体。"""
    toks, i, n = [], 0, len(text)
    while i < n:
        if text.startswith("（　）", i):
            toks.append("（　）")
            i += 3
        elif text[i] == "（":
            j = text.find("）", i + 1)
            if 0 < j - i <= 5 and "（" not in text[i + 1:j]:
                toks.append(text[i:j + 1])
                i = j + 1
            else:
                toks.append("（")
                i += 1
        elif text[i] == "＿":
            j = i
            while j < n and text[j] == "＿":
                j += 1
            toks.append(text[i:j])
            i = j
        else:
            toks.append(text[i])
            i += 1
    return toks


def wrap(text: str, name: str, size: float, width_pt: float):
    """按真实字宽断行，返回 [(行文本, 字距压缩磅值)]。

    规则：避头（标点不在行首）、末行仅 2—4 字时并回上一行并轻度压缩字距、
    「（N分）」不单独成行。
    """
    maxw = width_pt * 0.992
    lines, cur, curw = [], [], 0.0
    for t in tokenize(text):
        w = text_w(t, name, size)
        if cur and curw + w > maxw:
            carry = []
            if t[0] in NON_HEAD:
                for _ in range(2):
                    if len(cur) > 1:
                        ch = cur.pop()
                        curw -= text_w(ch, name, size)
                        carry.insert(0, ch)
                        if ch not in NON_HEAD:
                            break
            lines.append("".join(cur))
            cur = carry + [t]
            curw = sum(text_w(x, name, size) for x in cur)
        else:
            cur.append(t)
            curw += w
    if cur:
        lines.append("".join(cur))
    lines = lines or [""]

    spcs = [0.0] * len(lines)
    if len(lines) > 1 and len(lines[-1]) <= 4:
        merged = lines[-2] + lines[-1]
        w = text_w(merged, name, size)
        if w <= maxw / 0.945:                     # 压缩幅度 ≤ 5.5%
            lines[-2] = merged
            lines.pop()
            spcs = spcs[:len(lines)]
            spcs[-1] = -(w - maxw) / max(1, len(merged) - 1)
        elif lines[-1].startswith("（"):           # （4分）不单独成行
            for _ in range(3):
                if len(lines[-2]) > 10:
                    lines[-1] = lines[-2][-1] + lines[-1]
                    lines[-2] = lines[-2][:-1]
    return list(zip(lines, spcs))


def wrap_paras(paras, name, size, width_pt):
    out = []
    for p in paras:
        out.extend(wrap(p, name, size, width_pt))
    return out


# ---------------------------------------------------------------- 图表块
TBL_PAD_X = 0.06          # 单元格左右内边距（英寸）
TBL_PAD_Y = 4             # 单元格上下内边距（磅）
TBL_LS = 1.55             # 单元格行高系数


def img_size(path: str, width_in: float):
    """按给定宽度返回 (宽, 高)，保持原始比例。"""
    from PIL import Image
    with Image.open(path) as im:
        w, h = im.size
    return width_in, width_in * h / float(w)


def table_height(rows, widths, width_in, pt):
    """估算表格高度（英寸）：逐行按最小列宽断行，取最多行数。"""
    total = float(sum(widths)) or 1.0
    h = 0.0
    for r in rows:
        n = 1
        for i, c in enumerate(r):
            cw = width_in * widths[i] / total - 2 * TBL_PAD_X
            n = max(n, len(wrap(str(c), F_SONG, pt, max(cw, 0.4) * 72)))
        h += n * pt * TBL_LS / 72 + 2 * TBL_PAD_Y / 72
    return h


def pack_opts(opts, size, width_pt):
    """选项按宽度排成若干行；单个选项过长时折行并缩进续行。"""
    maxw = width_pt * 0.992
    sep = text_w("　　", F_SONG, size)
    rows, cur, curw = [], [], 0.0
    for o in opts:
        w = text_w(o, F_SONG, size)
        if w > maxw:
            if cur:
                rows.append("　　".join(cur))
                cur, curw = [], 0.0
            # 保留 wrap() 给出的字距压缩值：末行并回时靠 spc 才真正收得进版心，
            # 丢掉 spc 会让渲染端二次折行（估算 8 行、实际 10 行 → 冲出页面）
            rows.extend((("　" + t) if i else t, spc)
                        for i, (t, spc) in enumerate(
                            wrap(o, F_SONG, size, width_pt - size)))
            continue
        if cur and curw + sep + w > maxw:
            rows.append("　　".join(cur))
            cur, curw = [o], w
        else:
            curw += (sep + w) if cur else w
            cur.append(o)
    if cur:
        rows.append("　　".join(cur))
    return rows


# ================================================================ 绘制
def _rgb(hexstr):
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_run_font(run, name, size, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color if color is not None else _rgb("#000000")
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    lat = rPr.find(qn("a:latin"))
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set("typeface", name)
    ea, cs = rPr.find(qn("a:ea")), rPr.find(qn("a:cs"))
    if lat is not None and ea is not None:
        rPr.remove(ea)
        lat.addnext(ea)
        if cs is not None:
            rPr.remove(cs)
            ea.addnext(cs)


def add_textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    return tb, tf


def add_lines(tf, lines, name, size, spacing, bold=False, color=None,
              first_space=0.0, mode="exact"):
    for i, item in enumerate(lines):
        txt, spc = item if isinstance(item, tuple) else (item, 0.0)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = spacing if mode == "multiple" else Pt(spacing)
        p.space_before = Pt(first_space if i == 0 else 0)
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = txt
        set_run_font(r, name, size, bold, color)
        if spc:
            r.font._rPr.set("spc", str(int(round(spc * 100))))


TIMING_TMPL = """<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
    <p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
      <p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>
        <p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" grpId="0" nodeType="clickEffect">
            <p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
              <p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                <p:tgtEl><p:spTgt spid="__SPID__"/></p:tgtEl>
                <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
              </p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>
              <p:animEffect transition="in" filter="slide(fromBottom)">
                <p:cBhvr><p:cTn id="7" dur="600"/><p:tgtEl><p:spTgt spid="__SPID__"/></p:tgtEl></p:cBhvr>
              </p:animEffect>
            </p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn></p:par>
      </p:childTnLst></p:cTn></p:par>
    </p:childTnLst></p:cTn>
    <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
    <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
    </p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>
</p:timing>"""


def add_fly_in(slide, shape):
    """单个形状的「飞入（自底部）」进入动画，单击触发。"""
    timing = etree.fromstring(
        TIMING_TMPL.replace("__SPID__", str(shape.shape_id)).encode("utf-8"))
    sld = slide._element
    old = sld.find(qn("p:timing"))
    if old is not None:
        sld.remove(old)
    sld.append(timing)


class Deck:
    def __init__(self, spec: dict):
        self.spec = spec
        self.params = spec.get("params") or {}
        theme_key = self.params.get("theme", "quiet")
        self.theme = dict(THEMES.get(theme_key, THEMES["quiet"]))
        self.theme.update(self.params.get("theme_overrides") or {})
        self.lang = self.params.get("language", "zh")
        self.strs = STRINGS.get(self.lang, STRINGS["zh"])
        self.answer_style = self.params.get("answer_style", "inline")
        self.animate = self.params.get("animate", True)
        self.prs = Presentation()
        self.prs.slide_width = Inches(SW)
        self.prs.slide_height = Inches(SH)
        self.log = []
        self.q_reserve = 0.0        # 由 build() 预扫描确定

    # ------------------------------------------------------------ 基础
    def new_slide(self):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = _rgb(self.theme["bg"])
        return s

    def ink(self):
        return _rgb(self.theme["ink"])

    def muted(self):
        return _rgb(self.theme["muted"])

    # ------------------------------------------------------------ 封面
    def cover(self, sp):
        s = self.new_slide()
        title = sp.get("title") or self.spec.get("title", "")
        tb, tf = add_textbox(s, 0.9, 2.10, SW - 1.8, 1.9, anchor=MSO_ANCHOR.MIDDLE)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = Pt(66)
        r = p.add_run()
        r.text = title
        set_run_font(r, F_SONG, sp.get("title_pt", 46), True, self.ink())

        sub = sp.get("subtitle", self.spec.get("subtitle", ""))
        if sub:
            tb, tf = add_textbox(s, 0.9, 4.10, SW - 1.8, 0.7)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            r = tf.paragraphs[0].add_run()
            r.text = sub
            set_run_font(r, F_SONG, 28, False, self.muted())

        foot = sp.get("footer", self.spec.get("footer", ""))
        if foot:
            tb, tf = add_textbox(s, 0.9, 5.90, SW - 1.8, 0.5)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            r = tf.paragraphs[0].add_run()
            r.text = foot
            set_run_font(r, F_KAI, 24, False, self.muted())

    # ------------------------------------------------------------ 章节页
    def section(self, sp):
        s = self.new_slide()
        tb, tf = add_textbox(s, 0.9, 2.55, SW - 1.8, 1.2, anchor=MSO_ANCHOR.MIDDLE)
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        r = tf.paragraphs[0].add_run()
        r.text = sp.get("title", "")
        set_run_font(r, F_SONG, sp.get("title_pt", 44), True, self.ink())
        sub = sp.get("subtitle")
        if sub:
            tb, tf = add_textbox(s, 0.9, 3.85, SW - 1.8, 0.6)
            tf.paragraphs[0].alignment = PP_ALIGN.CENTER
            r = tf.paragraphs[0].add_run()
            r.text = sub
            set_run_font(r, F_KAI, 26, False, self.muted())

    # ------------------------------------------------------------ 答案表
    def answer_table(self, sp):
        s = self.new_slide()
        title = sp.get("title") or ("%s　%s" % (sp.get("section", ""),
                                                self.strs["answer_key"]))
        tb, tf = add_textbox(s, 0.85, 0.62, 9.5, 0.75)
        r = tf.paragraphs[0].add_run()
        r.text = title
        set_run_font(r, F_SONG, sp.get("title_pt", 34), True, self.ink())

        answers = [str(a) for a in sp.get("answers", [])]
        if not answers:
            return
        per_row = int(sp.get("per_row", 10))
        # 答案较少时收窄到实际数量，避免右侧出现大片空列
        per_row = max(1, min(per_row, len(answers)))
        start = int(sp.get("start", 1))
        groups = [answers[i:i + per_row] for i in range(0, len(answers), per_row)]
        rows = len(groups) * 2
        cols = per_row + 1

        top = 1.85
        rh = min(0.95, (6.6 - top) / max(rows, 1))
        tw = 11.6
        c0 = 1.55
        gt = s.shapes.add_table(rows, cols, Inches(0.85), Inches(top),
                                Inches(tw), Inches(rh * rows)).table
        gt.first_row = False
        gt.horz_banding = False
        gt.columns[0].width = Inches(c0)
        cw = (tw - c0) / per_row
        for i in range(1, cols):
            gt.columns[i].width = Inches(cw)
        for i in range(rows):
            gt.rows[i].height = Inches(rh)

        n = start
        for gi, g in enumerate(groups):
            for ci in range(cols):
                cell = gt.cell(gi * 2, ci)
                if ci == 0:
                    self._cell(cell, "题号", 22, F_SONG, True, self.ink())
                    self._cell(gt.cell(gi * 2 + 1, 0), "答案", 22,
                               F_SONG, True, self.ink())
                    continue
                if ci - 1 >= len(g):
                    # 末组不足一行：留空格，不越界
                    self._cell(cell, "", 22, F_SONG, False, self.ink())
                    self._cell(gt.cell(gi * 2 + 1, ci), "", 22,
                               F_HEI, False, self.ink())
                    continue
                self._cell(cell, str(n + ci - 1), 22, F_SONG, True, self.ink())
                self._cell(gt.cell(gi * 2 + 1, ci), g[ci - 1], 26, F_HEI, True,
                           _rgb(self.theme["info"]))
            n += len(g)

    def _cell(self, cell, text, size, font, bold, color,
              align=PP_ALIGN.CENTER, wrap=False):
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(self.theme["bg"])
        self._border(cell)
        if wrap:
            cell.margin_left = cell.margin_right = Inches(TBL_PAD_X)
            cell.margin_top = cell.margin_bottom = Pt(TBL_PAD_Y)
        else:
            cell.margin_left = cell.margin_right = 0
            cell.margin_top = cell.margin_bottom = 0
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = wrap
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = 1.15 if wrap else 1.0
        r = p.add_run()
        r.text = text
        set_run_font(r, font, size, bold, color)

    # ------------------------------------------------------------ 图表块
    def _media_item(self, m):
        """把 image / table / row 块规格转成渲染项（含高度，单位英寸）。"""
        k = m.get("kind", "table")
        if k == "image":
            w = float(m.get("width") or W_MAT)
            _, hh = img_size(m["path"], w)
            return {"t": "img", "path": m["path"], "w": w, "h": hh, "x": m.get("x")}
        if k == "row":
            items = list(m.get("items") or [])
            gap = float(m.get("gap", 0.4))
            fr = m.get("widths") or [1.0 / max(len(items), 1)] * len(items)
            avail = CW - gap * (len(items) - 1)
            cells, x, hmax = [], 0.0, 0.0
            for it, f in zip(items, fr):
                sub = self._media_item(dict(it, width=avail * f))
                cells.append((sub, x))
                x += sub["w"] + gap
                hmax = max(hmax, sub["h"])
            return {"t": "row", "cells": cells, "h": hmax}
        rows = m.get("rows") or []
        widths = m.get("widths") or [1] * max(len(r) for r in rows)
        w = float(m.get("width") or W_MAT)
        pt = float(m.get("pt") or 15)
        hh = table_height(rows, widths, w, pt)
        return {"t": "tbl", "rows": rows, "widths": widths, "w": w,
                "pt": pt, "h": hh, "align": m.get("align") or [],
                "head": m.get("head", True), "x": m.get("x")}

    def _draw_media(self, s, it, y):
        if it["t"] == "row":
            for sub, off in it["cells"]:
                self._draw_media_at(s, sub, ML + off, y)
            return
        x = it.get("x")
        if x is None:
            x = ML + (CW - it["w"]) / 2
        self._draw_media_at(s, it, x, y)

    def _draw_media_at(self, s, it, x, y):
        if it["t"] == "img":
            s.shapes.add_picture(it["path"], Inches(x), Inches(y),
                                 width=Inches(it["w"]))
            return
        if it["t"] == "row":
            for sub, off in it["cells"]:
                self._draw_media_at(s, sub, x + off, y)
            return
        rows, widths = it["rows"], it["widths"]
        total = float(sum(widths)) or 1.0
        nr, nc = len(rows), len(widths)
        gt = s.shapes.add_table(nr, nc, Inches(x), Inches(y),
                                Inches(it["w"]), Inches(it["h"])).table
        gt.first_row = False
        gt.horz_banding = False
        for i, ww in enumerate(widths):
            gt.columns[i].width = Inches(it["w"] * ww / total)
        for ri, r in enumerate(rows):
            n = 1
            for i, c in enumerate(r):
                cw = it["w"] * widths[i] / total - 2 * TBL_PAD_X
                n = max(n, len(wrap(str(c), F_SONG, it["pt"], max(cw, 0.4) * 72)))
            gt.rows[ri].height = Inches(
                n * it["pt"] * TBL_LS / 72 + 2 * TBL_PAD_Y / 72)
        al = it.get("align") or []
        for ri, r in enumerate(rows):
            for ci in range(nc):
                txt = str(r[ci]) if ci < len(r) else ""
                a = al[ci] if ci < len(al) else "left"
                align = {"center": PP_ALIGN.CENTER,
                         "right": PP_ALIGN.RIGHT}.get(a, PP_ALIGN.LEFT)
                bold = bool(it.get("head", True)) and ri == 0
                self._cell(gt.cell(ri, ci), txt, it["pt"], F_SONG, bold,
                           self.ink(), align=align, wrap=True)

    def _border(self, cell):
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for e in tcPr.findall(qn(tag)):
                tcPr.remove(e)
        for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):
            ln = etree.Element(qn(tag))
            ln.set("w", "9525")
            ln.set("cap", "flat")
            ln.set("cmpd", "sng")
            ln.set("algn", "ctr")
            sf = etree.SubElement(ln, qn("a:solidFill"))
            etree.SubElement(sf, qn("a:srgbClr")).set("val",
                                                      self.theme["grid"].lstrip("#"))
            etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")
            tcPr.insert(0, ln)

    # ------------------------------------------------------------ 选择题页
    def _choice_opts(self, tf, orow, size, first=False):
        """把选项行写进文本框；first=True 时沿用首个空段落（独立选项框）。

        orow 的元素可以是 str，也可以是 (文本, 字距压缩磅值)。
        """
        for i, item in enumerate(orow):
            row, spc = item if isinstance(item, tuple) else (item, 0.0)
            p = tf.paragraphs[0] if (first and i == 0) else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.5
            if not first:
                p.space_before = Pt(size * 0.40 if i == 0 else 0)
            else:
                p.space_before = Pt(0)
            p.space_after = Pt(0)
            r = p.add_run()
            r.text = row
            set_run_font(r, F_SONG, size, False, self.ink())
            if spc:
                r.font._rPr.set("spc", str(int(round(spc * 100))))

    def choice(self, sp):
        s = self.new_slide()
        no = sp.get("no")
        stem = sp.get("stem", "")
        if no is not None:
            stem = "%s．%s" % (no, stem)
        items = list(sp.get("items") or [])
        opts = list(sp.get("options") or [])
        media = sp.get("media")
        mit = self._media_item(media) if media else None
        mh_pt = (mit["h"] * 72 + 22) if mit else 0.0

        size = 28
        for cand in (28, 27, 26, 25, 24, 23, 22, 21, 20):
            sl = wrap(stem, F_SONG, cand, CW * 72)
            il = [x for it in items for x in wrap(it, F_SONG, cand, CW * 72)]
            orow = pack_opts(opts, cand, CW * 72)
            h = ((len(sl) + len(il) + len(orow)) * cand * CHOICE_LF
                 + cand * 0.40 + mh_pt)
            if h <= AVAIL_PT:
                size = cand
                break
        sl = wrap(stem, F_SONG, size, CW * 72)
        il = [x for it in items for x in wrap(it, F_SONG, size, CW * 72)]
        orow = pack_opts(opts, size, CW * 72)

        nline = len(sl) + len(il) + len(orow)
        h_pt = nline * size * CHOICE_LF + size * 0.40
        y0 = TOP                                    # 统一顶端，不垂直居中

        if not mit:
            tb, tf = add_textbox(s, ML, y0, CW, h_pt / 72 + 0.10)
            add_lines(tf, sl + il, F_SONG, size, 1.5, bold=False,
                      color=self.ink(), mode="multiple")
            self._choice_opts(tf, orow, size)
        else:
            # 题干/小项 → 图表 → 选项，三段依次下排
            h1 = (len(sl) + len(il)) * size * CHOICE_LF / 72.0
            tb, tf = add_textbox(s, ML, y0, CW, h1 + 0.04)
            add_lines(tf, sl + il, F_SONG, size, 1.5, bold=False,
                      color=self.ink(), mode="multiple")
            self._draw_media(s, mit, y0 + h1 + 4 / 72.0)
            if orow:
                tb, tf = add_textbox(s, ML, y0 + h1 + mit["h"] + 18 / 72.0,
                                     CW, len(orow) * size * CHOICE_LF / 72.0 + 0.04)
                self._choice_opts(tf, orow, size, first=True)

        ans = sp.get("answer")
        if not ans or self.answer_style == "none":
            self.log.append("choice %s: %.0fpt" % (no, size))
            return size

        if self.answer_style == "right_big":
            ab, atf = add_textbox(s, 11.35, 2.6, 1.6, 2.0, anchor=MSO_ANCHOR.MIDDLE)
            atf.paragraphs[0].alignment = PP_ALIGN.CENTER
            ar = atf.paragraphs[0].add_run()
            ar.text = str(ans)
            set_run_font(ar, F_HEI, 96, False, _rgb(self.theme["accent"]))
            if self.animate:
                add_fly_in(s, ab)
            self.log.append("choice %s: %.0fpt" % (no, size))
            return size

        # inline：答案字母嵌入题干括号，红色，与题干基线对齐
        li, k = -1, -1
        for i, (ln, _) in enumerate(sl):
            p = ln.find("（　）")
            if p >= 0:
                li, k = i, p
                break
        if k < 0:
            self.log.append("choice %s: %.0fpt (无括号，答案未嵌入)" % (no, size))
            return size

        x_off = text_w(sl[li][0][:k], F_SONG, size) + k * sl[li][1]
        box = size * 1.02
        bx = ML + (x_off + 1.5 * size - box / 2) / 72.0
        by = y0 + li * size * CHOICE_LF / 72.0

        ab, atf = add_textbox(s, bx, by, box / 72.0, size * CHOICE_LF / 72.0)
        ap = atf.paragraphs[0]
        ap.alignment = PP_ALIGN.CENTER
        ap.line_spacing = 1.5
        ap.space_before = Pt(0)
        ap.space_after = Pt(0)
        ar = ap.add_run()
        ar.text = str(ans)
        set_run_font(ar, F_HEI, size, True, _rgb(self.theme["accent"]))
        if self.animate:
            add_fly_in(s, ab)
        self.log.append("choice %s: %.0fpt" % (no, size))
        return size

    # ------------------------------------------------------------ 阅读/解题页
    @staticmethod
    def _block_h(lines, spacing, pad=0.0):
        return len(lines) * spacing + pad

    def _split_blocks(self, blocks):
        """按页面容量切分：材料可逐行续页，图表占位，答案放不下时独占一页并重印问题。"""
        pages, cur = [], {"mat": [], "q": [], "ans": [], "label": None}
        h = 0.0

        def flush():
            nonlocal cur
            pages.append(cur)
            cur = {"mat": [], "q": [], "ans": [], "label": None}

        for b in blocks:
            kind = b.get("kind", "material")
            if kind == "material":
                lines = wrap_paras(b.get("text") or [], F_KAI, MAT_PT,
                                   W_MAT * 72)
                i = 0
                while i < len(lines):
                    rem = int((APP_AVAIL_PT - h) // MAT_SP)
                    if rem < 1:
                        flush()
                        h = 0.0
                        rem = int(APP_AVAIL_PT // MAT_SP)
                    left = len(lines) - i
                    take = min(rem, left)
                    # 末页只剩 1—2 行时，从上一页匀几行下来
                    if left > rem and left - take < 3 and take > 4:
                        take = left - 3
                    for ln in lines[i:i + take]:
                        cur["mat"].append({"t": "line", "text": ln[0],
                                           "spc": ln[1]})
                        h += MAT_SP
                    i += take
                    if i < len(lines):
                        flush()
                        h = 0.0
            elif kind in ("image", "table", "row"):
                it = self._media_item(b)
                ih = it["h"] * 72 + 10
                if h + ih > APP_AVAIL_PT and cur["mat"]:
                    flush()
                    h = 0.0
                cur["mat"].append(it)
                h += ih
            elif kind == "question":
                ql = wrap_paras(b.get("text") or [], F_SONG, Q_PT, W_MAT * 72)
                need = len(ql) * Q_SP + (2 if cur["mat"] else 0)
                if h + need > APP_AVAIL_PT:
                    flush()
                    h = 0.0
                cur["q"] = ql
                h += need
            elif kind == "answer":
                # 答案不跟在材料后面：没有材料时落回统一起始线，有材料则另起一页
                if cur["mat"] or cur["ans"]:
                    flush()
                    h = 0.0
                al = wrap_paras(b.get("text") or [], F_HEI, ANS_PT, W_ANS * 72)
                cap = max(1, self._ans_cap())
                # 末页只剩 1—2 行时，先从首页匀几行下来
                first_cap = cap
                if len(al) > cap and len(al) - cap < 3:
                    first_cap = max(4, len(al) - 3)
                i = 0
                while i < len(al):
                    c = first_cap if i == 0 else cap
                    piece = al[i:i + c]
                    i += len(piece)
                    cur["ans"] = list(cur["ans"]) + piece
                    if i < len(al):
                        flush()
                        h = 0.0
                        cur["cont"] = True
        pages.append(cur)
        return pages

    def _prep_blocks(self, sp):
        """复制 blocks 并把题号并入首段（有材料并材料首段，否则并入问题）。"""
        blocks = [dict(b) for b in (sp.get("blocks") or [])]
        label = sp.get("label")
        if label:
            for b in blocks:
                if b.get("kind") in ("material", "question"):
                    t = list(b.get("text") or [])
                    if t:
                        t[0] = "%s．%s" % (label, t[0])
                        b["text"] = t
                        break
        return blocks

    def reading(self, sp):
        blocks = self._prep_blocks(sp)
        label = sp.get("label")
        pages = self._split_blocks(blocks)
        a_top = self._ans_top()
        for pg in pages:
            s = self.new_slide()
            ans_only = bool(pg["ans"]) and not pg["mat"] and not pg["q"]
            y = APP_TOP
            if ans_only and not pg.get("cont"):
                qtext = self._question_text(sp)
                if qtext:
                    ql2 = wrap_paras([qtext], F_SONG, Q_PT, W_MAT * 72)
                    tb, tf = add_textbox(s, ML, APP_TOP, W_MAT,
                                         self.q_reserve + 0.03)
                    add_lines(tf, ql2, F_SONG, Q_PT, Q_SP, bold=True,
                              color=self.ink())
            if ans_only:
                y = a_top

            if pg["mat"]:
                y = self._draw_flow(s, pg["mat"], y)
            if pg["q"]:
                if pg["mat"]:
                    y += 2 / 72.0
                hh = len(pg["q"]) * Q_SP / 72.0
                tb, tf = add_textbox(s, ML, y, W_MAT, hh + 0.03)
                add_lines(tf, pg["q"], F_SONG, Q_PT, Q_SP, bold=True,
                          color=self.ink())
                y += hh
            if pg["ans"]:
                y = (y + 6 / 72.0) if pg["mat"] else a_top
                hh = len(pg["ans"]) * ANS_SP / 72.0 + 2 * ANS_PAD / 72.0
                shape = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(ML),
                                           Inches(y), Inches(CW), Inches(hh))
                shape.fill.solid()
                shape.fill.fore_color.rgb = _rgb(self.theme["answer_fill"])
                shape.line.fill.background()
                shape.shadow.inherit = False
                tf = shape.text_frame
                tf.word_wrap = True
                tf.auto_size = MSO_AUTO_SIZE.NONE
                tf.vertical_anchor = MSO_ANCHOR.TOP
                tf.margin_left = tf.margin_right = Inches(ANS_MX)
                tf.margin_top = tf.margin_bottom = Pt(ANS_PAD)
                add_lines(tf, pg["ans"], F_HEI, ANS_PT, ANS_SP, bold=True,
                          color=self.ink())
                if self.animate:
                    add_fly_in(s, shape)
        self.log.append("reading %s: %d 页" % (label or "", len(pages)))
        return len(pages)

    def _draw_flow(self, s, items, y):
        """按顺序绘制「文字行 / 图片 / 表格」，返回末尾 y（英寸）。"""
        buf = []

        def flush_buf():
            nonlocal y, buf
            if not buf:
                return
            hh = len(buf) * MAT_SP / 72.0
            tb, tf = add_textbox(s, ML, y, W_MAT, hh + 0.03)
            add_lines(tf, buf, F_KAI, MAT_PT, MAT_SP, color=self.ink())
            y += hh
            buf = []

        for it in items:
            if it["t"] == "line":
                buf.append((it["text"], it.get("spc", 0.0)))
                continue
            flush_buf()
            self._draw_media(s, it, y)
            y += it["h"] + 10 / 72.0
        flush_buf()
        return y

    def _question_text(self, sp):
        label = sp.get("label")
        qt = ""
        for b in (sp.get("blocks") or []):
            if b.get("kind") == "question":
                qt = "".join(b.get("text") or [])
                break
        if not qt:
            return ""
        return ("%s．%s" % (label, qt)) if label else qt

    def _ans_top(self):
        """所有答案色块的统一起始线（英寸），全篇一致。"""
        return APP_TOP + self.q_reserve + Q_GAP_BEFORE_ANS

    def _ans_cap(self):
        """答案页色块最多容纳的行数。

        以「正文底边不越过下边距」为准（色块内部的下内边距可略微吃掉余量）。
        """
        room = (SH - BOT - self._ans_top()) * 72 - ANS_PAD
        return max(1, int(room // ANS_SP))

    # ------------------------------------------------------------ 要点页
    def bullets(self, sp):
        title = sp.get("title", "")
        t_top = 0.585          # 34 号标题（固定 44 磅行距）补偿到 42pt 基线
        items = list(sp.get("items") or [])
        per_page = int(sp.get("per_page") or self.params.get("bullets_per_page") or 6)
        col = int(sp.get("columns", 1))
        width = (CW - 0.4) / col if col > 1 else CW
        chunks = [items[i:i + per_page * col] for i in range(0, max(len(items), 1),
                                                           per_page * col)] or [[]]
        for pi, chunk in enumerate(chunks):
            s = self.new_slide()
            y = t_top if title else APP_TOP
            if title:
                tb, tf = add_textbox(s, ML, y, CW, 0.8)
                p = tf.paragraphs[0]
                p.line_spacing = Pt(44)       # 固定行距，首行字顶才与全篇对齐
                r = p.add_run()
                r.text = title if pi == 0 else "%s（续）" % title
                set_run_font(r, F_SONG, 34, True, self.ink())
                y += 0.92
            if chunk:
                size = float(sp.get("body_pt", 28))
                body = ["•　%s" % it for it in chunk]
                lines = wrap_paras(body, F_SONG, size, width * 72)
                hh = len(lines) * size * 1.6 / 72.0
                tb, tf = add_textbox(s, ML, y, width, hh + 0.05)
                add_lines(tf, lines, F_SONG, size, size * 1.6, color=self.ink())
        self.log.append("bullets: %d 页" % len(chunks))

    # ------------------------------------------------------------ 组装
    def build(self, out_path):
        # 预扫描：所有「问题」在最宽行宽下的最大行数 → 决定答案色块的统一起始线
        mx = 0
        for sp in self.spec.get("slides", []):
            if sp.get("type") != "reading":
                continue
            qt = self._question_text(sp)
            if qt:
                mx = max(mx, len(wrap(qt, F_SONG, Q_PT, W_MAT * 72)))
        self.q_reserve = mx * Q_SP / 72.0

        for sp in self.spec.get("slides", []):
            t = sp.get("type")
            if t == "cover":
                self.cover(sp)
            elif t == "section":
                self.section(sp)
            elif t == "answer_table":
                self.answer_table(sp)
            elif t == "choice":
                self.choice(sp)
            elif t == "reading":
                self.reading(sp)
            elif t == "bullets":
                self.bullets(sp)
            elif t == "blank":
                self.new_slide()
            else:
                raise ValueError("未知页面类型: %r" % t)
        self.prs.save(out_path)
        return len(self.prs.slides._sldIdLst)


def main(argv):
    if len(argv) < 2:
        print("用法: python deck_builder.py spec.json [out.pptx]")
        return 2
    spec_path = argv[1]
    out = argv[2] if len(argv) > 2 else os.path.splitext(spec_path)[0] + ".pptx"
    with open(spec_path, encoding="utf-8") as fh:
        spec = json.load(fh)
    deck = Deck(spec)
    n = deck.build(out)
    for line in deck.log:
        print(" ", line)
    print("总页数:", n)
    print("saved:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
