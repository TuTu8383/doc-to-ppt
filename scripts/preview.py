# -*- coding: utf-8 -*-
"""生成预览并自检：PPTX → PDF → 逐页量测

用法：
    python preview.py deck.pptx [--out DIR] [--dpi 110]

做的事：
1. 调用 LibreOffice 或 WPS 把 pptx 转成 pdf
2. 用 pymupdf 逐页量测：首行字顶、行底越界、右边界越界
3. 用 python-pptx 检查：答案色块起始线是否一致、飞入动画页数、是否有切换效果
4. 输出一份可读报告，并把每页 PNG 存到 DIR
"""
from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys

PAGE_W_PT = 13.333 * 72          # 960
PAGE_H_PT = 7.5 * 72             # 540
SAFE_BOTTOM = 520                # 允许的最低文字底部（pt）
SAFE_RIGHT = 955                 # 允许的最右文字边界（pt）


def find_converter():
    for exe in ("soffice", "libreoffice"):
        p = shutil.which(exe)
        if p:
            return ("soffice", p)
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(p):
            return ("soffice", p)
    hits = glob.glob(os.path.expandvars(
        r"%LOCALAPPDATA%\Kingsoft\WPS Office\*\office6\kwpsconvert.exe"))
    if hits:
        return ("wps", sorted(hits)[-1])
    return (None, None)


def to_pdf(pptx, pdf):
    kind, exe = find_converter()
    if not kind:
        print("[!] 未找到 LibreOffice / WPS，跳过 PDF 预览")
        return False
    if os.path.exists(pdf):
        os.remove(pdf)
    if kind == "soffice":
        subprocess.run([exe, "--headless", "--convert-to", "pdf",
                        "--outdir", os.path.dirname(pdf) or ".", pptx],
                       check=False, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, timeout=300)
        made = os.path.join(os.path.dirname(pdf) or ".",
                            os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
        if made != pdf and os.path.exists(made):
            os.replace(made, pdf)
    else:
        subprocess.run([exe, "ppt2pdf", os.path.abspath(pptx),
                        "-o", os.path.abspath(pdf)],
                       check=False, stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL, timeout=300)
    return os.path.exists(pdf)


def main(argv):
    if len(argv) < 2:
        print("用法: python preview.py deck.pptx [--out DIR] [--dpi 110]")
        return 2
    pptx = argv[1]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(pptx)), "_preview")
    dpi = 110
    if "--out" in argv:
        out_dir = argv[argv.index("--out") + 1]
    if "--dpi" in argv:
        dpi = int(argv[argv.index("--dpi") + 1])
    os.makedirs(out_dir, exist_ok=True)
    pdf = os.path.join(out_dir, "preview.pdf")

    ok = to_pdf(pptx, pdf)
    print("=" * 60)
    print("文件:", os.path.basename(pptx))

    # ---------------- pptx 结构检查
    from pptx import Presentation
    from pptx.oxml.ns import qn
    prs = Presentation(pptx)
    n = len(prs.slides._sldIdLst)
    anim = [i for i, s in enumerate(prs.slides, 1)
            if s._element.find(qn("p:timing")) is not None]
    trans = [i for i, s in enumerate(prs.slides, 1)
             if s._element.find(qn("p:transition")) is not None]
    print("页数:", n)
    print("飞入动画页数:", len(anim), anim if len(anim) < 45 else "")
    print("切换效果页:", trans if trans else "无（符合要求）")

    tops, worst = {}, []
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            try:
                if sh.shape_type == 1 and sh.fill.type == 1:
                    tops.setdefault(str(sh.fill.fore_color.rgb), []).append(
                        (i, round(sh.top / 914400, 3)))
            except Exception:
                pass
    for color, lst in tops.items():
        if len(lst) >= 2:
            ys = sorted({y for _, y in lst})
            print("色块 %s: %d 个, 起始线 %s" % (color, len(lst), ys))

    if ok:
        import pymupdf
        doc = pymupdf.open(pdf)
        print("PDF 页数:", doc.page_count)
        first = []
        for i, page in enumerate(doc, 1):
            lines = [l["bbox"] for b in page.get_text("dict")["blocks"]
                     for l in b.get("lines", [])]
            if not lines:
                continue
            first.append(round(min(b[1] for b in lines), 1))
            for b in lines:
                if b[3] > SAFE_BOTTOM or b[2] > SAFE_RIGHT:
                    worst.append((i, round(b[2], 1), round(b[3], 1)))
        print("首行字顶取值:", sorted(set(first)))
        print("越界/贴边行:", worst if worst else "0（符合要求）")
        for i, page in enumerate(doc, 1):
            page.get_pixmap(dpi=dpi).save(
                os.path.join(out_dir, "s%02d.png" % i))

    print("预览图:", out_dir)
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
