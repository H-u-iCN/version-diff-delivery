#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""DOCX / Word 文档的近似分页渲染 + 接触印相表 + 新旧对比。

不依赖 LibreOffice / Word：用 python-docx 读 WordprocessingML，按页面尺寸与页边距
把段落、表格、内嵌图按顺序流式排版，超出页高就分页，再拼成接触印相表。

子命令：
  sheet   --docx <file> --out <sheet.png> [--cols 4] [--thumb-w 440]
  compare --left <old.docx> --right <new.docx> --out <cmp.png> [--thumb-w 380]
  pages   --docx <file> --outdir <dir>

注意：分页是近似的。Word 的排版引擎（字距、行距、标点挤压、孤行控制、样式继承）
远比这里复杂，页的切分位置可能与 Word 实际分页有出入；内容本身是按顺序完整排下来的。
足够看出「改了哪儿」，不能当排版定稿依据。
"""

import argparse
import json
import os
import sys
from io import BytesIO

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover
    print(json.dumps({"ok": False, "error": "缺少 Pillow，请先 pip install Pillow"}, ensure_ascii=False))
    sys.exit(2)

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except ImportError:  # pragma: no cover
    print(json.dumps({"ok": False, "error": "缺少 python-docx，请先 pip install python-docx"},
                     ensure_ascii=False))
    sys.exit(2)

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.dont_write_bytecode = True  # 别在 skill 包里留下 __pycache__
from sheetkit import build_compare, build_sheet, font

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # pragma: no cover
    pass

PPI = 110
EMU_PER_INCH = 914400.0
DEFAULT_SIZE = 10.5  # 五号，Word 中文文档常用正文大小
LINE_RATIO = 1.5      # 行高倍数
C_TEXT = (32, 34, 40)
C_RULE = (143, 166, 191)
C_BORDER = (200, 206, 214)


def emu_px(v):
    return (v or 0) / EMU_PER_INCH * PPI


def rgb_of(color_obj, default=C_TEXT):
    try:
        if color_obj is None or color_obj.type is None:
            return default
        return tuple(color_obj.rgb)
    except Exception:
        return default


def run_font(run, para, default_size):
    size = None
    try:
        if run.font.size is not None:
            size = run.font.size.pt
    except Exception:
        size = None
    if size is None:
        try:
            if para.style is not None and para.style.font.size is not None:
                size = para.style.font.size.pt
        except Exception:
            size = None
    if size is None:
        size = default_size
    bold = False
    try:
        if run.font.bold:
            bold = True
    except Exception:
        pass
    kind = 'doc_bold' if bold else 'doc'
    return size, kind, rgb_of(run.font.color)


def wrap_text(draw, text, fnt, max_px):
    """按字符宽度折行（中文逐字断行，英文按空格优先）。"""
    if not text:
        return []
    lines = []
    cur = ''
    for ch in text:
        if ch == '\n':
            lines.append(cur)
            cur = ''
            continue
        trial = cur + ch
        if draw.textlength(trial, font=fnt) > max_px and cur:
            # 英文单词尽量不拆：回退到最后一个空格
            cut = cur.rfind(' ')
            if cut > 0 and ch.isascii() and ch not in ' ':
                lines.append(cur[:cut])
                cur = cur[cut + 1:] + ch
            else:
                lines.append(cur)
                cur = ch
        else:
            cur = trial
    lines.append(cur)
    return lines


class Flow:
    """把内容按页宽流式排到多张页图上。"""

    def __init__(self, page_w_emu, page_h_emu, ml, mr, mt, mb):
        self.W = int(emu_px(page_w_emu))
        self.H = int(emu_px(page_h_emu))
        self.left = emu_px(ml)
        self.right = self.W - emu_px(mr)
        self.top = emu_px(mt)
        self.bottom = self.H - emu_px(mb)
        self.width = max(40, self.right - self.left)
        self.pages = []
        self.new_page()

    def new_page(self):
        img = Image.new('RGB', (self.W, self.H), (255, 255, 255))
        self.pages.append(img)
        self.d = ImageDraw.Draw(img)
        self.y = self.top

    def space_left(self):
        return self.bottom - self.y

    def ensure(self, h):
        if self.y + h > self.bottom:
            self.new_page()
            return True
        return False

    def draw_line(self, segments, line_h, align):
        """segments: [(text, fnt, color)]，按对齐方式落一行。"""
        total = sum(self.d.textlength(t, font=f) for t, f, _ in segments)
        if 'CENTER' in align:
            x = self.left + (self.width - total) / 2.0
        elif 'RIGHT' in align:
            x = self.right - total
        else:
            x = self.left
        for t, f, c in segments:
            self.d.text((x, self.y), t, font=f, fill=c)
            x += self.d.textlength(t, font=f)
        self.y += line_h


def paragraph_items(para, doc):
    """把段落拆成文本块与图片块，并标记分页符。"""
    items = []
    for run in para.runs:
        el = run._element
        brs = el.findall('.//' + qn('w:br'))
        for br in brs:
            if br.get(qn('w:type')) == 'page':
                items.append(('pagebreak', None, None, None, None))
        blips = el.findall('.//' + qn('a:blip'))
        if blips:
            rid = blips[0].get(qn('r:embed'))
            try:
                blob = doc.part.related_parts[rid].blob
            except Exception:
                blob = None
            ext = el.findall('.//' + qn('wp:extent'))
            cx = int(ext[0].get('cx')) if ext else 0
            cy = int(ext[0].get('cy')) if ext else 0
            if blob:
                items.append(('image', blob, cx, cy, None))
            continue
        if run.text:
            size, kind, color = run_font(run, para, DEFAULT_SIZE)
            items.append(('text', run.text, size, kind, color))
    return items


def layout_paragraph(flow, para, doc, errors, page_no_getter):
    align = str(para.alignment) if para.alignment is not None else 'LEFT'
    items = paragraph_items(para, doc)
    if not items:
        flow.y += 6
        return
    # 一行由若干 (text, font, color) 段组成
    line = []
    line_h = 0
    para_w = flow.width

    def flush():
        nonlocal line, line_h
        if line:
            flow.draw_line(line, line_h or 14, align)
            line = []
            line_h = 0

    for kind, payload, a, b, c in items:
        if kind == 'pagebreak':
            flush()
            flow.new_page()
            continue
        if kind == 'image':
            w = emu_px(a)
            h = emu_px(b)
            try:
                im = Image.open(BytesIO(payload)).convert('RGBA')
            except Exception as exc:
                errors.append('图片解码失败: %s' % exc)
                continue
            if w <= 0 or h <= 0:
                w = min(emu_px(a), 0) or im.width
                h = im.height
            if w > para_w:
                scale = para_w / float(w)
                w, h = int(w * scale), int(h * scale)
            if h > flow.space_left():
                flush()
                flow.ensure(h)
                if h > flow.space_left():
                    h = flow.space_left()
            flush()
            im = im.resize((max(1, int(w)), max(1, int(h))), Image.LANCZOS)
            flow.pages[-1].paste(im, (int(flow.left), int(flow.y)), im)
            flow.y += h + 4
            continue
        # 文本：元组为 (kind, text, size, fontkind, color)
        text, size, fkind, color = payload, a, b, c
        fnt = font(fkind, size * PPI / 72.0)
        lh = size * PPI / 72.0 * LINE_RATIO
        for seg in text.split('\n'):
            for ln in wrap_text(flow.d, seg, fnt, para_w):
                if not ln:
                    flush()
                    flow.y += lh
                    continue
                tw = flow.d.textlength(ln, font=fnt)
                if line and (flow.d.textlength(''.join(t for t, _, _ in line), font=fnt) + tw) > para_w:
                    flush()
                line.append((ln, fnt, color))
                line_h = max(line_h, lh)
                if flow.space_left() < line_h:
                    flush()
                    flow.ensure(line_h)
                    line.append((ln, fnt, color))
                    line_h = max(line_h, lh)
    flush()
    flow.y += 5


def layout_table(flow, table, errors):
    grid = table._tbl.find(qn('w:tblGrid'))
    cols = []
    if grid is not None:
        for gc in grid.findall(qn('w:gridCol')):
            cols.append(int(gc.get(qn('w:w')) or 0) * 635)
    ncol = max(len(cols), 1)
    total = sum(cols) if cols else 0
    if not total:
        cols = [flow.width / ncol] * ncol
    else:
        cols = [c * (flow.width / total) for c in cols]
    for row in table.rows:
        cells = row.cells
        heights = []
        wrapped = []
        for ci, cell in enumerate(cells):
            cw = cols[ci] if ci < len(cols) else cols[-1]
            text = cell.text
            fnt = font('doc', 9 * PPI / 72.0)
            lines = []
            for seg in text.split('\n'):
                lines.extend(wrap_text(flow.d, seg, fnt, max(20, cw - 8)))
            wrapped.append((lines, fnt))
            heights.append(len(lines) * 9 * PPI / 72.0 * 1.35 + 6)
        rh = max(heights) if heights else 20
        if flow.space_left() < rh:
            flow.new_page()
        x = flow.left
        for ci, cell in enumerate(cells):
            cw = cols[ci] if ci < len(cols) else cols[-1]
            y = flow.y
            flow.d.rectangle([x, y, x + cw, y + rh], outline=C_RULE, width=1)
            lines, fnt = wrapped[ci]
            ty = y + 3
            for ln in lines:
                flow.d.text((x + 4, ty), ln, font=fnt, fill=C_TEXT)
                ty += 9 * PPI / 72.0 * 1.35
            x += cw
        flow.y += rh
    flow.y += 8


def render_docx(docx_path, outdir, ppi=None):
    if not os.path.isdir(outdir):
        os.makedirs(outdir, exist_ok=True)
    doc = Document(docx_path)
    sec = doc.sections[0]
    flow = Flow(sec.page_width, sec.page_height, sec.left_margin, sec.right_margin,
                sec.top_margin, sec.bottom_margin)
    errors = []
    for child in doc.element.body.iterchildren():
        tag = child.tag.split('}')[-1]
        if tag == 'p':
            layout_paragraph(flow, Paragraph(child, doc), doc, errors, None)
        elif tag == 'tbl':
            try:
                layout_table(flow, Table(child, doc), errors)
            except Exception as exc:
                errors.append('表格渲染失败: %s' % exc)
    paths = []
    for i, img in enumerate(flow.pages, start=1):
        p = os.path.join(outdir, 'page-%02d.png' % i)
        img.save(p)
        paths.append(p)
    return paths, errors, (flow.W, flow.H)


def main():
    ap = argparse.ArgumentParser(description="DOCX 接触印相表 / 新旧对比")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("sheet", help="整份文档拼成一张接触印相表")
    s.add_argument("--docx", required=True)
    s.add_argument("--out", required=True)
    s.add_argument("--cols", type=int, default=4)
    s.add_argument("--thumb-w", type=int, default=440)
    s.add_argument("--title", default="")
    s.add_argument("--workdir", default="")

    c = sub.add_parser("compare", help="新旧两版逐页并排对比")
    c.add_argument("--left", required=True)
    c.add_argument("--right", required=True)
    c.add_argument("--out", required=True)
    c.add_argument("--thumb-w", type=int, default=380)
    c.add_argument("--title", default="")
    c.add_argument("--workdir", default="")
    c.add_argument("--no-mark", action="store_true")

    p = sub.add_parser("pages", help="逐页导出 PNG")
    p.add_argument("--docx", required=True)
    p.add_argument("--outdir", required=True)

    args = ap.parse_args()
    tmp = args.workdir or os.path.join(os.path.dirname(os.path.abspath(args.out)) or ".", "_doc_work")
    try:
        if args.cmd == "pages":
            pages, errors, size = render_docx(args.docx, args.outdir)
            res = {"ok": True, "mode": "pages", "files": pages, "pages": len(pages),
                   "errors": errors, "page_size": size}
        elif args.cmd == "sheet":
            work = os.path.join(tmp, os.path.splitext(os.path.basename(args.docx))[0])
            pages, errors, size = render_docx(args.docx, work)
            out, dims = build_sheet(pages, args.out, args.cols, args.thumb_w, args.title)
            res = {"ok": True, "mode": "sheet", "files": [out], "pages": len(pages),
                   "errors": errors, "image_size": dims}
        else:
            lp, err_l, _ = render_docx(args.left, os.path.join(tmp, "old"))
            rp, err_r, _ = render_docx(args.right, os.path.join(tmp, "new"))
            out, dims, changed = build_compare(lp, rp, args.out, args.thumb_w, args.title,
                                               mark_changes=not args.no_mark)
            res = {"ok": True, "mode": "compare", "files": [out], "pages": len(rp),
                   "changed_pages": changed, "image_size": dims, "errors": err_l + err_r}
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": "%s: %s" % (type(exc).__name__, exc)},
                         ensure_ascii=False, indent=2))
        sys.exit(1)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
