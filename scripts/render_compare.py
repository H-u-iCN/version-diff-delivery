#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""前后对比图渲染（Pillow）。把"旧版 vs 新版"渲染成可直接交付给用户的 PNG。

三种模式：
  text  --left A --right B      代码 / 文本 / 配置 / Markdown：左右并排，红删绿增，带行号
  image --left A --right B      位图（png/jpg/webp）：左右并排，并标出变化区域
  pair  --left A.png B.png --right C.png D.png    已渲染好的页面/幻灯片：逐页并排网格

所有命令向 stdout 输出 JSON，例如：
  {"ok": true, "mode": "text", "files": ["out.png"], "rows": 42, "added": 12, "removed": 3}
"""

import argparse
import difflib
import json
import os
import sys
from datetime import datetime

try:
    from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont
except ImportError:  # pragma: no cover
    print(json.dumps({"ok": False, "error": "缺少 Pillow，请先 pip install Pillow"}, ensure_ascii=False))
    sys.exit(2)

try:  # Windows 控制台默认 GBK，强制 UTF-8 避免中文结果乱码
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # pragma: no cover
    pass

MONO_FONTS = [
    r"C:\Windows\Fonts\consola.ttf",
    r"C:\Windows\Fonts\CascadiaMono.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
    "/System/Library/Fonts/Menlo.ttc",
]
CJK_FONTS = [
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]

C_BG = (255, 255, 255)
C_BORDER = (222, 225, 231)
C_HEAD = (30, 36, 48)
C_HEAD_FG = (255, 255, 255)
C_DEL_BG = (255, 235, 235)
C_DEL_FG = (146, 34, 34)
C_DEL_GUT = (246, 219, 219)
C_INS_BG = (231, 248, 236)
C_INS_FG = (20, 102, 52)
C_INS_GUT = (220, 240, 227)
C_EQ_FG = (42, 46, 54)
C_GUT_BG = (247, 248, 250)
C_GUT_FG = (134, 141, 154)
C_SEP_BG = (238, 240, 245)
C_SEP_FG = (108, 116, 132)
C_BOX = (255, 138, 0)


# ---------------------------------------------------------------- 字体

def _try_font(paths, size):
    for p in paths:
        if os.path.exists(p):
            for args in ((p, size, 0), (p, size)):
                try:
                    return ImageFont.truetype(*args)
                except Exception:
                    continue
    return ImageFont.load_default()


def load_fonts(size):
    return {"mono": _try_font(MONO_FONTS, size), "cjk": _try_font(CJK_FONTS, size)}


def _runs(text):
    """按 拉丁 / CJK 切段，保证中文与等宽 ASCII 混排时都能画出来。"""
    out = []
    cur, cur_cls = "", None
    for ch in text:
        cls = "cjk" if ord(ch) >= 0x2E80 else "mono"
        if cur_cls is None or cls == cur_cls:
            cur += ch
            cur_cls = cls
        else:
            out.append((cur_cls, cur))
            cur, cur_cls = ch, cls
    if cur:
        out.append((cur_cls, cur))
    return out


def measure(text, fonts):
    return sum(fonts[cls].getlength(run) for cls, run in _runs(text))


def draw_text(draw, xy, text, fonts, fill):
    x, y = xy
    for cls, run in _runs(text):
        f = fonts[cls]
        draw.text((x, y), run, font=f, fill=fill)
        x += f.getlength(run)
    return x


def line_height(fonts):
    hs = []
    for f in fonts.values():
        asc, desc = f.getmetrics()
        hs.append(asc + desc)
    return max(hs)


# ---------------------------------------------------------------- 文本 diff

def read_lines(path):
    with open(path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8", "utf-8-sig", "gbk", "latin-1"):
        try:
            return raw.decode(enc).splitlines()
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace").splitlines()


def build_rows(left, right):
    """返回逐行配对的行模型。kind: equal / del / ins / mod。"""
    rows = []
    sm = difflib.SequenceMatcher(a=left, b=right, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                rows.append({"ln": i1 + k + 1, "rn": j1 + k + 1,
                             "lt": left[i1 + k], "rt": right[j1 + k],
                             "kind": "equal", "del": False})
            continue
        lblock = left[i1:i2]
        rblock = right[j1:j2]
        for k in range(max(len(lblock), len(rblock))):
            ltxt = lblock[k] if k < len(lblock) else None
            rtxt = rblock[k] if k < len(rblock) else None
            if ltxt is not None and rtxt is not None:
                kind = "mod" if ltxt != rtxt else "equal"
            elif ltxt is not None:
                kind = "del"
            else:
                kind = "ins"
            rows.append({
                "ln": (i1 + k + 1) if ltxt is not None else None,
                "rn": (j1 + k + 1) if rtxt is not None else None,
                "lt": ltxt, "rt": rtxt, "kind": kind,
                "del": (ltxt is not None and rtxt is None),
            })
    return rows


def compress_rows(rows, context):
    """把大片未改动内容折叠成一行省略标记，避免图变成一堵墙。"""
    out = []
    i = 0
    n = len(rows)
    while i < n:
        if rows[i]["kind"] != "equal":
            out.append(rows[i])
            i += 1
            continue
        j = i
        while j < n and rows[j]["kind"] == "equal":
            j += 1
        run = rows[i:j]
        if len(run) <= context * 2 + 2:
            out.extend(run)
        else:
            out.extend(run[:context])
            out.append({"kind": "sep", "ln": None, "rn": None, "lt": None, "rt": None,
                        "count": len(run) - context * 2})
            out.extend(run[-context:])
        i = j
    return out


def render_text(left_path, right_path, out_path, title="", context=3, max_rows=140,
                font_size=14, max_width=1700):
    left = read_lines(left_path)
    right = read_lines(right_path)
    rows = build_rows(left, right)
    added = sum(1 for r in rows if r["kind"] == "ins")
    removed = sum(1 for r in rows if r["kind"] == "del")
    modified = sum(1 for r in rows if r["kind"] == "mod")
    rows = compress_rows(rows, context)

    fonts = load_fonts(font_size)
    lh = line_height(fonts) + 6
    pad = 8
    gut_w = 46

    ltxt = [r["lt"] for r in rows if r.get("lt") is not None]
    rtxt = [r["rt"] for r in rows if r.get("rt") is not None]
    lw = int(min(max((measure(t, fonts) for t in ltxt), default=100), max_width))
    rw = int(min(max((measure(t, fonts) for t in rtxt), default=100), max_width))
    panel_w = max(lw, rw)
    col_w = int(gut_w + panel_w + pad * 2)
    total_w = int(col_w * 2 + 1)

    head_h = 46
    foot_h = 26
    pages = [rows[i:i + max_rows] for i in range(0, len(rows), max_rows)] or [[]]

    written = []
    stem, ext = os.path.splitext(out_path)
    for pi, page in enumerate(pages):
        img = Image.new("RGB", (total_w, head_h + len(page) * lh + foot_h), C_BG)
        d = ImageDraw.Draw(img)

        d.rectangle([0, 0, total_w, head_h], fill=C_HEAD)
        lname = os.path.basename(left_path)
        rname = os.path.basename(right_path)
        draw_text(d, (pad, 8), "旧 · " + lname, fonts, C_HEAD_FG)
        draw_text(d, (col_w + pad, 8), "新 · " + rname, fonts, C_HEAD_FG)
        legend = "红 = 旧版被删/改    绿 = 新版新增/改    灰 = 相同"
        draw_text(d, (pad, 26), legend, fonts, (176, 186, 204))

        y = head_h
        for r in page:
            if r["kind"] == "sep":
                d.rectangle([0, y, total_w, y + lh], fill=C_SEP_BG)
                msg = "⋯ 省略 %d 行相同内容 ⋯" % r["count"]
                draw_text(d, (pad, y + 3), msg, fonts, C_SEP_FG)
                y += lh
                continue
            for side, col_x in ((0, 0), (1, col_w + 1)):
                if side == 0:
                    no, txt, kind = r["ln"], r["lt"], r["kind"]
                else:
                    no, txt, kind = r["rn"], r["rt"], r["kind"]
                if kind in ("del", "mod") and side == 0:
                    bg, fg, gbg = C_DEL_BG, C_DEL_FG, C_DEL_GUT
                elif kind in ("ins", "mod") and side == 1:
                    bg, fg, gbg = C_INS_BG, C_INS_FG, C_INS_GUT
                else:
                    bg, fg, gbg = C_BG, C_EQ_FG, C_GUT_BG
                d.rectangle([col_x, y, col_x + col_w, y + lh], fill=bg)
                d.rectangle([col_x, y, col_x + gut_w, y + lh], fill=gbg)
                d.rectangle([col_x, y, col_x + col_w, y + lh], outline=C_BORDER)
                if txt is None:
                    d.rectangle([col_x + gut_w, y, col_x + col_w, y + lh], outline=C_BORDER)
                    continue
                num = str(no) if no is not None else ""
                draw_text(d, (col_x + gut_w - 8 - measure(num, fonts), y + 3), num, fonts, C_GUT_FG)
                shown = txt
                if measure(shown, fonts) > panel_w:
                    while shown and measure(shown + "…", fonts) > panel_w:
                        shown = shown[:-1]
                    shown += "…"
                draw_text(d, (col_x + gut_w + pad, y + 3), shown, fonts, fg)
            y += lh

        foot = "%s差异行：新增 %d / 删除 %d / 修改 %d    渲染 %d 行%s" % (
            (title + "    ") if title else "",
            added, removed, modified, len(page),
            "（共 %d 页）" % len(pages) if len(pages) > 1 else "")
        draw_text(d, (pad, y + 4), foot, fonts, C_GUT_FG)

        target = out_path if len(pages) == 1 else "%s_p%d%s" % (stem, pi + 1, ext)
        img.save(target)
        written.append(target)

    return {"ok": True, "mode": "text", "files": written, "rows": len(rows),
            "added": added, "removed": removed, "modified": modified}


# ---------------------------------------------------------------- 位图对比

def load_rgb(path):
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    return im.convert("RGB")


def detect_regions(a, b, work_w=320, min_area_ratio=0.0006):
    """在缩小图上找差异连通块，返回相对坐标 (x0,y0,x1,y1) 列表。

    走 RGB 逐通道差分而不是灰度差分：同明度的换色（例如米色换成蓝色）
    灰度差只有个位数，转 RGB 后单通道差能到 100+，才不会漏检。
    """
    def shrink(im):
        w, h = im.size
        if w <= work_w:
            return im, 1.0
        nh = max(1, int(h * work_w / w))
        return im.resize((work_w, nh), Image.LANCZOS), w / float(work_w)

    sa, sx = shrink(a)
    sb, _ = shrink(b)
    if sa.size != sb.size:
        sb = sb.resize(sa.size, Image.LANCZOS)
    diff = ImageChops.difference(sa, sb)
    r, g, bl = diff.split()
    mask = ImageChops.lighter(ImageChops.lighter(r, g), bl)
    mask = mask.point(lambda v: 255 if v > 10 else 0)
    # 膨胀一圈：让"只有边缘不同"的换色块也能连成一个完整区域
    mask = mask.filter(ImageFilter.MaxFilter(5))
    mask = mask.filter(ImageFilter.GaussianBlur(0.6)).point(lambda v: 255 if v > 40 else 0)

    w, h = mask.size
    px = mask.load()
    seen = [[False] * w for _ in range(h)]
    boxes = []
    min_area = max(12, int(w * h * min_area_ratio))
    for y in range(h):
        for x in range(w):
            if px[x, y] != 255 or seen[y][x]:
                continue
            stack = [(x, y)]
            seen[y][x] = True
            x0, x1, y0, y1, area = x, x, y, y, 0
            while stack:
                cx, cy = stack.pop()
                area += 1
                x0, x1 = min(x0, cx), max(x1, cx)
                y0, y1 = min(y0, cy), max(y1, cy)
                for nx, ny in ((cx - 1, cy), (cx + 1, cy), (cx, cy - 1), (cx, cy + 1)):
                    if 0 <= nx < w and 0 <= ny < h and not seen[ny][nx] and px[nx, ny] == 255:
                        seen[ny][nx] = True
                        stack.append((nx, ny))
            if area >= min_area:
                # 小图是原图的 1/sx，坐标要乘回去才落在原图像素上
                boxes.append((x0 * sx, y0 * sx, x1 * sx, y1 * sx))
    return boxes


def draw_boxes(img, boxes, label_start=1):
    d = ImageDraw.Draw(img)
    fonts = load_fonts(13)
    for i, (x0, y0, x1, y1) in enumerate(boxes):
        x0 = max(0, int(x0)); y0 = max(0, int(y0))
        x1 = min(img.width - 1, int(x1)); y1 = min(img.height - 1, int(y1))
        if x1 <= x0 or y1 <= y0:
            continue
        w = max(2, img.width // 400)
        d.rectangle([x0, y0, x1, y1], outline=C_BOX, width=w)
        tag = str(i + label_start)
        tw = int(measure(tag, fonts)) + 8
        th = 18
        tx = min(max(0, x0), img.width - tw)
        ty = min(max(0, y0 - th), img.height - th)
        d.rectangle([tx, ty, tx + tw, ty + th], fill=C_BOX)
        draw_text(d, (tx + 4, ty + 2), tag, fonts, (255, 255, 255))
    return len(boxes)


def compose_images(parts, out_path, title="", foot="", max_h=1200, fonts=None, gap=24):
    """parts: [(label, PIL.Image), ...] 横向并排一张对比图。"""
    fonts = fonts or load_fonts(14)
    lh = line_height(fonts)
    head_h = 46
    foot_h = 28
    scaled = []
    target_h = min(max(im.height for _, im in parts), max_h)
    for label, im in parts:
        nh = target_h
        nw = max(1, int(im.width * nh / im.height))
        scaled.append((label, im.resize((nw, nh), Image.LANCZOS)))
    widths = [im.width for _, im in scaled]
    total_w = sum(widths) + gap * (len(scaled) - 1) + 24
    total_h = head_h + target_h + foot_h
    img = Image.new("RGB", (total_w, total_h), C_BG)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, total_w, head_h], fill=C_HEAD)
    x = 12
    for label, im in scaled:
        draw_text(d, (x, 8), label, fonts, C_HEAD_FG)
        draw_text(d, (x, 26), "%d×%d" % (im.width, im.height), fonts, (176, 186, 204))
        x += im.width + gap
    if title:
        draw_text(d, (total_w - measure(title, fonts) - 12, 26), title, fonts, (176, 186, 204))
    x = 12
    y = head_h
    for label, im in scaled:
        img.paste(im, (x, y))
        d.rectangle([x, y, x + im.width - 1, y + im.height - 1], outline=C_BORDER)
        x += im.width + gap
    d.rectangle([12, head_h + target_h // 2, total_w - 12, head_h + target_h // 2], fill=(0, 0, 0))
    if foot:
        draw_text(d, (12, head_h + target_h + 6), foot, fonts, C_GUT_FG)
    img.save(out_path)
    return out_path


def render_image(left_path, right_path, out_path, title="", mark=True, fonts=None):
    fonts = fonts or load_fonts(14)
    a = load_rgb(left_path)
    b = load_rgb(right_path)
    boxes = []
    if mark and a.size == b.size:
        boxes = detect_regions(a, b)
    if boxes:
        draw_boxes(a, boxes)
        draw_boxes(b, boxes)
    parts = [
        ("旧 · " + os.path.basename(left_path), a),
        ("新 · " + os.path.basename(right_path), b),
    ]
    if boxes:
        foot = "检出 %d 处变化区域（橙色框）" % len(boxes)
    elif a.size != b.size:
        foot = "尺寸不同：旧 %d×%d → 新 %d×%d（未做像素级比对）" % (a.width, a.height, b.width, b.height)
    else:
        foot = "未检出明显像素差异"
    compose_images(parts, out_path, title=title, foot=foot, fonts=fonts)
    return {"ok": True, "mode": "image", "files": [out_path], "changed_regions": len(boxes)}


# ---------------------------------------------------------------- 页面配对

def render_pair(left_paths, right_paths, out_path, title="", cell_h=520, fonts=None):
    fonts = fonts or load_fonts(14)
    lh = line_height(fonts)
    head_h = 46
    foot_h = 28
    count = max(len(left_paths), len(right_paths))
    cells = []
    for i in range(count):
        lp = left_paths[i] if i < len(left_paths) else None
        rp = right_paths[i] if i < len(right_paths) else None
        li = load_rgb(lp) if lp else None
        ri = load_rgb(rp) if rp else None
        cells.append((li, ri))
    colw = [0, 0]
    for li, ri in cells:
        for k, im in enumerate((li, ri)):
            if im is None:
                continue
            nw = max(1, int(im.width * cell_h / im.height))
            colw[k] = max(colw[k], nw)
    gap = 20
    total_w = colw[0] + colw[1] + gap * 3
    total_h = head_h + count * (cell_h + 22) + foot_h
    img = Image.new("RGB", (total_w, total_h), C_BG)
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, total_w, head_h], fill=C_HEAD)
    draw_text(d, (12, 8), "旧版页面", fonts, C_HEAD_FG)
    draw_text(d, (12 + colw[0] + gap, 8), "新版页面", fonts, C_HEAD_FG)
    if title:
        draw_text(d, (total_w - measure(title, fonts) - 12, 8), title, fonts, C_HEAD_FG)
    y = head_h
    for i, (li, ri) in enumerate(cells):
        x = 12
        for k, im in enumerate((li, ri)):
            if im is not None:
                nw = max(1, int(im.width * cell_h / im.height))
                nh = cell_h
                im2 = im.resize((nw, nh), Image.LANCZOS)
                img.paste(im2, (x, y))
                d.rectangle([x, y, x + nw - 1, y + nh - 1], outline=C_BORDER)
            x += colw[k] + gap
        draw_text(d, (12, y + cell_h + 2), "P%d" % (i + 1), fonts, C_GUT_FG)
        y += cell_h + 22
    draw_text(d, (12, y + 4), "共 %d 页" % count, fonts, C_GUT_FG)
    img.save(out_path)
    return {"ok": True, "mode": "pair", "files": [out_path], "pages": count}


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(description="渲染前后对比图")
    sub = ap.add_subparsers(dest="mode", required=True)

    t = sub.add_parser("text", help="文本/代码并排 diff")
    t.add_argument("--left", required=True)
    t.add_argument("--right", required=True)
    t.add_argument("--out", required=True)
    t.add_argument("--title", default="")
    t.add_argument("--context", type=int, default=3)
    t.add_argument("--max-rows", type=int, default=140)
    t.add_argument("--font-size", type=int, default=14)
    t.add_argument("--json-out", default="")

    m = sub.add_parser("image", help="位图并排 + 变化区域")
    m.add_argument("--left", required=True)
    m.add_argument("--right", required=True)
    m.add_argument("--out", required=True)
    m.add_argument("--title", default="")
    m.add_argument("--no-mark", action="store_true")
    m.add_argument("--json-out", default="")

    p = sub.add_parser("pair", help="页面/幻灯片逐页并排")
    p.add_argument("--left", nargs="+", required=True)
    p.add_argument("--right", nargs="+", required=True)
    p.add_argument("--out", required=True)
    p.add_argument("--title", default="")
    p.add_argument("--cell-height", type=int, default=520)
    p.add_argument("--json-out", default="")

    args = ap.parse_args()
    try:
        if args.mode == "text":
            res = render_text(args.left, args.right, args.out, args.title,
                              args.context, args.max_rows, args.font_size)
        elif args.mode == "image":
            res = render_image(args.left, args.right, args.out, args.title, not args.no_mark)
        else:
            res = render_pair(args.left, args.right, args.out, args.title, args.cell_height)
    except Exception as exc:  # noqa: BLE001
        res = {"ok": False, "error": "%s: %s" % (type(exc).__name__, exc),
               "at": datetime.now().isoformat(timespec="seconds")}
        print(json.dumps(res, ensure_ascii=False, indent=2))
        sys.exit(1)
    if getattr(args, "json_out", ""):
        with open(args.json_out, "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=2)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
