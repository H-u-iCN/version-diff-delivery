#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""接触印相表拼图公共件：把若干页 PNG 拼成一张大图，或把新旧两版逐页并排。

被 render_deck.py（PPTX）与 render_doc.py（DOCX）共用。
拼图每页带 p01/p02 标签，手机点开大图即可缩放看细节。
"""

import os
import sys

from PIL import Image, ImageDraw, ImageFont

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from render_compare import detect_regions, load_rgb
except Exception:  # pragma: no cover
    detect_regions = None
    load_rgb = None

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:  # pragma: no cover
    pass

FONTS = {
    'cn': r'C:\Windows\Fonts\msyh.ttc',
    'cn_bold': r'C:\Windows\Fonts\msyhbd.ttc',
    'doc': r'C:\Windows\Fonts\simsun.ttc',
    'doc_bold': r'C:\Windows\Fonts\msyhbd.ttc',
    'en': r'C:\Windows\Fonts\georgia.ttf',
    'en_bold': r'C:\Windows\Fonts\georgiab.ttf',
}
FALLBACK_FONTS = [r'C:\Windows\Fonts\msyh.ttc', r'C:\Windows\Fonts\simhei.ttf',
                  r'C:\Windows\Fonts\simsun.ttc',
                  '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc']
_fcache = {}

C_BG = (247, 248, 250)
C_CARD = (255, 255, 255)
C_BORDER = (214, 218, 226)
C_LABEL = (108, 116, 132)
C_HEAD = (30, 36, 48)
C_BOX = (255, 138, 0)


def font(kind, size_px):
    key = (kind, int(size_px))
    if key not in _fcache:
        path = FONTS.get(kind)
        if not path or not os.path.exists(path):
            path = next((p for p in FALLBACK_FONTS if os.path.exists(p)), None)
        if path is None:
            _fcache[key] = ImageFont.load_default()
        else:
            try:
                _fcache[key] = ImageFont.truetype(path, max(6, int(size_px)))
            except Exception:
                _fcache[key] = ImageFont.load_default()
    return _fcache[key]


def load_img(path):
    if load_rgb:
        return load_rgb(path)
    im = Image.open(path)
    if im.mode in ('RGBA', 'LA', 'P'):
        im = im.convert('RGBA')
        bg = Image.new('RGBA', im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    return im.convert('RGB')


def build_sheet(page_paths, out_path, cols=4, thumb_w=440, title="", gap=14, pad=18):
    """把若干页 PNG 拼成一张接触印相表。"""
    imgs = [load_img(p) for p in page_paths]
    if not imgs:
        raise ValueError('没有可拼的页面')
    thumb_h = int(thumb_w * imgs[0].height / imgs[0].width)
    label_h = 22
    cell_h = thumb_h + label_h
    rows = (len(imgs) + cols - 1) // cols
    head_h = 52 if title else 0
    W = cols * thumb_w + (cols + 1) * gap + pad
    H = head_h + rows * cell_h + (rows + 1) * gap + pad
    sheet = Image.new('RGB', (W, H), C_BG)
    d = ImageDraw.Draw(sheet)
    if title:
        d.rectangle([0, 0, W, head_h], fill=C_HEAD)
        d.text((pad, 16), title, font=font('cn_bold', 20), fill=(255, 255, 255))
    for i, im in enumerate(imgs):
        r, c = divmod(i, cols)
        x = pad + gap + c * (thumb_w + gap)
        y = head_h + pad + gap + r * (cell_h + gap)
        th = im.resize((thumb_w, thumb_h), Image.LANCZOS)
        d.rectangle([x - 3, y - 3, x + thumb_w + 2, y + thumb_h + 2], fill=C_CARD)
        sheet.paste(th, (x, y))
        d.rectangle([x, y, x + thumb_w - 1, y + thumb_h - 1], outline=C_BORDER)
        d.text((x, y + thumb_h + 3), 'p%02d' % (i + 1), font=font('cn', 15), fill=C_LABEL)
    sheet.save(out_path)
    return out_path, (W, H)


def build_compare(left_paths, right_paths, out_path, thumb_w=440, title="", gap=14, pad=18,
                  mark_changes=True, page_label="p%02d"):
    """新旧两版逐页并排；变化处在新版一侧打橙框并编号。"""
    n = max(len(left_paths), len(right_paths))
    lf = font('cn', 15)
    lf_b = font('cn_bold', 15)
    head_h = 78 if title else 0
    probe = load_img(left_paths[0] if left_paths else right_paths[0])
    thumb_h = int(thumb_w * probe.height / probe.width)
    label_h = 24
    cell_h = thumb_h + label_h
    W = 2 * thumb_w + 3 * gap + pad
    H = head_h + n * cell_h + (n + 1) * gap + pad
    sheet = Image.new('RGB', (W, H), C_BG)
    d = ImageDraw.Draw(sheet)
    if title:
        d.rectangle([0, 0, W, head_h], fill=C_HEAD)
        d.text((pad, 14), title, font=font('cn_bold', 20), fill=(255, 255, 255))
        d.text((pad, 42), '旧版', font=lf_b, fill=(176, 186, 204))
        d.text((pad + thumb_w + gap, 42), '新版', font=lf_b, fill=(176, 186, 204))

    changed = []
    for i in range(n):
        x0 = pad + gap
        x1 = pad + 2 * gap + thumb_w
        y = head_h + pad + gap + i * (cell_h + gap)
        lp = left_paths[i] if i < len(left_paths) else None
        rp = right_paths[i] if i < len(right_paths) else None
        boxes = []
        if lp and rp and mark_changes and detect_regions:
            try:
                boxes = detect_regions(load_img(lp), load_img(rp))
            except Exception:
                boxes = []
        for slot, path in ((0, lp), (1, rp)):
            if not path:
                x = x0 if slot == 0 else x1
                d.rectangle([x, y, x + thumb_w, y + thumb_h], fill=(235, 237, 241),
                            outline=C_BORDER)
                d.text((x + 12, y + 12), '（该版本无此页）', font=lf, fill=C_LABEL)
                continue
            im = load_img(path)
            th = im.resize((thumb_w, thumb_h), Image.LANCZOS)
            x = x0 if slot == 0 else x1
            d.rectangle([x - 3, y - 3, x + thumb_w + 2, y + thumb_h + 2], fill=C_CARD)
            sheet.paste(th, (x, y))
            d.rectangle([x, y, x + thumb_w - 1, y + thumb_h - 1], outline=C_BORDER)
            if slot == 1 and boxes:
                sx = thumb_w / im.width
                sy = thumb_h / im.height
                for k, (bx0, by0, bx1, by1) in enumerate(boxes):
                    d.rectangle([x + bx0 * sx, y + by0 * sy, x + bx1 * sx, y + by1 * sy],
                                outline=C_BOX, width=3)
                    d.text((x + bx0 * sx + 2, y + by0 * sy + 1), str(k + 1), font=lf_b, fill=C_BOX)
        d.text((x0, y + thumb_h + 4), page_label % (i + 1), font=lf, fill=C_LABEL)
        if boxes:
            changed.append(i + 1)
    sheet.save(out_path)
    return out_path, (W, H), changed
