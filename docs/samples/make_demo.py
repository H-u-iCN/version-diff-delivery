# -*- coding: utf-8 -*-
"""生成演示素材：合成一份 PPTX(两版) + DOCX(两版) + 一张位图(两版) + 一段代码(两版)。

只用来给 README 出效果图，内容为虚构演示数据，不含任何真实文档内容。
"""
import os
import sys

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from docx import Document
from docx.shared import Pt as DPt, RGBColor as DRGB

OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)

NAVY = RGBColor(0x1B, 0x2A, 0x3A)
RED = RGBColor(0x9E, 0x2A, 0x25)
CREAM = RGBColor(0xFD, 0xF6, 0xE6)
GREY = RGBColor(0x3A, 0x3A, 0x3A)
LIGHT = RGBColor(0xF2, 0xEC, 0xDD)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

W, H = Inches(13.333), Inches(7.5)


def blank(prs, bg):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    r = s.shapes.add_shape(1, 0, 0, W, H)  # rectangle
    r.fill.solid()
    r.fill.fore_color.rgb = bg
    r.line.fill.background()
    r.shadow.inherit = False
    return s


def text(slide, x, y, w, h, s, size, color, bold=False, align=PP_ALIGN.LEFT, font='Microsoft YaHei'):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = s
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font
    return tb


def rect(slide, x, y, w, h, color, shape=1):
    r = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid()
    r.fill.fore_color.rgb = color
    r.line.fill.background()
    r.shadow.inherit = False
    return r


def build_pptx(path, version):
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    title = '版本对比交付 Skill' if version == 'v1' else '版本对比交付 Skill（修订版）'
    tagline = '改文件 · 留原版 · 出对比图' if version == 'v1' else '改文件 · 留原版 · 出对比图 · 支持 PPT 与 Word'

    # 1 封面
    s = blank(prs, NAVY)
    rect(s, 8.9, 0, 4.45, 7.5, RED)
    rect(s, 8.86, 0, 0.05, 7.5, CREAM)
    text(s, 0.9, 2.1, 7.6, 1.1, title, 40, WHITE, bold=True)
    text(s, 0.95, 3.25, 7.2, 0.6, tagline, 18, RGBColor(0xC9, 0xD4, 0xE0))
    text(s, 0.95, 4.35, 7.2, 0.5, 'MiniMax Code / Mavis  ·  version-diff-delivery', 13,
         RGBColor(0x8E, 0x9E, 0xB0))

    # 2 目录 + 表格
    s = blank(prs, CREAM)
    text(s, 0.7, 0.5, 6.0, 0.5, '目录  CONTENTS', 15, RED, bold=True)
    text(s, 0.7, 1.0, 6.0, 0.7, '它到底解决什么问题', 32, NAVY, bold=True)
    items = ['保留原版，改动另存 _V1 / _V2', '代码改动 → 左右并排 diff 图',
             '图片改动 → 并排图 + 变更区域框', 'PPT / Word → 整份拼成一张接触印相表']
    for i, it in enumerate(items):
        y = 2.05 + i * 0.72
        rect(s, 0.72, y + 0.06, 0.22, 0.22, RED)
        text(s, 1.12, y - 0.04, 5.6, 0.45, it, 15, GREY)
    rows, cols = 5, 3
    tb = s.shapes.add_table(rows, cols, Inches(7.1), Inches(1.1), Inches(5.5), Inches(3.1)).table
    data = [['版本', '改动内容', '交付形态'],
            ['V0', '原始文件', '不覆盖'],
            ['V1', '标题文案', '并排 diff 图'],
            ['V2', '新增常量', '并排 diff 图'],
            ['V3', '图片换色', '橙框标注图']]
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = tb.cell(ri, ci)
            cell.text = val
            p = cell.text_frame.paragraphs[0]
            p.runs[0].font.size = Pt(12)
            p.runs[0].font.bold = (ri == 0)
            p.runs[0].font.color.rgb = WHITE if ri == 0 else GREY
            p.runs[0].font.name = 'Microsoft YaHei'
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY if ri == 0 else (WHITE if ri % 2 else LIGHT)

    # 3 时间线
    s = blank(prs, CREAM)
    text(s, 0.7, 0.5, 6.0, 0.5, '演进  ROADMAP', 15, RED, bold=True)
    text(s, 0.7, 1.0, 9.0, 0.7, '每轮迭代都留痕，都可回看', 32, NAVY, bold=True)
    rect(s, 0.9, 3.35, 11.4, 0.06, RED)
    years = ['2024', '2025', '2026 Q1', '2026 Q2', '现在']
    labels = ['第一版流程', '补图例', '加 PPT', '加 Word', '公开可用']
    for i, (yr, lb) in enumerate(zip(years, labels)):
        x = 1.15 + i * 2.35
        rect(s, x, 3.14, 0.42, 0.42, NAVY, shape=9)  # oval
        text(s, x - 0.35, 2.45, 1.2, 0.4, yr, 17, RED, bold=True, align=PP_ALIGN.CENTER)
        text(s, x - 0.5, 3.85, 1.5, 0.4, lb, 13, GREY, align=PP_ALIGN.CENTER)

    # 4 双栏
    s = blank(prs, CREAM)
    text(s, 0.7, 0.5, 9.0, 0.5, '机制  HOW IT WORKS', 15, RED, bold=True)
    text(s, 0.7, 1.0, 9.0, 0.7, '三步走完一版迭代', 32, NAVY, bold=True)
    rect(s, 0.7, 2.1, 5.9, 3.6, NAVY)
    rect(s, 6.95, 2.1, 5.9, 3.6, WHITE)
    for i, t in enumerate(['查基线：脚本判定，不手算版本号', '开新版本：原版字节不动', '渲染对比：图 + 汇总表']):
        rect(s, 1.05, 2.55 + i * 0.95, 0.3, 0.3, RED, shape=9)
        text(s, 1.55, 2.5 + i * 0.95, 4.9, 0.5, t, 15, WHITE)
    for i, t in enumerate(['原版永远保留在原处', '新版命名 _V1 / _V2 递增', '对比图进 .mavis-compare/']):
        text(s, 7.35, 2.55 + i * 0.95, 5.1, 0.5, '· ' + t, 15, GREY)

    # 5 图文混排
    s = blank(prs, NAVY)
    img = Image.new('RGB', (900, 1200), (28, 42, 58))
    dr = ImageDraw.Draw(img)
    for i in range(14):
        dr.rectangle([60 + i * 40, 200 + (i % 5) * 60, 120 + i * 40, 900],
                     fill=(40 + i * 9, 70 + i * 6, 110 + i * 4))
    dr.ellipse([260, 420, 640, 800], fill=(158, 42, 37))
    im = Image.new('RGB', (640, 900), (28, 42, 58))
    im.paste(img.crop((130, 150, 770, 1050)).resize((640, 900)), (0, 0))
    tmp = os.path.join(OUT, 'demo_photo.png')
    im.save(tmp)
    s.shapes.add_picture(tmp, Inches(8.4), 0, height=H)
    text(s, 0.9, 1.5, 6.6, 1.8, '03', 72, RED, bold=True)
    text(s, 0.95, 3.0, 6.8, 0.6, '效果一览', 30, WHITE, bold=True)
    text(s, 0.95, 3.9, 6.8, 1.4,
         '演示页：图与文同屏\n用于验证图文混排页的还原度', 15, RGBColor(0xC9, 0xD4, 0xE0))

    # 6 收尾
    s = blank(prs, NAVY)
    text(s, 0.9, 0.75, 8.0, 0.8, '小结  SUMMARY', 15, RED, bold=True)
    text(s, 0.9, 1.2, 9.0, 0.7, '四件被记住的事', 32, WHITE, bold=True)
    for i, (a, b) in enumerate([('原版不动', 'V0 永远是原始状态'),
                                 ('版本递增', '_V1 / _V2 不覆盖'),
                                 ('图先于话', '对比图 + 汇总表'),
                                 ('可复核', '哈希与页码随交付给出')]):
        y = 2.35 + i * 0.95
        text(s, 0.95, y, 3.0, 0.5, a, 19, RED, bold=True)
        text(s, 4.0, y + 0.05, 7.0, 0.5, b, 15, RGBColor(0xC9, 0xD4, 0xE0))
    prs.save(path)
    return path


def build_docx(path, version):
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = 'Microsoft YaHei'
    st.font.size = DPt(10.5)

    def para(s, size=10.5, bold=False, color=None, align=None, space_after=4):
        p = doc.add_paragraph()
        r = p.add_run(s)
        r.font.size = DPt(size)
        r.font.bold = bold
        if color:
            r.font.color.rgb = DRGB(*color)
        if align is not None:
            p.alignment = align
        p.paragraph_format.space_after = DPt(space_after)
        return p

    para('版本对比交付 Skill', 22, True, (0x1B, 0x2A, 0x3A), 1, 6)
    para('—— 演示文档，用于说明每一版如何被留痕与对比', 11, False, (0x8E, 0x9E, 0xB0), 1, 14)
    if version == 'v1':
        para('一、它解决什么问题', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    else:
        para('一、它解决什么问题（修订版）', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    para('日常改文件最怕两件事：一是把原版改没了，二是说"改了什么"全靠嘴。这里把两件事都变成可复核的产物：'
         '原版原地保留，新版按 _V1、_V2 递增留档，每轮末尾给出可视化对比与版本链表格。', 10.5, False, None,
         None, 8)
    for t in ['原版是 V0，字节不动，任何时候都能回到起点',
              '改动落在新版本文件上，命名规则固定，不靠记忆',
              '对比图与汇总表同时给出，图看位置，表看清单']:
        para('• ' + t, 10.5, False, None, None, 3)

    para('二、版本链长什么样', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    tb = doc.add_table(rows=4, cols=4)
    tb.style = 'Table Grid'
    for ci, v in enumerate(['版本', '文件', '改动', '对比形态']):
        c = tb.cell(0, ci)
        c.text = v
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = DPt(9.5)
    rows = [('V0', 'config.ts', '原始状态', '—'),
            ('V1', 'config_V1.ts', '标题文案', '并排 diff 图'),
            ('V2', 'config_V2.ts', '新增常量', '并排 diff 图')]
    for ri, row in enumerate(rows, start=1):
        for ci, v in enumerate(row):
            c = tb.cell(ri, ci)
            c.text = v
            c.paragraphs[0].runs[0].font.size = DPt(9.5)

    para('三、每轮回复固定给三样东西', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    para('第一样是版本链表格，列出文件、版本与本轮改动；第二样是每个改动文件一张对比图，'
         '图上能直接看出改了哪儿；第三样是改动摘要，图里看不出来的细节写在这里。', 10.5)
    para('四、适用与不适用', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    para('适用于交付型文件：代码、配置、图片、PPT、Word、表格、视频。临时脚本与缓存不参与版本化。', 10.5)
    para('不适用于：只读的排查、纯问答，以及用户明确要求原地覆盖的场景——那种情况直接照做，不额外留版。', 10.5)
    para('五、每种文件类型怎么比', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    tb2 = doc.add_table(rows=5, cols=3)
    tb2.style = 'Table Grid'
    for ci, v in enumerate(['文件类型', '主形式', '附带形式']):
        c = tb2.cell(0, ci)
        c.text = v
        c.paragraphs[0].runs[0].font.bold = True
        c.paragraphs[0].runs[0].font.size = DPt(9.5)
    rows2 = [('代码 / 配置', '左右并排 diff 图', 'unified diff 代码块'),
             ('图片', '并排图 + 变更区域框', '尺寸与色值参数表'),
             ('PPT / Word', '整份接触印相表', '新旧逐页并排 + 橙框'),
             ('表格', '单元格 before/after', 'CSV 文本 diff')]
    for ri, row in enumerate(rows2, start=1):
        for ci, v in enumerate(row):
            c = tb2.cell(ri, ci)
            c.text = v
            c.paragraphs[0].runs[0].font.size = DPt(9.5)
    para('六、依赖与安装', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    para('装完下一个会话生效，不依赖运行时重启。依赖只有 Python 3 与 Pillow；', 10.5)
    para('PPTX 路线另需 python-pptx，DOCX 路线另需 python-docx。缺哪个，对应路线会明确报错并给兜底形式。', 10.5)
    para('七、怎么复现上面的图', 15, True, (0x9E, 0x2A, 0x25), None, 6)
    para('仓库里的 docs/ 目录放了全部效果图，它们由本 skill 自己的脚本渲染合成样例生成：', 10.5)
    for t in ['render_compare.py text / image —— 代码与位图对比',
              'render_deck.py sheet / compare —— PPT 接触印相表与逐页对比',
              'render_doc.py sheet / compare —— Word 接触印相表与逐页对比']:
        para('• ' + t, 10.5, False, None, None, 3)
    para('汇报人：千晖　　2026 年 10 月', 10.5, False, None, 2, 10)
    doc.add_page_break()
    para('附录  术语与命名约定', 15, True, (0x9E, 0x2A, 0x25), None, 8)
    para('V0：原版。任何时候都指向最初的字节状态。', 10.5)
    para('V1 / V2 / V3：按轮次递增的新版本，命名规则为「文件名_V<N>.扩展名」，重名自动跳号。', 10.5)
    para('基线：做本轮对比时对照的那一版，默认是上一轮的最新版本。', 10.5)
    para('接触印相表：把整份文档或整份 PPT 的每一页缩略图拼成一张大图，每页带 p01、p02 标签，'
         '手机点开即可缩放查看细节。', 10.5)
    para('changed_pages：对比图里被判定为有变化的页码列表，是核对用的硬指标。', 10.5)
    para('汇报人：千晖　　2026 年 10 月', 10.5, False, None, 2, 10)
    doc.save(path)
    return path


def build_cover(path, color):
    im = Image.new('RGB', (1000, 640), (250, 246, 238))
    dr = ImageDraw.Draw(im)
    dr.rectangle([90, 90, 470, 380], fill=color)
    dr.rectangle([560, 120, 910, 250], fill=(60, 62, 70))
    dr.rectangle([560, 300, 910, 470], fill=(225, 218, 200))
    dr.ellipse([120, 430, 440, 570], fill=color)
    im.save(path)
    return path


if __name__ == '__main__':
    build_pptx(os.path.join(OUT, 'demo_v1.pptx'), 'v1')
    build_pptx(os.path.join(OUT, 'demo_v2.pptx'), 'v2')
    build_docx(os.path.join(OUT, 'demo_v1.docx'), 'v1')
    build_docx(os.path.join(OUT, 'demo_v2.docx'), 'v2')
    build_cover(os.path.join(OUT, 'cover_v1.png'), (196, 154, 92))
    build_cover(os.path.join(OUT, 'cover_v2.png'), (37, 99, 235))
    with open(os.path.join(OUT, 'demo_v1.ts'), 'w', encoding='utf-8') as f:
        f.write("// 站点配置\nTITLE = '春季上新'\nSUBTITLE = '副标题'\n\nNAV = ['a', 'b']\n")
    with open(os.path.join(OUT, 'demo_v2.ts'), 'w', encoding='utf-8') as f:
        f.write("// 站点配置\nTITLE = '春季上新·第二版'\nSUBTITLE = '副标题'\nSITE_NAME = '春季上新站'\n\nNAV = ['a', 'b', 'c']\n")
    print('demo assets ready in', OUT)
