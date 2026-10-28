# -*- coding: utf-8 -*-
"""用本 skill 自己的脚本，把 docs/samples/ 里的样例重新渲染成 docs/ 下的 6 张效果图。

用法（在仓库根目录执行）：
    python docs/samples/render_all.py

依赖：Pillow、python-pptx、python-docx。渲染全部走 ../scripts/ 下的脚本，
不额外引入任何依赖。
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..', '..'))
SAMPLES = HERE
DOCS = os.path.join(ROOT, 'docs')
SCRIPTS = os.path.join(ROOT, 'scripts')
PY = sys.executable


def run(args):
    print('>', ' '.join(os.path.basename(a) for a in args[:2]), '...')
    subprocess.check_call([PY] + args)


def main():
    os.makedirs(DOCS, exist_ok=True)
    work = os.path.join(SAMPLES, '_work')
    os.makedirs(work, exist_ok=True)

    run([os.path.join(SCRIPTS, 'render_compare.py'), 'text',
         '--left', os.path.join(SAMPLES, 'demo_v1.ts'),
         '--right', os.path.join(SAMPLES, 'demo_v2.ts'),
         '--out', os.path.join(DOCS, 'demo-code-diff.png'),
         '--title', '代码改动 → 左右并排 diff（红删绿增）'])

    run([os.path.join(SCRIPTS, 'render_compare.py'), 'image',
         '--left', os.path.join(SAMPLES, 'cover_v1.png'),
         '--right', os.path.join(SAMPLES, 'cover_v2.png'),
         '--out', os.path.join(DOCS, 'demo-image-compare.png'),
         '--title', '图片改动 → 并排图 + 变更区域框'])

    run([os.path.join(SCRIPTS, 'render_deck.py'), 'sheet',
         '--pptx', os.path.join(SAMPLES, 'demo_v1.pptx'),
         '--out', os.path.join(DOCS, 'demo-pptx-sheet.png'),
         '--cols', '3', '--thumb-w', '520',
         '--title', 'PPT 整份拼成一张接触印相表',
         '--workdir', os.path.join(work, 'p1')])

    run([os.path.join(SCRIPTS, 'render_deck.py'), 'compare',
         '--left', os.path.join(SAMPLES, 'demo_v1.pptx'),
         '--right', os.path.join(SAMPLES, 'demo_v2.pptx'),
         '--out', os.path.join(DOCS, 'demo-pptx-compare.png'),
         '--thumb-w', '520',
         '--title', 'PPT 新旧逐页并排，变化处打橙框',
         '--workdir', os.path.join(work, 'p2')])

    run([os.path.join(SCRIPTS, 'render_doc.py'), 'sheet',
         '--docx', os.path.join(SAMPLES, 'demo_v1.docx'),
         '--out', os.path.join(DOCS, 'demo-docx-sheet.png'),
         '--cols', '3', '--thumb-w', '520',
         '--title', 'Word 整份拼成一张接触印相表',
         '--workdir', os.path.join(work, 'd1')])

    run([os.path.join(SCRIPTS, 'render_doc.py'), 'compare',
         '--left', os.path.join(SAMPLES, 'demo_v1.docx'),
         '--right', os.path.join(SAMPLES, 'demo_v2.docx'),
         '--out', os.path.join(DOCS, 'demo-docx-compare.png'),
         '--thumb-w', '520',
         '--title', 'Word 新旧逐页并排，变化处打橙框',
         '--workdir', os.path.join(work, 'd2')])

    print('done -> %s' % DOCS)


if __name__ == '__main__':
    main()
