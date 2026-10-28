---
name: version-diff-delivery
description: |
  版本化迭代 + 前后对比交付：改任何文件（代码 / 配置 / 图片 / PPT / 文档 / 表格 / 视频）都保留原版，
  另写 _V1、_V2、_V3 递增版本；每轮回复结尾给出旧版与新版的可视化对比（并排图 / 差异高亮图 /
  页面对）+ 一张版本链汇总表。
  Use this skill whenever the agent edits, rewrites, regenerates, or generates a file and then reports
  the result. Triggers on: 改完了看看差异, 前后对比, 对比图, 保留原版, 出个 V1 版本, 改了哪些地方,
  diff 一下, 版本迭代, 迭代版本, before/after comparison, show me what changed.
  Do NOT use for read-only inspection, pure Q&A, or replies that change no file.
descriptions:
  zh-Hans: "版本化迭代 + 前后对比交付：改文件必留原版并出 _V1/_V2 版本，每轮结尾给可视化前后对比图和汇总表。"
displayNames:
  zh-Hans: "版本对比交付"
---

# 版本对比交付

每一轮改完文件，都要把「改了什么」以看得见的形式交给用户：原版留着不动，新版另存为递增版本，
回复结尾给对比图 + 版本链表格。

## 安装与依赖

```bash
git clone https://github.com/H-u-iCN/version-diff-delivery.git ~/.minimax/skills/version-diff-delivery
```

Windows（PowerShell）：

```powershell
git clone https://github.com/H-u-iCN/version-diff-delivery.git "$env:USERPROFILE\.minimax\skills\version-diff-delivery"
```

装完**下一个会话生效**，不需要重启 runtime。

依赖：Python 3 + Pillow（必需）；PPTX 路线另需 python-pptx，DOCX 路线另需 python-docx。
缺哪个，对应路线会明确报错并给兜底形式，不会静默跳过。

## Inputs to collect

- **交付文件清单**：只有用户要拿走的东西才版本化。临时脚本、缓存、日志、构建产物不参与。
- **基线是谁**：上一轮的最新版本（`_V1`、`_V2`…）；没有历史版本时用原文件本身。原文件永远是 V0。
- **对比基准**：默认跟「上一版」比；用户想看累计改动时，跟原版（V0）比，标题里写清楚。
- **用户是否明确要求原地覆盖**：明确说了就直接改，不版本化——用户的明确指令优先于本 skill 的默认。
- **这个文件是不是运行中就生效的文件**：站点配置、线上素材、构建入口、脚本等；决定收尾要不要交代
  「新版本如何顶替生效」。

## Procedure

1. **先确认基线和下一个版本名**，别手算版本号：
   ```powershell
   python "<skill_dir>/scripts/version_file.py" chain --file "<基线文件路径>"
   ```
   原因：版本号跳号、重名覆盖是这个流程最容易出错的地方，脚本的判定是确定性的。
   输出里 `base_path` 是基线，`next_path` 是本轮要写的文件。

2. **开新版本工作副本**（原版和所有旧版本不动）：
   ```powershell
   python "<skill_dir>/scripts/version_file.py" snap --file "<基线文件路径>" --label "<本轮改了什么>"
   ```
   原因：`snap` 只做一次拷贝，原版因此天然被保留；直接编辑原版会让"原始状态"丢失。
   之后的编辑只在新版本文件上做。若目标已存在，脚本会自动跳到下一个空号。

3. **在新版本上完成改动**。改完顺手确认原版字节数/内容没变——用户要的是"原样还在"。

4. **交代生效路径**：新版本默认**不自动顶替**原文件。若该文件是运行中就生效的东西（站点配置、线上素材、
   构建入口、脚本），收尾必须写清「新版本在哪个路径、要不要复制回原路径才生效」；用户明确要即时生效，
   就按其指示复制回去，并说明原版仍在、可回溯。
   原因：用户要的正是原样保留，自动顶替等于把 V0 悄悄换成 V1；而不交代生效路径，用户会误以为改动已经跑起来了。

5. **渲染对比图**，输出到 `<项目根>/.mavis-compare/`（项目根 = git 仓库根 `git rev-parse --show-toplevel`，
   没有仓库就用当前工作目录）。文件名：`<名称>_<from>vs<to>_<YYYYMMDD-HHMM>.png`。
   按文件类型选模式，各类型的具体配方见 `references/comparison-formats.md`：
   - 代码 / 配置 / 文本 / Markdown / JSON / CSV → `render_compare.py text`
   - 位图 → `render_compare.py image`
   - PPTX / 演示文稿 → `render_deck.py compare`（逐页并排，变化页自动打橙框标出改了哪儿）
   - DOCX / Word → `render_doc.py compare`（逐页并排）+ `render_doc.py sheet`（全文拼一张接触印相表）
   - 单份 PPT / Word 概览（只展示不比对）→ `render_deck.py sheet` / `render_doc.py sheet`
     （全部页面拼成一张接触印相表，每页带 p01/p02 标签，手机点开就能放大看）
   - 已渲染好的页面 / 幻灯片 PNG → `render_compare.py pair`
   ```powershell
   python "<skill_dir>/scripts/render_compare.py" text --left "<旧文件>" --right "<新文件>" `
     --out "<项目根>/.mavis-compare/app_V0vsV1_20261027-1340.png" --title "app.ts V0 -> V1"
   ```
   原因：图比文字更快看出"哪里变了"，用户要的就是这个而不是一段复述。

6. **收尾三段式**（顺序固定，这是本 skill 的交付契约）：
   - **版本链表格**：文件 | 版本 | 本轮改动
   - **每个改动文件一张对比图**：用 `<media src="绝对路径" />` 发出，发之前确认文件真实存在
   - **改动摘要**：每个文件 1–3 句，说清改了什么、为什么改；图里看不出来的细节（数值、参数、
     取舍理由）写在这里

## Output contract

每轮回复结尾固定给出：

```markdown
## 本轮改动对比

| 文件 | 版本 | 改动摘要 |
|---|---|---|
| site.config.ts | V0 → V1 | 标题文案更新，NAV 增补 'c'，新增 FOOTER 常量 |

<media src="C:/.../.mavis-compare/site.config_V0vsV1_20261027-1340.png" />

- **site.config.ts**（V1）：`TITLE` 由「旧标题」改为「新标题·第二版」；`NAV` 增加 `'c'`；新增 `FOOTER` 常量。
  原版 `site.config.ts` 未改动。
```

- 图默认每个改动文件一张，汇总表一张；图过多时按文件分组，不要用一张大图糊弄多个文件。
- 对比形式可扩展：图之外可叠加 unified diff 代码块、字段级 before/after 表格、参数差异表、
  关键帧并排——见 `references/comparison-formats.md`。
- 版本链永远完整可追溯：V0（原版）→ V1 → V2 …，任何一版都不会被覆盖或删除。

## Failure handling

- **Pillow 缺失**：脚本会返回 `ok: false`。退回到 unified diff 代码块 + 字段级 before/after 表格，
  并在回复里说明"对比图未能生成（缺少 Pillow）"，不要静默跳过对比这一环。
- **差异行数过多**：`render_compare.py text` 会自动分页成 `_p1`、`_p2`…多张图，全部交付，不要只给第一张。
- **中文字形缺失**：脚本自动在 `msyh.ttc` / `simhei.ttf` / 思源黑体之间探测；若系统一个中文字体都没有，
  中文会显示成方块——此时提示安装字体，而不是改用会丢中文的纯 ASCII 字体。
- **DOCX / PPTX 页面图是近似渲染**：两者都不走 LibreOffice，而是读结构自己重画版面（DOCX 用 python-docx
  流式排版，PPTX 用 python-pptx 读形状），**不需要 soffice**。代价是分页/字距/阴影等与真实引擎有出入，
  看"改了哪儿"完全够用，**不能当排版或视觉定稿依据**。PPTX 脚本另会回报 `overflow`（文字溢出文本框的页），
  DOCX/PPTX 都会回报 `errors`（渲染失败的元素），这些都要如实转述，不要吞掉。
  要真实版面时才用 `render_docx_pages.py`（依赖 LibreOffice，本机没装会直接报
  `Missing LibreOffice executable`）。
- **对比图打不开**：先验证路径与文件存在，再发 `<media>` 标签；发不出去的图不算交付。

## Windows (win32) platform notes

- Python 解释器命令是 `python`（不是 `python3`）。脚本已内置 `sys.stdout.reconfigure(encoding="utf-8")`，
  控制台输出中文不会乱码；需要稳定读取结果时加 `--json-out <路径>`，再用 Read 工具读该 JSON 文件。
- 中文字体优先级：`C:\Windows\Fonts\msyh.ttc` → `simhei.ttf` → 思源黑体；等宽字体 `consola.ttf`。
- 对比图目录 `.mavis-compare/` 建议加入项目的 `.gitignore`，避免对比产物进版本库。
- PowerShell 5.1 不支持 `&&`，脚本命令分行执行，每条检查退出码。

## Examples

**Input**：用户说「把封面图的主色换成蓝色，另外出一版」

**Output**：
- `cover.png` 原版保留；`cover_V1.png` 为新版本（同目录）
- `.mavis-compare/cover_V0vsV1_20261027-1340.png` 并排图，橙色框标出换色区域
- 收尾给出版本链表格 + 该图 + 一句改动说明

**Input**：用户说「给这个配置加个 FOOTER 常量」

**Output**：
- `site.config.ts` 原版保留；`site.config_V1.ts` 为新版本
- `.mavis-compare/site.config_V0vsV1_*.png` 左右并排 diff 图（红删绿增，带行号）
- 收尾给出版本链表格 + 该图 + 改动摘要；原版未被改动这一点写进摘要
