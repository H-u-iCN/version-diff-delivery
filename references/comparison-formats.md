# 对比形式选型表

按文件类型选主形式，选不出图的形式用文字对比兜底——重点是让用户一眼看出"改了哪里"。

## 一览

| 文件类型 | 主形式 | 命令 / 手法 |
|---|---|---|
| 代码、配置、Markdown、JSON、YAML、CSV、日志 | 左右并排 diff 图（红删绿增，带行号） | `render_compare.py text` |
| PNG / JPG / WEBP 等位图 | 并排图 + 橙色变更区域框 | `render_compare.py image` |
| PPTX / 演示文稿 | 逐页并排拼图（变化页打橙框）+ 单份接触印相表 | `render_deck.py compare` / `sheet` |
| DOCX / Word | 逐页并排拼图（变化页打橙框）+ 单份接触印相表 | `render_doc.py compare` / `sheet` |
| PDF | 页面图 → `render_compare.py pair`；否则文本差异表 | 见下方 |
| XLSX / 表格数据 | 单元格级 before/after 表格（必要时另附 CSV 文本 diff 图） | 见下方"表格类" |
| 视频 | 关键帧并排 + 时间码表 | 见下方"视频类" |
| 纯文本小改动（< 5 行） | unified diff 代码块 | `git diff --no-index` |

## 文本 / 代码类

```powershell
python "<skill_dir>/scripts/render_compare.py" text `
  --left "<旧文件>" --right "<新文件>" `
  --out "<项目根>/.mavis-compare/<名称>_V0vsV1_20261027-1340.png" `
  --title "<文件名> V0 -> V1" --context 3
```

- `--context N`：变更行上下保留 N 行上下文，默认 3；改动密集时调小。
- `--max-rows`：单图最多渲染多少行，默认 140；超出自动分页成 `_p1`、`_p2`…，全部交付。
- 附带形式：`git diff --no-index -- <旧> <新>` 拿到标准 unified diff，配合对比图一起给，
  适合让用户复制到 PR 或 issue 里。

## 位图类

```powershell
python "<skill_dir>/scripts/render_compare.py" image `
  --left "<旧图>" --right "<新图>" `
  --out "<项目根>/.mavis-compare/<名称>_V0vsV1_20261027-1340.png" --title "<文件名> V0 -> V1"
```

- 差异检测走 RGB 逐通道差分，同明度换色（米色→蓝）也能检出；橙色框带序号角标。
- 两侧尺寸不同时不做像素比对，页脚会写明"尺寸不同"，此时应补充尺寸/分辨率参数表。
- 需要"没改但被重新编码"这类信息时，补一张参数表（尺寸、主色、文件大小、位深）。

## 页面 / 幻灯片类（PPTX / DOCX / PDF）

### PPTX：接触印相表 + 逐页并排（首选路线）

`render_deck.py` 用 python-pptx 读结构、Pillow 重画页面，**不依赖 LibreOffice / PowerPoint**。
渲染思路来自用户工程「经典原著阅读分享」里的 `tools/render_preview.py`，此处通用化后加了分组形状、
圆角矩形/三角、渐变兜底、图片失败可见化、每页溢出检测。

新旧两版逐页并排、变化页自动打橙框标出改了哪儿：

```powershell
python "<skill_dir>/scripts/render_deck.py" compare `
  --left "<旧.pptx>" --right "<新.pptx>" `
  --out "<项目根>/.mavis-compare/<名称>_V0vsV1_20261027-1340.png" `
  --title "<名称> V0 -> V1" --thumb-w 380
```

只展示一份 deck（不改、只想让用户看全貌）：

```powershell
python "<skill_dir>/scripts/render_deck.py" sheet --pptx "<deck.pptx>" `
  --out "<项目根>/.mavis-compare/<名称>_全览_20261027-1340.png" --cols 4 --thumb-w 440
```

- 拼图每页带 `p01`、`p02` 标签；手机点开大图就能缩放看细节——这是用户明确要的展示形态。
- 输出 JSON 里三个字段要如实转述：`changed_slides`（变化页码）、`overflow`（文字溢出文本框的页）、
  `errors`（图片渲染失败的页）。
- 逐页 PNG 单独导出（供 `render_compare.py pair` 用）：
  `render_deck.py slides --pptx <f> --outdir <dir>`。
- 这是**近似渲染**：阴影、渐变、复杂图表、艺术字与最终放映会有出入。看"改了哪儿"够用，
  当视觉定稿依据不够。

### DOCX / Word：接触印相表 + 逐页并排

`render_doc.py` 用 python-docx 读 WordprocessingML，按页面尺寸与页边距把段落、表格、内嵌图
**流式排版**，超出页高就分页，再拼成接触印相表。**不依赖 LibreOffice / Word**。

新旧两版逐页并排、变化处自动打橙框：

```powershell
python "<skill_dir>/scripts/render_doc.py" compare `
  --left "<旧.docx>" --right "<新.docx>" `
  --out "<项目根>/.mavis-compare/<名称>_V0vsV1_20261027-1340.png" `
  --title "<名称> V0 -> V1" --thumb-w 420
```

只展示一份文档：

```powershell
python "<skill_dir>/scripts/render_doc.py" sheet --docx "<文档.docx>" `
  --out "<项目根>/.mavis-compare/<名称>_全览_20261027-1340.png" --cols 4 --thumb-w 440
```

- 输出 JSON 里 `changed_pages` 是有变化的页码，`errors` 是渲染失败的元素（如图片解码失败），
  两项都要如实转述。
- 逐页 PNG 单独导出：`render_doc.py pages --docx <f> --outdir <dir>`。
- **分页是近似的**：Word 的排版引擎（字距、行距、标点挤压、孤行控制、样式继承）比这里复杂，
  页的切分位置可能与 Word 实际分页有出入；内容本身按顺序完整排下来。
- 要真实版面时走 `render_docx_pages.py`（需 LibreOffice，本机没装会报
  `Missing LibreOffice executable`），这时改用结构/文本差异表兜底。

## 表格类（XLSX / CSV）

- 主形式：字段级 before/after 表——列 | 行定位 | 旧值 | 新值 | 变化类型（改值/新增行/删行）。
- 辅形式：把两个版本各导出成 CSV 后走 `text` 模式出并排 diff 图，适合改动集中在单元格文本时。
- 公式与格式变化（颜色、合并单元格、公式串）要单列一行说明，图上看不出来。

## 视频类

- 主形式：关键帧并排（改动点前后各取一帧，`pair` 模式）。
- 辅形式：时间码表——时间点 | 旧画面描述 | 新画面描述 | 变化类型。

## 非图像对比形式清单（图之外的补充）

这些形式可以和图叠加使用，用来补图里表达不了的信息：

1. **unified diff 代码块**：精确到行，可复制。
2. **字段级 before/after 表**：配置、结构化数据首选。
3. **参数差异表**：尺寸、分辨率、颜色值、文件大小、时长。
4. **版本链表格**：文件 | 版本 | 时间 | 改动摘要，一条链走到底。
5. **结构树差异**：目录结构、页码、章节层级变动。
6. **文字要点清单**：取舍理由、没做什么、已知限制。

## 交付自检

- 每个改动文件都有对应的对比产物（图或表），没有"漏网"文件。
- 图和文字说的是同一件事，数字对得上（新增/删除行数与表格一致）。
- 表格里出现的数字来自实际渲染结果或实际读取的元数据，不是估的。
