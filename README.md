# vibe_docx

中文：把难以直接让 AI 编辑的 Word 文档，桥接成一个更易读、更易改、也更适合版本控制的 Markdown 中间层，再尽可能保留样式地写回 DOCX。

English: A DOCX-to-Markdown bridge for AI-assisted editing. It converts style-heavy Word documents into a readable, editable Markdown intermediate format, then writes the content back to DOCX while preserving as much Word structure and styling as practical.

## Why This Exists | 为什么要做这个项目

中文：

- 直接让 AI 改 `.docx` 很别扭，因为 Word 的结构和样式信息不透明，不适合做细粒度编辑。
- 纯粹把 `.docx` 转成普通 Markdown 又会丢掉大量 Word 语义，例如段落样式、首行缩进、公式对象、表格元数据、字符级格式等。
- `vibe_docx` 的目标不是“把 Word 完全变成 Markdown”，而是提供一个对 AI 和人类都更友好的中间表示：可见内容尽量保持普通 Markdown，可见处尽量可编辑，不得不保留的 Word 专有信息则通过隐藏 marker 存下来。

English:

- Asking an AI to edit `.docx` directly is awkward because Word files are rich, opaque containers rather than clean text-first documents.
- A plain DOCX-to-Markdown export loses too much: paragraph styles, first-line indent, OMML equations, table layout, merged cells, run-level formatting, and more.
- `vibe_docx` is not trying to turn Word into pure Markdown. It creates a practical intermediate representation: visible content stays as normal Markdown when possible, while Word-specific metadata is preserved through hidden markers.

## Core Workflow | 核心工作流

```text
DOCX -> Markdown bridge -> human / AI edits -> DOCX
```

中文：

1. 从 DOCX 导出 Markdown。
2. 在 Markdown 中编辑正文、公式、表格和一部分样式。
3. 使用原 DOCX 作为模板，将 Markdown 回写成新的 DOCX。
4. 如有需要，再跑一遍 roundtrip 验证报告。

English:

1. Export a DOCX file into Markdown.
2. Edit the Markdown with a human or an AI.
3. Rebuild a DOCX, optionally using the original DOCX as a style/template source.
4. Run a roundtrip validation report when you want a quick fidelity check.

## What It Currently Preserves | 当前已支持的保真项

中文：

- 标题与普通段落
- 加粗、斜体、粗斜体
- 字体颜色
- 下划线、删除线
- 高亮与字符背景色
- 字符样式级 override
- 段落样式、段落对齐、首行缩进
- 行内公式与独立公式，支持 OMML payload marker
- Markdown 表格与 Word 表格元数据
- 表格列宽、单元格元数据、合并单元格、单元格内部换行/换段
- roundtrip 统计校验：标题、样式、加粗、下划线、删除线、高亮、字符样式、表格元数据、公式数量、字体表等

English:

- Headings and normal paragraphs
- Bold, italic, and bold+italic runs
- Text color
- Underline and strikethrough
- Highlight and run background fill
- Character-style overrides
- Paragraph styles, alignment, and first-line indent
- Inline and display math with optional OMML payload markers
- Markdown tables plus Word table metadata
- Table widths, cell metadata, merged cells, and line/paragraph breaks inside cells
- Roundtrip checks for headings, styles, bold, underline, strike, highlight, character styles, table metadata, math counts, font table, and more

## Quick Start | 快速开始

### Requirements | 依赖

```bash
python -m pip install -r requirements.txt
```

或：

```bash
python -m pip install python-docx lxml
```

中文：脚本目前是单文件原型，直接运行 `docx_md_bridge.py` 即可。

English: The current implementation is a single-file prototype. Run `docx_md_bridge.py` directly.

### 1. DOCX -> Markdown

```bash
python3 docx_md_bridge.py docx2md input.docx output.md
```

如果你不想在 Markdown 中嵌入 OMML payload marker：

```bash
python3 docx_md_bridge.py docx2md input.docx output.md --no-embed-omml
```

### 2. Markdown -> DOCX

```bash
python3 docx_md_bridge.py md2docx edited.md rebuilt.docx --template-docx input.docx
```

中文：`--template-docx` 很重要。它允许脚本从原始 DOCX 继承样式表、字体与可复用结构，否则很多 Word 专有样式无法恢复。

English: `--template-docx` matters. It lets the bridge reuse the original style definitions, fonts, and other Word assets. Without it, many Word-specific styles cannot be faithfully restored.

### 3. Roundtrip Validation

```bash
python3 docx_md_bridge.py roundtrip input.docx
```

这会生成：

- `input.md`
- `input.roundtrip.docx`
- `input.roundtrip.md`
- `input.roundtrip_report.json`

## Markdown Protocol | Markdown 标记协议

中文：这个项目不是使用“纯 Markdown”，而是使用“Markdown + visible HTML wrappers + hidden markers”。

English: This project does not use plain Markdown alone. It uses Markdown plus visible HTML wrappers and hidden markers.

Examples:

```md
**bold**
*italic*
<span style="color:#C00000">Red text</span>
<u data-docx-underline="double">double underline</u>
~~strikethrough~~
<span style="background-color:#CCE5FF" data-docx-highlight="yellow">highlighted</span>
<span data-docx-rstyle="MyCharacterStyle">character style</span>
<!--DOCX_PSTYLE:...-->
<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
<!--DOCX_PALIGN:center-->
<!--DOCX_TABLE_META:...-->
<!--OMML_INLINE_Z:...-->
<!--OMML_BLOCK_Z:...-->
```

Detailed marker reference:

- 中文详细说明见 [docx_md_bridge_markers.md](docx_md_bridge_markers.md)
- The detailed marker reference currently lives in [docx_md_bridge_markers.md](docx_md_bridge_markers.md)

## Repository Layout | 仓库结构

```text
vibe_docx/
├── README.md
├── docx_md_bridge.py
└── docx_md_bridge_markers.md
```

## Recommended Use Cases | 推荐使用场景

中文：

- 固定模板、固定风格的作业、报告、技术文档
- 需要让 AI 大幅修改内容，但又不想手工重排 Word 格式
- 需要把内容放进 Git 做 diff、review 和版本控制

English:

- Homework, reports, or technical documents with stable templates
- AI-assisted rewriting where direct Word editing is too brittle
- Workflows that benefit from Git diffs, code review, and text-first editing

## Non-Goals | 非目标

中文：

- 不是一个“任意 DOCX 完全无损 roundtrip”的通用转换器
- 不是一个完整的 Word 渲染引擎
- 不是 Pandoc 的替代品，而是偏向固定样式文档和 AI 编辑场景的桥接工具

English:

- Not a universal lossless converter for arbitrary DOCX files
- Not a full Word rendering engine
- Not a replacement for Pandoc; it is a bridge aimed at fixed-style documents and AI editing workflows

## Current Status | 当前状态

中文：这是一个已经可用的原型，重点在于让“DOCX -> AI editable text -> DOCX”这条链路变得实际可用。它已经覆盖了一批对真实文档很关键的样式与表格能力，但距离“覆盖 Word 的绝大部分功能”仍有空间。

English: This is a working prototype focused on making the “DOCX -> AI-editable text -> DOCX” loop practical. It already covers a meaningful subset of Word features that matter in real documents, but it does not aim to cover all of Word.

## Notes For Public Release | 面向公开发布的说明

中文：如果你想在自己的文档上使用这套桥接协议，建议优先准备一份稳定模板 DOCX，并把它作为 `--template-docx` 的来源。很多字符样式和段落样式只有在目标文档样式表中存在时才能真正恢复。

English: If you want to use this bridge on your own documents, start with a stable template DOCX and reuse it via `--template-docx`. Many paragraph and character styles can only be restored when the target document actually contains those style definitions.