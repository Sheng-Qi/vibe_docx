# docx_md_bridge Markdown 标记说明

English note: this is the detailed marker reference for the bridge protocol. The main overview is bilingual in `README.md`, while this file remains Chinese-first and implementation-oriented.

本文档描述 `docx_md_bridge.py` 在 DOCX <-> Markdown 桥接过程中会使用到的 Markdown 标记格式，包括：

- 普通可见 Markdown 语法
- 为了保真而插入的隐藏 HTML 注释 marker
- 表格、公式、段落样式等元数据的作用范围

这份说明以当前脚本实现为准，不是泛化的 Markdown 规范。

注意：目前脚本仅支持 DOCX 中的一部分格式特征，对于过于复杂的 Word 文档可能无法完全保真。本文档描述的是脚本目前已经实现并会真正读写的那些标记和语法。

## 1. 总体原则

脚本生成的 Markdown 分成两层：

- 可见层：尽量保持为可直接编辑的普通 Markdown，例如标题、段落、加粗、斜体、表格、`$...$` 公式。
- 隐藏层：用 HTML 注释保存 Word 专有信息，例如段落样式、首行缩进、段落对齐、表格合并单元格、OMML 数学对象原始载荷等。

如果你删除了隐藏 marker，可见内容通常还可以转回 DOCX，但会丢失对应的 Word 专有格式。

## 2. 可见 Markdown 语法

### 2.1 标题

Heading 会导出成标准 Markdown 标题：

```md
# 一级标题
## 二级标题
```

当前脚本通过 `#` 的数量识别 heading level，并在回写 DOCX 时映射到 Word 的 Heading 样式。

### 2.2 普通段落

普通段落直接写成普通文本：

```md
This is a normal paragraph.
```

如果段落前面有隐藏 marker，这些 marker 会作用到紧随其后的那个 block。

### 2.3 加粗、斜体、粗斜体

当前脚本会把相邻且格式相同的 run 合并后，再输出为最小化的 emphasis 标记：

```md
**bold**
*italic*
***bold and italic***
```

说明：

- `**...**` 表示加粗
- `*...*` 表示斜体
- `***...***` 表示同时加粗和斜体
- 脚本会尽量避免生成冗余的 `****` 之类标记

### 2.4 字体颜色

字体颜色使用可见的 HTML `span`，不使用隐藏 marker：

```md
<span style="color:#C00000">Red text</span>
<span style="color:#00B050">**Green bold**</span>
<span style="color:#0070C0">*Blue italic*</span>
<span style="color:#7030A0">***Purple bold italic***</span>
```

说明：

- 当前实现支持 `#RRGGBB` 形式的 6 位十六进制颜色
- 推荐把加粗/斜体写在颜色 `span` 内部，即 `<span style="color:#RRGGBB">**text**</span>`
- 这套语法对普通段落和表格单元格都生效
- 如果删除这个 `span`，文本仍然存在，但颜色信息会丢失

### 2.5 下划线与删除线

下划线与删除线使用可见语法：

```md
<u>underlined</u>
~~strikethrough~~
<u>~~both~~</u>
<u data-docx-underline="double">double underline</u>
```

说明：

- 下划线使用 HTML `<u>...</u>`
- 普通删除线使用 Markdown 的 `~~...~~`
- 如果需要保留非默认下划线类型，可以在 `<u>` 上增加 `data-docx-underline`，例如 `double`
- 推荐嵌套顺序是：最外层颜色/背景/字符样式 `span`，中间 `<u>`，再内层 `~~...~~`，最内层才是 `**...**` / `*...*`

### 2.6 文字高亮与背景色

高亮/背景色也使用可见 `span`：

```md
<span style="background-color:#FFFF00">Highlight</span>
<span style="background-color:#CCE5FF">Background fill</span>
<span style="background-color:#FFFF00" data-docx-highlight="yellow">Word highlight</span>
```

说明：

- `background-color:#RRGGBB` 表示字符级背景色
- 如果这是 Word 自带 highlight 颜色，脚本会额外写出 `data-docx-highlight="..."`，例如 `yellow`
- 如果只保留 `background-color` 而删掉 `data-docx-highlight`，通常仍能保留背景色，但可能从 Word 的 highlight 语义退化为普通底纹

### 2.7 字符样式级 override

字符样式通过 `span` 的 `data-docx-rstyle` 保存：

```md
<span data-docx-rstyle="Emphasis">Styled text</span>
<span style="color:#C00000" data-docx-rstyle="MyCharacterStyle">**Styled and colored**</span>
```

说明：

- `data-docx-rstyle` 保存的是 Word run/character style 的 style id
- 回写 DOCX 时，脚本会优先在模板或目标文档样式表中查找同名 character style 并重新套用
- 如果目标 DOCX 没有这个字符样式，文本内容会保留，但该字符样式本身无法恢复

### 2.8 行内公式

可见层写成：

```md
This is inline math: $x^2 + y^2$.
```

如果需要保留原始 OMML 对象，会紧跟一个隐藏 marker：

```md
This is inline math: $x^2 + y^2$<!--OMML_INLINE_Z:...-->
```

### 2.9 独立公式块

可见层写成：

```md
$$
\frac{a}{b}
$$
```

如果需要保留原始 OMML 对象，会在公式块后紧跟一个隐藏 marker：

```md
$$
\frac{a}{b}
$$
<!--OMML_BLOCK_Z:...-->
```

### 2.10 表格

表格主体仍然是普通 Markdown 表格：

```md
| Col A | Col B |
| --- | --- |
| A1 | B1 |
```

但如果原始 Word 表格包含列宽、合并单元格、单元格样式等结构信息，会在表格上方加一个隐藏的 `DOCX_TABLE_META` marker。

### 2.11 表格单元格中的换行与换段

表格 cell 内部的文本支持两种层级：

```md
| Line 1<br>Line 2 | Para 1<br><br>Para 2 |
```

含义：

- 单个 `<br>`：同一个段落内部的换行
- 两个连续的 `<br><br>`：单元格内部开始一个新段落

这条规则只针对表格单元格内容。

## 3. 隐藏 marker 总览

当前实现中会用到以下隐藏 marker：

```md
<!--OMML_INLINE_Z:...-->
<!--OMML_BLOCK_Z:...-->
<!--DOCX_PSTYLE:...-->
<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
<!--DOCX_PALIGN:center-->
<!--DOCX_TABLE_META:...-->
<!--EMPTY_P-->
```

下面分别说明。

## 4. 数学公式 marker

### 4.1 `OMML_INLINE_Z`

格式：

```md
<!--OMML_INLINE_Z:BASE64_ZLIB_OMML_XML-->
```

作用：

- 绑定在一个行内公式 `$...$` 后面
- 保存原始 Word OMML 数学对象
- 回写 DOCX 时优先恢复为真正的 Word 数学对象，而不是纯文本公式

如果删掉它：

- `$...$` 的可见数学文本仍在
- 但回写时可能只能生成近似公式，或者退化为普通文本

### 4.2 `OMML_BLOCK_Z`

格式：

```md
<!--OMML_BLOCK_Z:BASE64_ZLIB_OMML_XML-->
```

作用：

- 绑定在一个独立公式块 `$$...$$` 后面
- 保存 display math 的原始 OMML 结构

## 5. 段落相关 marker

这些 marker 一般写在某个 block 前面，并作用于“下一个非空 block”。

### 5.1 `DOCX_PSTYLE`

格式：

```md
<!--DOCX_PSTYLE:BASE64_JSON-->
```

payload 是未经压缩的 base64 JSON，结构类似：

```json
{"style_id":"13","style_name":"Preset Styles1"}
```

作用：

- 保存非 heading 段落的 Word 段落样式
- 用于恢复自定义段落样式

典型形式：

```md
<!--DOCX_PSTYLE:...-->
This paragraph should use a custom Word paragraph style.
```

说明：

- 导出时，默认 Normal 段落不会生成这个 marker
- Heading 主要由 `#` 决定，`DOCX_PSTYLE` 不应用来替代 heading 级别

### 5.2 `DOCX_FIRST_LINE_INDENT_CM`

格式：

```md
<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
```

作用：

- 指定下一个段落的首行缩进，单位是厘米
- 现在它是显式 marker，不再强制对所有普通段落自动启用

示例：

```md
<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
Indented paragraph.
```

### 5.3 `DOCX_PALIGN`

格式：

```md
<!--DOCX_PALIGN:center-->
```

作用：

- 指定下一个段落的对齐方式
- 当前脚本只在非默认对齐时输出这个 marker

常见值：

- `left`
- `right`
- `center`
- `both`
- 以及 Word `w:jc` 里使用的其他值

示例：

```md
<!--DOCX_PALIGN:center-->
Centered paragraph.
```

说明：

- 顶层普通段落导出时，默认 `left/start` 不会显式输出，以减少噪声
- 表格 cell 内部段落的对齐信息通常保存在 `DOCX_TABLE_META` 里，而不是单独吐成外层 `DOCX_PALIGN`

### 5.4 `EMPTY_P`

格式：

```md
<!--EMPTY_P-->
```

作用：

- 表示一个真正的空段落
- 用于保留 Word 中的空白段落

示例：

```md
First paragraph.

<!--EMPTY_P-->

Second paragraph.
```

注意：

- 普通 Markdown 的空行只是 block 分隔
- `<!--EMPTY_P-->` 才表示要在 DOCX 中生成一个空段落对象

## 6. 表格 marker

### 6.1 `DOCX_TABLE_META`

格式：

```md
<!--DOCX_TABLE_META:BASE64_ZLIB_JSON-->
| ... |
| --- |
| ... |
```

作用：

- 作用于紧随其后的那张 Markdown 表格
- 保存 Word 表格的结构性元数据
- 可见表格仍然保留为普通 Markdown，方便编辑

payload 是 `zlib` 压缩后的 JSON，再做 base64 编码。

### 6.2 当前 `DOCX_TABLE_META` 会保存什么

当前实现里，这个 marker 可能包含以下字段：

```json
{
  "style_id": "aa",
  "layout": "autofit",
  "alignment": "center",
  "table_width": {"type": "auto", "w": "0"},
  "grid_widths": [8359, 1842, 2410, 1337],
  "row_count": 3,
  "col_count": 4,
  "merges": [
    {"row": 0, "col": 0, "rowspan": 1, "colspan": 4}
  ],
  "rows_meta": [
    {},
    {"header": true},
    {}
  ],
  "cells": {
    "0,0": {
      "width": {"type": "dxa", "w": "13948"},
      "vertical_align": "center",
      "shading_fill": "D9EAF7",
      "paragraphs": [
        {"alignment": "center"},
        {}
      ]
    }
  }
}
```

### 6.3 字段含义

#### 表级字段

- `style_id`：Word 表格样式 id
- `layout`：表格布局，例如 `autofit` 或 `fixed`
- `alignment`：整张表格的水平对齐方式
- `table_width`：Word `tblW`
- `grid_widths`：`tblGrid/gridCol` 列宽数组
- `row_count` / `col_count`：逻辑网格大小

#### 合并单元格

- `merges`：单元格合并信息

格式：

```json
{"row": 0, "col": 0, "rowspan": 1, "colspan": 4}
```

表示：

- 从第 `0` 行、第 `0` 列开始
- 横向跨 `4` 列
- 纵向跨 `1` 行

这就是 Word 合并单元格在 Markdown 里的兼容方案：

- 可见层仍然保留占位空列，确保 Markdown 表格可解析
- 真正的 merge 语义写在隐藏 marker 的 `merges` 字段里

#### 行级字段

- `rows_meta`：每一行的行级元数据数组

可能字段：

- `height`: 行高设置
- `header`: 是否为表头行
- `cant_split`: 是否不允许跨页拆分

#### 单元格级字段

- `cells`：字典，key 形如 `"row,col"`

例如：

```json
"1,2": {
  "width": {"type": "dxa", "w": "2410"},
  "vertical_align": "center",
  "shading_fill": "D9EAF7",
  "paragraphs": [
    {"alignment": "center"},
    {"alignment": "left", "first_line_indent_cm": 0.74}
  ]
}
```

含义：

- `width`：单元格宽度
- `vertical_align`：垂直对齐
- `shading_fill`：底纹填充色
- `paragraphs`：cell 内每个段落的段落级 override

### 6.4 cell 内段落与 `<br>` 的对应关系

在 cell 文本里：

- `A<br>B` 表示同一段中换一行
- `A<br><br>B` 表示 cell 内两个段落

例如：

```md
| <span style="color:#C00000">Title</span><br>Subtitle | Para 1<br><br>Para 2 |
```

## 7. marker 的作用范围与消费方式

### 7.1 段落类 marker

以下 marker 属于“挂在下一个 block 前面”的类型：

- `DOCX_PSTYLE`
- `DOCX_FIRST_LINE_INDENT_CM`
- `DOCX_PALIGN`

它们会被解析器暂存，并在遇到下一个非空 block 时消费。

典型例子：

```md
<!--DOCX_PSTYLE:...-->
<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
<!--DOCX_PALIGN:center-->
Actual paragraph text.
```

### 7.2 表格类 marker

`DOCX_TABLE_META` 只作用于其后紧跟的那一张 Markdown 表格：

```md
<!--DOCX_TABLE_META:...-->
| A | B |
| --- | --- |
| 1 | 2 |
```

### 7.3 公式类 marker

- `OMML_INLINE_Z` 只绑定在它前面的那个 `$...$`
- `OMML_BLOCK_Z` 只绑定在它前面的那个 `$$...$$`

## 8. 编辑建议

### 8.1 可以放心手改的部分

- 普通段落文本
- 标题文本
- `**bold**` / `*italic*` / `***both***`
- `<span style="color:#RRGGBB">...</span>`
- 表格里的可见单元格文本
- 公式的可见 LaTeX 文本

### 8.2 尽量不要手改的部分

- `OMML_INLINE_Z`
- `OMML_BLOCK_Z`
- `DOCX_PSTYLE` 的 payload
- `DOCX_TABLE_META` 的 payload

这些 payload 不是为了人工编辑设计的。它们是脚本生成的保真数据。

### 8.3 如果删掉隐藏 marker，会发生什么

- 删除公式 marker：通常还能保留可见公式文本，但 Word 原生公式对象可能丢失
- 删除 `DOCX_PSTYLE`：自定义段落样式可能退回 Normal
- 删除 `DOCX_FIRST_LINE_INDENT_CM`：首行缩进丢失
- 删除 `DOCX_PALIGN`：非默认对齐丢失
- 删除 `DOCX_TABLE_META`：列宽、合并单元格、cell 样式、cell 内段落 override 等表格细节会退化

## 9. 最小示例

下面是一个包含多种 marker 的示例：

```md
<!--DOCX_PSTYLE:eyJzdHlsZV9pZCI6IjEzIiwic3R5bGVfbmFtZSI6IlByZXNldCBTdHlsZXMxIn0=-->
Results and Discussion (+Conclusion) Analysis

<!--DOCX_FIRST_LINE_INDENT_CM:0.74-->
Indented paragraph with **bold**, *italic*, and ***both***.
Indented paragraph with **bold**, *italic*, ***both***, and <span style="color:#C00000">red text</span>.

Inline math: $x^2 + y^2$<!--OMML_INLINE_Z:...-->

$$
\frac{a}{b}
$$
<!--OMML_BLOCK_Z:...-->

<!--DOCX_TABLE_META:...-->
| Merged heading |  |  |
| --- | --- | --- |
| A<br>B | Para 1<br><br>Para 2 | C |
```

## 10. 当前实现边界

虽然目前已经支持了不少 Word 专有信息，但它仍然不是完整的 DOCX 语法树导出器。当前文档所覆盖的 marker，是脚本现在已经实现并会真正读写的那部分。

如果后续脚本再增加字符样式、图片、页眉页脚、列表编号等支持，这份文档也需要同步更新。