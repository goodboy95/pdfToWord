# 10_Word 生成（OpenXML 实现）

## 10.1 文档目的

本文档说明如何将《03_数据结构与IR规范》中的 `DocumentIr` 渲染为 `.docx`，重点覆盖：

* OpenXML SDK 的写入结构（Document/Body/SectionProperties）
* 页面尺寸：**强制 A4** 或 **跟随 PDF 页尺寸**
* 段落：尽量保持原文（换行、标题/正文样式）
* 表格：生成带边框表格，支持 **rowspan/colspan 合并单元格**
* 字体：中文/英文混排时的 EastAsia/Ascii 字体设置
* 分页：每 PDF 页对应 Word 分页
* 常见坑与建议的实现方式（owner 矩阵法）

---

## 10.2 输入输出与整体策略

### 输入

* `DocumentIr doc`
* `DocxWriteOptions opt`

  * `PageSizeMode`：A4 / FollowPdf
  * `FontEastAsia`：默认“微软雅黑”（或宋体）
  * `FontAscii`：默认“Calibri”
  * `DefaultFontSizeHalfPoints`：如 21（10.5pt）或 24（12pt）
  * `Margins`：默认页边距（可配置）
  * `KeepPageBreakPerPdfPage`：true（固定）

### 输出

* `.docx`（OpenXML WordprocessingDocument）

### 总体策略（符合需求“像原文一样”）

1. **分页优先**：每页处理完插入 PageBreak（或 SectionBreak，见下文）
2. **结构优先**：表格结构与合并单元格优先于细节样式
3. **文字尽量原样**：保持段落与段内换行（`\n`）
4. 页面尺寸按用户选项决定（A4 或跟随 PDF）

---

## 10.3 OpenXML 文档骨架

### 10.3.1 最小骨架

* `WordprocessingDocument`

  * `MainDocumentPart`

    * `Document`

      * `Body`

        * `Paragraph` / `Table` / `Paragraph(PageBreak)`
        * `SectionProperties`（页尺寸/页边距/页眉页脚等）

### 10.3.2 SectionProperties 放置位置

Word 的页面大小和边距由 `SectionProperties` 控制，通常放在 body 最后一个元素里。

**两种模式：**

* **全局一致页面大小**（强制 A4）：仅需要一个 SectionProperties（结尾）
* **按页变化大小**（跟随 PDF 页）：需要 **每页一个 section**（用 SectionBreak），否则无法每页不同 PageSize

> 重要：
> 如果用户选择“跟随 PDF 页一致”，且 PDF 可能存在不同大小页面，那么必须采用 **SectionBreak** 每页一个 section。
> 如果 PDF 所有页同尺寸，也可以只设置一次，但实现上建议统一按“每页一个 section”以简化一致性。

---

## 10.4 页面尺寸与边距

### 10.4.1 单位换算（核心）

OpenXML 中页面大小单位为 **twips**（1/20 point；1 inch = 1440 twips）。

PDF 渲染得到的是像素，需要映射到 Word 页面尺寸。由于渲染使用 DPI（dot per inch）：

* `inches = pixels / dpi`
* `twips = inches * 1440 = pixels * 1440 / dpi`

因此：

* `pageWidthTwips = round(widthPx * 1440.0 / dpi)`
* `pageHeightTwips = round(heightPx * 1440.0 / dpi)`

> 注意：这里的 widthPx/heightPx 应使用 **裁切后的页图尺寸**还是原始页图尺寸？
> 建议：**使用裁切后的页图尺寸**作为页面内容基准，但页面大小应对应用户视觉阅读：
>
> * 若“去页眉/页脚”是“不要这些内容”，输出 docx 也应不包含这些区域，因此用裁切后尺寸更一致。
> * 若你希望 Word 页面仍保持原 PDF 页面大小但内容少一截，需要另行设计；本项目按“裁切后就是最终内容”处理更直观。

### 10.4.2 A4 模式

A4：8.27" × 11.69"（约 595×842 pt）
twips：

* 宽：`8.27 * 1440 ≈ 11906`
* 高：`11.69 * 1440 ≈ 16838`

实现建议：

* 直接使用固定值（不用从像素换算）
* 页边距建议（可配置）：

  * 上下左右：`1440` twips（=1 inch）或更符合中文排版的 2.0cm（约 1134 twips）

### 10.4.3 FollowPdf 模式

按上面的像素→twips换算，并考虑方向：

* 若宽 > 高，视为横向页面（Landscape）

  * Word 中可通过交换 PageSize 的 width/height 或设置 `Orient=Landscape`（两种都可）

### 10.4.4 SectionProperties 生成

每个 section 的关键属性：

* `PageSize`：`Width/Height/Orient`
* `PageMargin`：`Top/Bottom/Left/Right`（twips）

> 本项目不需要复杂页眉页脚对象（我们做的是裁切去除），所以无需写入 HeaderPart/FooterPart。

---

## 10.5 段落写入（Paragraph Rendering）

### 10.5.1 IR → OpenXML 段落映射

对 `ParagraphBlockIr`：

* `Role=Title` → 使用标题样式（如 Heading1）
* `Role=Body` → Normal（或自定义 Body 样式）

“像原文一样”的重点：

* 保留段内换行（`\n`）
* 尽量不主动重排（不做自动分词/改写）

### 10.5.2 段内换行实现

Word 的段内换行是 `Break`：

* 将 text 按 `\n` 分割
* 每段片段写一个 Run；片段之间插入 `<w:br/>`

### 10.5.3 样式与字体设置（中文/英文混排关键）

必须设置：

* `RunFonts` 的 `Ascii` 与 `EastAsia`
* 字号（HalfPoints）
  建议统一通过样式控制，而不是每个 Run 设置一遍（效率更好、文档更干净）。

**推荐做法：**

* 在文档开头写入自定义 styles：

  * `BodyText`：Normal 基础上设置字体、字号、行距、首行缩进（如需要）
  * `Title1`：Heading1 基础上设置字体/字号（也可直接用内置 Heading1）
* 每个段落只引用 styleId

> 如果你不想写 styles.xml，也可以在段落/Run 上直接写 `RunProperties`，但会重复较多。

### 10.5.4 首行缩进/行距（尽量像原文）

原文是图片版面，“完全一致”不现实，但可提供一个默认策略：

* 默认：不强制首行缩进（避免和原文差距过大）
* 可选增强：Body 段落设置首行缩进 2 字（约 420 twips）与 1.2~1.5 行距
  由于你明确“希望像原文一样”，建议：
* MVP：**仅保留换行与段落分隔**，不强加缩进
* 后续：提供 UI 选项“启用中文常规排版（首行缩进/行距）”

---

## 10.6 表格写入（Table Rendering）

### 10.6.1 表格基本构建

对 `TableBlockIr`：

* 列数：`NCols`
* 行数：`Rows.Count`

OpenXML 表格由：

* `Table`

  * `TableProperties`（边框等）
  * `TableGrid`（列宽，可选）
  * `TableRow`

    * `TableCell`

### 10.6.2 表格边框（满足“只要有线就行”）

设置 `TableBorders`：

* `Top/Bottom/Left/Right/InsideH/InsideV` 全部 `Single`
* 线宽用默认即可（无需仿线粗）

### 10.6.3 列宽策略（简单可靠）

为了让表格布局不要乱：

* **简单模式**：不设置列宽，让 Word 自动分配（多数情况下可用）
* **推荐模式**：按页面可用宽度均分列宽（更稳定）

  * 页面可用宽度 = `PageWidth - LeftMargin - RightMargin`
  * 每列宽 = 可用宽度 / NCols

OpenXML 列宽单位常用 `dxa`（twips），可直接用 twips。

---

## 10.7 合并单元格实现（rowspan/colspan）— 核心

你的表格经常出现合并单元格，这是最关键部分。

### 10.7.1 为什么需要 owner 矩阵法

直接按行遍历 cell 并立即 merge 容易出现：

* merge 顺序错导致覆盖关系混乱
* 纵向 merge 需要在被合并的 cell 上写 `VerticalMerge continue`，这要求你知道每个网格位置是否被某个主 cell 覆盖

因此推荐：

1. 先基于 IR 构造一个 `owner[r,c] = (r0,c0)` 矩阵
2. 再一次性生成 TableRow/TableCell，并按 owner 设置 GridSpan/VerticalMerge

### 10.7.2 owner 矩阵构造规则

设：

* `R = Rows.Count`
* `C = NCols`

从 IR 建立“主单元格列表”，每个主单元格具有：

* 起点 `(r0,c0)`：该 cell 在行内的起始列位置（需要由行内累积 colspan 推导）
* `rowspan`, `colspan`

构造流程：

* 初始化 `owner[R,C] = null`
* 遍历每行 `r`：

  * 用 `colCursor=0`
  * 遍历该行 `Cells`：

    * 令该 cell 的主坐标为 `(r, colCursor)`
    * 将矩形覆盖区域内所有格的 owner 设为 `(r, colCursor)`

      * 若已占用则冲突（应在 IR 校验阶段处理；此处可 throw）
    * `colCursor += colspan`
  * 行结束后要求 `colCursor == C`（否则 IR 不合法）

### 10.7.3 OpenXML 写入规则（按网格位置逐格写）

生成一个 `Table`，对每个网格位置 `(r,c)` 都生成 `TableCell`（也就是总共 R*C 个 cell），然后用属性表示合并：

* 若 `(r,c)` 是主单元格（owner[r,c] == (r,c)）：

  * 横向合并：设置 `GridSpan = colspan`（如果 colspan>1）
  * 纵向合并：如果 rowspan>1，设置 `VerticalMerge = restart`
  * 写入文本内容到该 cell
* 若 `(r,c)` 不是主单元格（被覆盖）：

  * 横向被覆盖：仍然需要存在一个 cell，但它会被 GridSpan 吃掉一部分；实践中更稳的做法是：

    * 对于同一行中被 colspan 覆盖的后续格，仍创建 cell，但它会被 Word 忽略/折叠；为避免异常，可设置空文本
  * 纵向被覆盖：设置 `VerticalMerge = continue`
  * 不写文本或写空

> 实务建议：
> “横向合并”通常只需要在主 cell 设置 GridSpan，并且**不要再为被横向覆盖的格子设置任何特殊属性**（保持空）即可；
> “纵向合并”必须在每个被覆盖格设置 VerticalMerge continue。

### 10.7.4 更稳的写法（推荐）

* 对每行 `r`，按列 `c` 逐格生成 cell
* 如果发现该格属于某个主 cell 的横向覆盖区（owner 同主 cell，但 c != c0），则：

  * 生成空 cell（不设 GridSpan，不设 VerticalMerge，除非也在纵向覆盖）
* 如果属于纵向覆盖且非主格，则设 `VerticalMerge continue`
* 仅主格写 GridSpan/VerticalMerge restart + 写文本

> 这样最不容易出现 Word 打开后表格损坏或布局错乱。

---

## 10.8 单元格内容写入（Text in cells）

### 10.8.1 写入方式

每个 cell 内通常放一个 `Paragraph`：

* `TableCell` → `Paragraph` → `Run` → `Text`

### 10.8.2 保留换行

若单元格文字包含 `\n`：

* 同段内插入 `Break`（与段落一致）
* 或拆成多个 Paragraph（更像 Excel 行内换行），两者都可。建议用 Break 简单。

### 10.8.3 清理文本

在写入前应用《03》的过滤：

* 去控制字符
* trim 逻辑：不强制 trim（以免破坏原样）；可只去掉两端不可见空白

---

## 10.9 分页实现

### 10.9.1 A4 模式（全局一致页面）

* 文档末尾放一次 `SectionProperties`（A4 + margins）
* 每页结束插入 `Paragraph`，其 `Run` 内含 `Break(Type=Page)`

### 10.9.2 FollowPdf 模式（每页一个 section）

为了每页不同 PageSize，需要用 **SectionBreak (NextPage)**：

做法：

* 每页内容写完后，创建一个“带 SectionProperties 的段落”作为该页结尾
* `SectionProperties` 里写入该页 pageWidth/pageHeight/margins
* `SectionType = NextPage`

这样 Word 会从下一页开始应用新的页面设置。

> 如果 PDF 所有页一致，也可以简化为一次 SectionProperties + PageBreak，但为统一实现逻辑，建议 FollowPdf 下全部使用 SectionBreak。

---

## 10.10 字体与样式（建议落地方案）

### 10.10.1 推荐默认字体

* 中文 EastAsia：微软雅黑（或宋体）
* 英文 Ascii：Calibri

### 10.10.2 样式策略（推荐实现）

创建并写入 Styles（可简化为最少两种）：

* `BodyText`（基于 Normal）
* `Title1`（基于 Heading1 或自定义）

在段落中只引用 styleId：

* `ParagraphProperties` → `ParagraphStyleId`

如果你暂时不想管理 styles：

* 每个 Run 设置 `RunProperties`（字体/字号）
* 这会让文档更“重”，但实现更快，MVP 可接受

---

## 10.11 页面边距策略

### 10.11.1 默认边距（建议）

* A4：左右 2.0cm，上下 2.0cm（约 1134 twips）
* FollowPdf：建议也用同一边距，除非你希望更贴近原版面

### 10.11.2 表格宽度适配

若使用“列宽均分”，则列宽基于可用宽度计算，避免表格超出页面。

---

## 10.12 常见坑与处理建议

### 10.12.1 表格合并导致 Word 打不开/修复

原因通常是：

* rowspan/colspan 冲突（owner 矩阵出现覆盖重叠）
* GridSpan 与 VerticalMerge 设置不一致
* 忘记给 TableCell 放 Paragraph（Word 结构不完整）

建议：

* 生成前先做 owner 矩阵校验（必须）
* 每个 TableCell 内至少有一个 Paragraph
* 对被覆盖 cell 统一写空 Paragraph

### 10.12.2 FollowPdf 页面尺寸不生效

原因：没有使用 SectionBreak 或 SectionProperties 放置位置错误。

建议：

* 每页末尾都插入带 `SectionProperties` 的段落（NextPage）
* 不要只用 PageBreak

### 10.12.3 中英混排字体异常

原因：只设置了 Ascii，没有设置 EastAsia。

建议：

* RunFonts 必须同时设置 `Ascii` 与 `EastAsia`
* 如果出现符号字体异常，再设置 `HighAnsi` 为 Calibri

---

## 10.13 推荐的写入器接口

```csharp
public sealed class DocxWriteOptions
{
    public PageSizeMode PageSizeMode { get; set; } = PageSizeMode.FollowPdf;
    public int Dpi { get; set; } = 300;

    public string FontEastAsia { get; set; } = "微软雅黑";
    public string FontAscii { get; set; } = "Calibri";
    public int DefaultFontSizeHalfPoints { get; set; } = 24;

    public int MarginTopTwips { get; set; } = 1134;
    public int MarginBottomTwips { get; set; } = 1134;
    public int MarginLeftTwips { get; set; } = 1134;
    public int MarginRightTwips { get; set; } = 1134;
}
```

写入器：

```csharp
public interface IDocxWriter
{
    Task WriteAsync(DocumentIr doc, DocxWriteOptions opt, Stream output, CancellationToken ct);
}
```

---

## 10.14 本模块验收点

* 能输出可被 Word 正常打开且可编辑的 docx
* A4 模式：

  * 页面为 A4，分页与 PDF 页一致
* FollowPdf 模式：

  * 页面尺寸可随页变化（通过 SectionBreak 生效）
* 表格：

  * 有边框线
  * 行列结构正确
  * 合并单元格（rowspan/colspan）正确或基本正确，不出现严重错位
* 段落：

  * 保留 `\n` 换行
  * title/body 样式基本可区分
* 中英混排字体不出现明显“乱码/缺字/奇怪替换”
