# 08_AI 集成与提示词（OpenAI 兼容）

## 8.1 文档目的

本文档定义 AI（OpenAI 兼容接口）在本项目中的集成方式、调用协议与提示词（Prompt）规范，目标是：

* 将 Gemini 用作**OCR/文本识别引擎**，而不是版面结构推断引擎
* 让返回结果 **严格为 JSON**，便于校验、重试、降级
* 支持两类核心能力：

  1. 表格文字回填（给定单元格 bbox → 返回每个 cell 的 text）
  2. 正文段落识别（遮罩表格后整页 → paragraphs）
* 给出完整可用的提示词模板、返回 JSON 规范、错误处理与重试策略

> 重要原则：
> **表格结构由 OpenCV 推断**（rowspan/colspan），Gemini 只识别文字并回填；这样对合并单元格场景稳定性最高。

---

## 8.2 集成概览

### 8.2.1 两类调用（必须实现）

* **T1：TableCell OCR（表格文字回填）**

  * 输入：表格裁剪图 + 单元格列表（id + bbox）
  * 输出：每个 id 的 text
  * 调用频率：每张表格 1 次（推荐），而不是每个单元格 1 次

* **T2：Page Text OCR（正文段落识别）**

  * 输入：遮罩掉表格区域后的整页图片
  * 输出：段落列表（role + text）
  * 调用频率：每页 1 次

### 8.2.2 不在本项目中使用 Gemini 做的事情

* 不让 Gemini 识别表格行列结构、合并单元格关系
* 不让 Gemini 输出 HTML/Markdown/富文本
* 不让 Gemini 进行“纠错补全式创作”（必须强调**照抄**）

---

## 8.3 AI Client 设计（架构与接口）

### 8.3.1 关键能力要求

* 统一封装 `HttpClient`
* 支持超时、重试、并发限流（SemaphoreSlim）
* 使用 OpenAI 兼容 `chat/completions` 接口，需提供 Base URL（带版本号）与 Model
* 支持图片编码（PNG/JPEG Base64 → `image_url` data URL）
* 严格 JSON 输出：失败则触发重试（使用更强约束 prompt）
* 输出内容做基础过滤（移除控制字符等）

### 8.3.2 推荐接口（与架构文档一致）

```csharp
public interface IGeminiClient
{
    Task<TableCellOcrResult> RecognizeTableCellsAsync(
        Bitmap tableImage,
        IReadOnlyList<CellBoxForOcr> cells,
        CancellationToken ct);

    Task<PageTextOcrResult> RecognizePageParagraphsAsync(
        Bitmap pageImageMasked,
        CancellationToken ct);
}
```

数据结构建议：

```csharp
public sealed class CellBoxForOcr
{
    public string Id { get; init; } = "";
    public int X { get; init; }
    public int Y { get; init; }
    public int W { get; init; }
    public int H { get; init; }
}
public sealed class TableCellOcrResult
{
    public Dictionary<string, string> TextById { get; init; } = new();
    public string? RawJson { get; init; } // 调试用
}
public sealed class PageTextOcrResult
{
    public List<ParagraphDto> Paragraphs { get; init; } = new();
    public string? RawJson { get; init; }
}
public sealed class ParagraphDto
{
    public string Role { get; init; } = "body"; // "title" or "body"
    public string Text { get; init; } = "";
}
```

---

## 8.4 调用策略：并发、超时、重试

### 8.4.1 并发限制（推荐默认）

* 全局 Gemini 并发：**2**（可配置 1/2/4）
* 因为每页通常 1 次正文 + 若干表格，限制并发可以降低超时/限流风险

### 8.4.2 超时建议

* 表格 OCR：60–120 秒（视表格大小与网络）
* 正文 OCR：60–120 秒（视页图大小）

### 8.4.3 重试策略（必须）

仅在“可恢复错误”重试 1 次（可配置）：

* 网络超时、HTTP 429/5xx
* 返回不是合法 JSON
* JSON 合法但 schema 不合法（缺字段/类型错误）
* 表格 OCR 结果缺失大量 id（例如缺失率 > 20%）

重试时应升级 prompt（更强约束、强调“只输出 JSON”）。

---

## 8.5 JSON 输出协议（强约束）

### 8.5.1 TableCell OCR 输出 schema

必须输出如下格式（只允许这些字段；不要多余解释）：

```json
{
  "cells": [
    {"id": "p1_t0_r0_c0", "text": "......"},
    {"id": "p1_t0_r0_c1", "text": "......"}
  ]
}
```

规则：

* `cells` 必须是数组
* 每个元素必须包含：

  * `id`：与输入一致的字符串
  * `text`：字符串（允许空串）
* 若无法识别某 cell：`text` 输出空串，不得编造
* 输出中不得包含 Markdown、代码块围栏、解释性文字

### 8.5.2 Page Text OCR 输出 schema

```json
{
  "paragraphs": [
    {"role": "title", "text": "......"},
    {"role": "body", "text": "......"}
  ]
}
```

规则：

* `role` 仅允许：`title` 或 `body`（无法判断统一给 body）
* `text` 必须为字符串
* 允许在 `text` 内使用 `\n` 表示段内换行
* 段落间不需要额外空行（由数组分隔）

---

## 8.6 提示词（Prompt）最终模板

> 说明：以下提示词以“系统/开发者提示 + 用户提示内容”为结构描述。实际调用时你可以将其拼成一个完整 prompt（或按 Gemini 支持的 message roles 分段）。
> 核心点：**强制 JSON 输出** + **照抄** + **不推断**。

### 8.6.1 T1：表格单元格文字回填 Prompt（强烈推荐使用）

**目标**：给定表格图片 + 单元格 bbox 列表，输出每个 id 的文字。

#### Prompt（建议直接用）

**指令：**

1. 你是一个 OCR 引擎。
2. 你将收到一张“表格图片”，以及一个 `cells` 列表，每个 cell 提供 `id` 和矩形坐标（x,y,w,h）。
3. 你的任务是识别每个矩形区域里的文字，并返回 JSON。
4. **必须严格输出 JSON，禁止输出任何解释、Markdown、代码块。**
5. **不得猜测/补全**。看不清就输出空字符串。
6. 必须保留原文：中文、英文字母、数字、标点、空格；不要改写数字格式。
7. 识别时不要把表格线当成字符；只输出文字内容。

**输出格式（必须一致）：**

```json
{"cells":[{"id":"...","text":"..."}, ...]}
```

**输入（由程序填充，示意）：**

* `cells`：`[{id,x,y,w,h}, ...]`

> 重要：程序应在 prompt 中附上 cells JSON（或以结构化输入字段传入），并附上表格图片。

#### 额外建议（减少错字）

在 prompt 末尾加一条：

* 若字符像 “O/0、I/1、S/5” 难分辨，请以图中真实形状为准，不要自作判断。

---

### 8.6.2 T1-重试版 Prompt（更强约束）

当第一次返回 JSON 不合法/缺 id 时使用：

强化条款：

* “如果输出包含任何非 JSON 字符，将被判为失败。”
* “必须包含输入中所有 id，顺序不限。”

---

### 8.6.3 T2：正文段落识别 Prompt（遮罩表格后的整页）

**目标**：输出段落列表，尽量保留原分段与换行，且不包含表格内容（因为表格已被遮罩）。

#### Prompt（建议直接用）

**指令：**

1. 你是 OCR 引擎，识别图片中的正文文字。
2. 图片中可能存在被遮罩的空白区域（那是表格区域），请忽略空白区域，不要输出表格内容。
3. 按阅读顺序输出段落列表。
4. **必须严格输出 JSON，禁止输出任何解释、Markdown、代码块。**
5. **不得猜测/补全**，看不清的字可用 `?` 代替（但不要大量生成），或省略该词。
6. 尽量保持原换行：段内换行使用 `\n`。
7. 若出现明显标题（居中/字号更大/“第X章”等），将该段 `role` 标为 `title`，否则 `body`。

**输出格式（必须一致）：**

```json
{"paragraphs":[{"role":"title|body","text":"..."}, ...]}
```

---

## 8.7 输入组织与图片大小控制

### 8.7.1 图片格式与压缩

* 表格图：PNG（线条清晰）或高质量 JPEG（体积小）
* 正文页图：JPEG（质量 85–92）通常足够，可显著降低传输体积

建议策略：

* 内部处理用无损（Mat/Bitmap）
* 发给 Gemini 前按策略编码：

  * 若图片边长 > 2500–3000px，可等比缩放到较合理尺寸（避免请求过大导致超时）
  * 但不要缩得太小，否则小字表格会损失

### 8.7.2 单元格 bbox 列表大小

对于单表格单次调用：

* 若 cells 数量非常大（> 300），可能导致 prompt 过长或处理变慢
  应对：
* 分两次调用（按行分块）
* 或对极小单元格（如纯序号列）可合并识别策略（增强项）

---

## 8.8 结果校验与回填规则（Gemini 输出后必须做）

### 8.8.1 表格回填校验

* 必须覆盖所有输入 `id`
* 允许 text 为空
* 过滤控制字符
* 若缺失率 > 阈值（默认 20%）：

  * 触发重试（使用重试 prompt）
  * 再失败则降级（见 8.10）

### 8.8.2 正文段落校验

* `paragraphs` 必须为数组，元素含 `role/text`
* role 非法时强制置为 `body`
* 清理不可见控制字符
* 若返回空 paragraphs：

  * 触发重试（更强 prompt）
  * 再失败：降级为单段文本（让 Gemini 输出全文，不分段）

---

## 8.9 与 IR 的映射（必须一致）

* 表格：

  * `cells[id].text` → 回填到对应 `TableCellIr.Text`
  * `id` 应与 `TableCellIr.CellId` 对齐
* 正文：

  * `paragraphs[]` → 每条生成一个 `ParagraphBlockIr`
  * `role` → `ParagraphRole.Title/Body`
  * `text` → `ParagraphBlockIr.Text`

---

## 8.10 降级策略（Gemini 层面建议）

当出现以下情况可触发降级：

* 多次返回非 JSON
* 大量缺 id / text 全空
* 网络持续失败

建议降级顺序：

1. **切换到更宽松识别 prompt**（仍严格 JSON，但允许用 `?` 占位）
2. 表格 OCR 降级为“整表文本”：

   * 输出 `{"lines":["...","..."]}` 或 `{"text":"..."}`
   * 写入 Word 时用制表符/换行拼成可读表格（牺牲结构但可读）
3. 正文降级为单段文本：

   * 输出 `{"text":"..."}`（只要能拿到正文）

> 具体触发条件和用户提示在《09_结果校验_重试_降级策略》中统一定义；本模块需提供错误分类与原始响应保存（可选）用于诊断。

---

## 8.11 日志与诊断建议（与安全平衡）

### 8.11.1 建议记录

* 请求类型（T1/T2）
* page/table 标识（不含正文）
* 耗时、图片尺寸、cells 数量
* 返回 JSON 是否通过校验、缺失率
* 错误码与重试次数

### 8.11.2 避免记录

* 完整正文内容（可能敏感）
* API Key
* 若需保存 RawJson 用于调试，建议仅在“保留诊断文件”模式下启用，并提示用户可能包含敏感信息。

---

## 8.12 本模块验收点

* T1：对含合并单元格表格，能稳定返回覆盖全部 id 的 JSON，并回填正确文字
* T2：遮罩表格后能输出合理段落数组，文本大致正确
* 返回结果 JSON 解析/校验失败时可重试，重试后成功率提升
* 并发与超时策略合理，不阻塞 UI，不造成大量失败
* 日志可定位问题（超时、缺 id、图片过大等）
