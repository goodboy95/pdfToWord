# 06_PDF渲染与图像预处理

## 6.1 文档目的

本文档描述“从 PDF 到可用于识别的页图”的完整处理链路，包括：

* PDF 渲染到位图/图片的技术选型与接口约定
* DPI 策略与质量/性能权衡
* 页眉页脚去除（裁切）在流水线中的位置与参数
* 图像预处理步骤（增强文字与表格线、降低噪声）
* 中间结果缓存与调试输出策略

该模块输出的结果将作为后续：

* OpenCV 表格检测与单元格分割
* Gemini 正文识别与表格文字回填
  的统一输入。

---

## 6.2 输入输出定义

### 输入

* `pdfPath`：PDF 文件路径
* `pageIndex`：0-based 页索引（由页码范围解析结果转换）
* `RenderOptions`：渲染参数（DPI、颜色模式等）
* `CropOptions`：页眉/页脚去除选项（裁切模式与比例）
* `PreprocessOptions`：预处理参数（是否 deskew、二值化方式等）

### 输出

* `PageImageBundle`（建议结构，示意）

  * `OriginalColor`：渲染后的原始彩色页图（Bitmap/Mat/PNG 路径）
  * `CroppedColor`：裁切后的彩色页图
  * `Gray`：灰度图
  * `Binary`：二值图（用于表格线检测）
  * `DeskewedColor`：可选（若 deskew 启用）
  * `MaskBase`：后续遮罩表格区域时使用的基图（一般为 CroppedColor）
  * `Meta`：尺寸信息、裁切像素等

> 注：十来页规模下，可用内存对象传递；为增强稳定性与调试能力，建议同时支持落盘缓存（见 6.9）。

---

## 6.3 PDF 渲染（Render）

### 6.3.1 渲染目标

将 PDF 指定页渲染为**高分辨率、清晰、无插值伪影尽可能少**的位图，以利于：

* 文字识别（Gemini）
* 表格线提取（OpenCV）

### 6.3.2 渲染库建议与要求

本项目允许多种渲染实现，但需要满足统一接口（`IPdfRenderer`）：

**渲染要求：**

* 支持按 DPI 渲染
* 输出稳定、无明显黑边/空白错位
* 能获取页数
* 渲染速度可接受（十来页）

**推荐：PDFium**

* 一般足够、速度快
* 注意随程序打包时的 native 依赖

### 6.3.3 DPI 策略

* 默认：**300 DPI**
* 用户可选：200 / 300 / 400

经验建议：

* 200：速度快、文件小，适合大字体或简单表格
* 300：平衡方案（默认）
* 400：小字密集、表格线细、合并单元格多时更稳，但耗时/内存增大

#### 自动建议（可选增强）

* 若检测到表格线断裂较多或 OCR 乱码偏多，可提示用户提高 DPI 重试（不强制自动重跑，以免耗时不可控）。

### 6.3.4 渲染输出格式

* 内存：`Bitmap`（WPF 友好）或 `OpenCvSharp.Mat`
* 建议统一内部处理为 `Mat`，UI 需要展示时再转 `BitmapSource`

**颜色格式建议：**

* 渲染出 `RGB` 或 `BGR` 均可，但 OpenCV 默认 `BGR`
* 避免 Alpha 通道带来的混合问题（如不必要）

---

## 6.4 页眉页脚去除（裁切 Crop）

### 6.4.1 为什么在渲染后立刻裁切

* 页眉页脚常常干扰正文识别与表格检测
* 早裁切能减少后续处理面积，提升性能
* 裁切是确定性操作，且用户可调参（通过预览）

### 6.4.2 选项与参数

* 模式：

  * `None`
  * `RemoveHeader`
  * `RemoveFooter`
  * `RemoveBoth`
* 参数：

  * `HeaderPercent`（默认 0.06）
  * `FooterPercent`（默认 0.06）

> 实际裁切像素应写入 `PageIr.Crop`（见 IR 文档），便于复现。

### 6.4.3 裁切计算

设渲染后原图高度 `H`：

* `cropTopPx = round(H * HeaderPercent)`（仅在去页眉模式）
* `cropBottomPx = round(H * FooterPercent)`（仅在去页脚模式）
* 裁切区域：`y in [cropTopPx, H - cropBottomPx)`

边界约束：

* `cropTopPx + cropBottomPx <= H * 0.5`（建议限制，避免用户误设导致裁掉正文）
* 若裁切后高度过小（比如 < 200px）：应提示参数异常并阻止转换

---

## 6.5 图像预处理总览（Preprocess）

预处理目标：在不破坏文字的前提下，增强表格线与文字对比度，提升：

* 表格线检测稳定性（OpenCV）
* 文本识别准确率（Gemini）

预处理建议产出两条分支：

* **用于表格线检测的 Binary 图**
* **用于 Gemini 的较干净的 Color/Gray 图**

---

## 6.6 预处理步骤（推荐默认链路）

以下步骤以 OpenCV（OpenCvSharp）实现为主，默认建议顺序如下：

### Step 1：转灰度（Gray）

* `Gray = cvtColor(CroppedColor, GRAY)`

目的：

* 降低后续处理复杂度
* 便于阈值化与对比度增强

### Step 2：对比度增强（可选但推荐）

两种方式二选一：

* 简单线性拉伸（快）
* **CLAHE**（对扫描件更稳）

目的：

* 提升淡字、浅线的可见度

参数建议（CLAHE）：

* clipLimit：2.0～3.0
* tileGridSize：8x8 或 16x16

### Step 3：去噪（轻量）

* `medianBlur(Gray, k=3)` 或 `GaussianBlur(Gray, (3,3), 0)`

目的：

* 去掉椒盐噪声，减少二值化后毛刺
* 注意不要过强，避免细线断裂

### Step 4：二值化（Binary）

推荐：

* 自适应阈值 `adaptiveThreshold`（扫描件常用）
  备选：
* Otsu（二值化更快但对光照不均不如自适应）

目的：

* 为表格线提取提供高对比黑白图

参数建议（adaptiveThreshold）：

* blockSize：31～51（随 DPI 可调）
* C：5～15（视背景灰度调整）

输出：

* `Binary`（白底黑字黑线，或反相也可，后续统一即可）

### Step 5：倾斜校正 deskew（可选，默认开启）

对扫描件倾斜（1~3 度）非常常见，deskew 能显著提高表格线垂直/水平一致性。

建议策略：

* 在 `Binary` 上检测主方向角度（霍夫线或最小外接矩形）
* 若角度绝对值 < 阈值（如 0.3°）则不旋转
* 旋转时使用 `warpAffine`，背景填白

注意：

* deskew 最好在“裁切后”进行，且对同页只做一次
* 旋转会改变 bbox，需要统一在 deskew 后图上进行表格检测（推荐：一旦 deskew，就让后续全程使用 deskewed 图）

---

## 6.7 参数与 DPI 的联动建议

为了让不同 DPI 下预处理效果稳定，部分参数建议按图像尺寸自适应：

* AdaptiveThreshold blockSize：

  * 200 DPI：31
  * 300 DPI：41
  * 400 DPI：51
* 形态学核尺寸（表格线提取会在表格算法文档详述）也需按宽高比例确定

原则：

* 不把参数写死到代码里，至少集中在一个 `PreprocessOptions` 配置对象里，便于后期调优。

---

## 6.8 输出用于后续模块的“标准图”

建议在预处理完成后，明确以下标准输出（供下游使用）：

* `PageColorForGemini`：通常为 **裁切后、deskew 后（若启用）的彩色图**

  * 用于正文识别（遮罩表格后）
* `PageBinaryForTable`：通常为 **裁切后、deskew 后（若启用）的二值图**

  * 用于表格线提取、表格 bbox 检测
* `PageGrayOptional`：供调试/备用

> 重要：如果 deskew 启用，下游所有 bbox 坐标都以 deskew 后的图为准，避免坐标不一致。

---

## 6.9 中间结果缓存与诊断包（建议）

### 6.9.1 缓存目录结构（建议）

每个 Job 一个目录（例如 `%TEMP%\Pdf2Word\{jobId}\`）：

```
job_{id}/
  pages/
    p001_original.png
    p001_cropped.png
    p001_gray.png
    p001_binary.png
    p001_deskew.png
  logs/
    job.log
  ir/
    doc_ir.json
```

### 6.9.2 清理策略

* 默认：任务完成后可自动清理（或保留最近 N 次）
* 提供 UI 选项：

  * “保留诊断文件”勾选（便于排查）
  * “导出诊断包”按钮（打包为 zip）

---

## 6.10 失败与异常处理

### 6.10.1 常见失败点

* PDF 渲染失败：文件损坏、加密、渲染库异常
* 内存不足：高 DPI 大图 + 并发过高
* 裁切参数异常：裁掉过多导致内容缺失
* deskew 角度检测失败：噪声导致角度误判

### 6.10.2 处理策略

* 渲染失败：标记该页失败，提示用户；若连续失败可终止任务
* 内存压力：提示用户降低 DPI 或并发
* 裁切异常：阻止开始并提示调整百分比；或在预览中直观看到
* deskew：若角度异常（比如 > 10°）可视为检测失败，禁用 deskew 回退原图（避免旋转毁图）

---

## 6.11 推荐接口设计（供架构实现对接）

### 6.11.1 渲染器接口

```csharp
public interface IPdfRenderer
{
    int GetPageCount(string pdfPath);
    Bitmap RenderPage(string pdfPath, int pageIndex0Based, int dpi);
}
```

### 6.11.2 预处理管线接口

```csharp
public interface IPageImagePipeline
{
    PageImageBundle Process(Bitmap renderedPage, CropOptions crop, PreprocessOptions opt);
}
```

`PageImageBundle`（示意）：

```csharp
public sealed class PageImageBundle
{
    public int PageNumber { get; init; }

    public Bitmap OriginalColor { get; init; } = default!;
    public Bitmap CroppedColor { get; init; } = default!;
    public Mat Gray { get; init; } = default!;
    public Mat Binary { get; init; } = default!;

    public Bitmap ColorForGemini { get; init; } = default!;
    public Mat BinaryForTable { get; init; } = default!;

    public CropInfo CropInfo { get; init; } = new();
    public (int W, int H) OriginalSizePx { get; init; }
    public (int W, int H) CroppedSizePx { get; init; }
}
```

---

## 6.12 本模块验收点

* 在 200/300/400 DPI 下能稳定渲染并输出页图
* 页眉/页脚裁切行为正确，裁切像素记录准确
* 输出二值图能明显区分表格线与背景（用于后续表格检测）
* deskew 开启时不会引入明显的旋转伪影或坐标错乱（必要时能回退）
* 缓存/诊断输出可复现问题（同一页可用缓存文件重跑后续模块）
