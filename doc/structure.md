# 项目结构说明

> 仅包含开发者编写的目录与代码文件（不包含 bin/obj 等构建产物或外部依赖目录）。

## 根目录
- AGENTS.md：开发约束与执行要求。
- Pdf2Word.sln：解决方案入口，包含 App/Core/Infrastructure/Tests。
- dev-doc/：项目整体开发方案与规范文档。
- doc/structure.md：当前文件，记录代码结构与职责说明。
- src/：应用源码。
- tests/：单元测试源码。

## dev-doc/
- 00_总览与快速开始.md：目标、里程碑、流程概览。
- 01_需求与范围.md：功能/非功能需求与范围边界。
- 02_架构设计.md：模块划分、接口与流程。
- 03_数据结构与IR规范.md：IR 结构与校验规范。
- 04_用户交互与UI设计.md：WPF 交互与界面规范。
- 05_页码范围解析规范.md：页码范围解析规则与测试建议。
- 06_PDF渲染与图像预处理.md：渲染/裁切/预处理流程。
- 07_表格检测与单元格分割算法.md：OpenCV 表格检测与合并单元格推断。
- 08_Gemini集成与提示词.md：Gemini 调用、prompt 与 JSON 协议。
- 09_结果校验-重试-降级策略.md：校验、重试与降级策略。
- 10_Word生成-OpenXML实现.md：OpenXML 写入策略。
- 11_版式还原策略-页眉页脚-页面尺寸.md：版式与页面尺寸策略。
- 12_工程与运维.md：性能、日志、测试、部署、隐私规范。
- Appendix/A1_配置项与默认值表.md：配置项默认值与范围。
- Appendix/A2_错误码与提示文案.md：错误码与用户提示文案。
- Appendix/A3_示例输入输出与样例JSON.md：示例输入输出（预留）。
- Appendix/A4_第三方依赖与许可证清单.md：第三方依赖与许可证要求。

## src/Pdf2Word.App/
- Pdf2Word.App.csproj：WPF 应用项目文件。
- App.xaml：应用资源与样式。
- App.xaml.cs：启动入口与依赖注入配置。
- MainWindow.xaml：主界面布局与控件绑定。
- MainWindow.xaml.cs：主窗体代码隐藏，设置 DataContext。
- appsettings.json：默认配置项。
- Models/PreviewOverlay.cs：预览叠加层矩形模型（裁切遮罩/表格框）。
- Services/DpapiApiKeyStore.cs：使用 DPAPI 加密保存 Gemini API Key。
- Services/BitmapSourceHelper.cs：System.Drawing.Bitmap 转 WPF ImageSource 辅助。
- Services/CompositeLogSink.cs：组合日志写入（UI + 文件）。
- Services/UiLogSink.cs：UI 日志接收器。
- ViewModels/MainViewModel.cs：主界面 MVVM 逻辑与命令。

## src/Pdf2Word.Core/
- Pdf2Word.Core.csproj：核心库项目文件。
- Models/BBox.cs：矩形坐标结构。
- Models/Enums.cs：枚举定义（页眉页脚、页面尺寸、状态等）。
- Models/ConvertJobRequest.cs：任务请求与结果模型。
- Models/JobProgress.cs：进度上报模型。
- Models/PageImageBundle.cs：渲染/预处理后的页面图像集合。
- Models/TableDetection.cs：表格检测结果结构。
- Models/Ir/DocumentIr.cs：Document IR 结构定义。
- Models/Ir/PageIr.cs：Page/Block/Table IR 结构定义。
- Logging/LogEntry.cs：日志结构与日志发布接口。
- Options/AppOptions.cs：全量配置项与默认值。
- Services/Interfaces.cs：核心接口定义（渲染、表格、Gemini、Docx 等）。
- Services/PageRangeParser.cs：页码范围解析实现。
- Services/TextCleaner.cs：文本清洗与乱码统计。
- Services/TwipsConverter.cs：像素到 twips 的换算。
- Services/TableIrBuilder.cs：表格检测结果转 Table IR。
- Services/IrBuilder.cs：段落/表格合并为 Page IR 的构建器。
- Services/ConversionService.cs：转换流程调度与重试/降级逻辑。
- Validation/TableStructureValidator.cs：表格结构校验（owner matrix）。

## src/Pdf2Word.Infrastructure/
- Pdf2Word.Infrastructure.csproj：基础设施实现项目文件。
- Pdf/PdfiumRenderer.cs：PDFium 渲染器实现。
- ImageProcessing/OpenCvPageImagePipeline.cs：图像裁切/预处理管线。
- Table/OpenCvTableEngine.cs：表格检测与单元格分割实现。
- Gemini/GeminiClient.cs：Gemini API 调用与响应解析。
- Gemini/GeminiPrompts.cs：Gemini Prompt 模板。
- Gemini/ImageEncoder.cs：图像编码与压缩。
- Docx/OpenXmlDocxWriter.cs：OpenXML 写入 docx 实现。
- Logging/FileLogSink.cs：诊断日志写入临时文件。
- Storage/TempStorage.cs：临时目录与诊断文件管理。

## tests/Pdf2Word.Tests/
- Pdf2Word.Tests.csproj：测试项目文件。
- Tests/PageRangeParserTests.cs：页码范围解析测试。
- Tests/TextCleanerTests.cs：文本清洗测试。
- Tests/TwipsConverterTests.cs：像素→twips 换算测试。
- Tests/TableStructureValidatorTests.cs：表格结构校验测试。
