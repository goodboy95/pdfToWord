using Microsoft.Extensions.Logging;
using Pdf2Word.Core.Logging;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Validation;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO.Compression;

namespace Pdf2Word.Core.Services;

public sealed class ConversionService : IConversionService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    private readonly IPdfRenderer _renderer;
    private readonly IPageImagePipeline _pipeline;
    private readonly ITableEngine _tableEngine;
    private readonly IGeminiClient _gemini;
    private readonly IDocxWriter _docxWriter;
    private readonly ITempStorage _tempStorage;
    private readonly ILogger<ConversionService> _logger;
    private readonly ILogSink? _logSink;

    public ConversionService(IPdfRenderer renderer,
        IPageImagePipeline pipeline,
        ITableEngine tableEngine,
        IGeminiClient gemini,
        IDocxWriter docxWriter,
        ITempStorage tempStorage,
        ILogger<ConversionService> logger,
        ILogSink? logSink = null)
    {
        _renderer = renderer;
        _pipeline = pipeline;
        _tableEngine = tableEngine;
        _gemini = gemini;
        _docxWriter = docxWriter;
        _tempStorage = tempStorage;
        _logger = logger;
        _logSink = logSink;
    }

    public async Task<ConvertResult> ConvertAsync(ConvertJobRequest request, IProgress<JobProgress> progress, CancellationToken ct)
    {
        var startedAt = DateTime.UtcNow;
        var result = new ConvertResult { Status = JobStatus.Running };
        PublishLog(JobStage.Init, ErrorSeverity.Warning, null, null, "开始转换任务");

        if (string.IsNullOrWhiteSpace(request.PdfPath) || !File.Exists(request.PdfPath))
        {
            return Fail(result, "E_CFG_INVALID_PDF_PATH", "PDF 路径无效。", ErrorSeverity.Recoverable, JobStage.Init);
        }

        var layout = request.Options.Layout;
        if (layout.HeaderPercent < 0 || layout.FooterPercent < 0 || layout.HeaderPercent + layout.FooterPercent > layout.MaxCropTotalPercent)
        {
            return Fail(result, "E_CFG_CROP_TOO_LARGE", "页眉/页脚裁切比例过大。", ErrorSeverity.Recoverable, JobStage.Init);
        }

        var outputPath = request.OutputPath;
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            if (request.Options.Output.FileNameMode == OutputFileNameMode.Custom)
            {
                return Fail(result, "E_CFG_OUTPUT_PATH_REQUIRED", "未指定输出路径。", ErrorSeverity.Recoverable, JobStage.Init);
            }

            var dir = string.IsNullOrWhiteSpace(request.Options.Output.OutputDirectory)
                ? Path.GetDirectoryName(request.PdfPath) ?? string.Empty
                : request.Options.Output.OutputDirectory;
            var baseName = Path.GetFileNameWithoutExtension(request.PdfPath);
            outputPath = Path.Combine(dir, baseName + "_converted.docx");
        }

        if (!request.Options.Output.Overwrite && File.Exists(outputPath))
        {
            return Fail(result, "E_OUTPUT_EXISTS", "输出文件已存在。", ErrorSeverity.Recoverable, JobStage.Init);
        }

        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        result.OutputPath = outputPath;

        int totalPages;
        try
        {
            totalPages = _renderer.GetPageCount(request.PdfPath);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to get page count.");
            return Fail(result, "E_PDF_GET_PAGECOUNT_FAILED", "无法读取 PDF 页数。", ErrorSeverity.Fatal, JobStage.PdfOpen);
        }

        var pageRange = PageRangeParser.Parse(request.PageRangeText, totalPages);
        if (pageRange.HasError)
        {
            return Fail(result, "E_CFG_INVALID_PAGE_RANGE", pageRange.ErrorMessage ?? "页码范围格式不正确。", ErrorSeverity.Recoverable, JobStage.Init);
        }
        foreach (var warning in pageRange.Warnings)
        {
            PublishLog(JobStage.Init, ErrorSeverity.Warning, null, null, warning);
        }

        if (pageRange.Pages.Count == 0)
        {
            return Fail(result, "E_CFG_INVALID_PAGE_RANGE", "页码范围为空。", ErrorSeverity.Recoverable, JobStage.Init);
        }

        _tempStorage.EnsureCreated();
        if (request.Options.Diagnostics.KeepTempFiles)
        {
            var meta = new
            {
                startedAtUtc = startedAt,
                pdfPath = request.PdfPath,
                outputPath,
                pageRange = request.PageRangeText,
                options = request.Options
            };
            WriteJson(_tempStorage.GetIrPath("meta.json"), meta);
        }

        try
        {
            var doc = new DocumentIr
            {
                Meta = new DocumentMeta
                {
                    SourcePdfPath = request.PdfPath,
                    GeneratedAtUtc = DateTime.UtcNow,
                    Options = new ConvertOptionsSnapshot
                    {
                        Dpi = request.Options.Render.Dpi,
                        PageRange = request.PageRangeText ?? string.Empty,
                        HeaderFooterMode = request.Options.Layout.HeaderFooterMode,
                        HeaderPercent = request.Options.Layout.HeaderPercent,
                        FooterPercent = request.Options.Layout.FooterPercent,
                        PageSizeMode = request.Options.Layout.PageSizeMode
                    }
                }
            };

            var semaphore = new SemaphoreSlim(Math.Max(1, request.Options.Runtime.PageConcurrency));
            var geminiSemaphore = new SemaphoreSlim(Math.Max(1, request.Options.Runtime.GeminiConcurrency));
            var pageResults = new List<PageIr>();
            var failures = new System.Collections.Concurrent.ConcurrentBag<FailureInfo>();
            var completed = 0;
            var failFastTriggered = false;
            using var failFastCts = request.Options.Validation.FailFast
                ? CancellationTokenSource.CreateLinkedTokenSource(ct)
                : null;
            var effectiveToken = failFastCts?.Token ?? ct;

            var tasks = pageRange.Pages.Select(async pageNumber =>
            {
                await semaphore.WaitAsync(effectiveToken).ConfigureAwait(false);
                try
                {
                    var pageIr = await ProcessPageAsync(request, pageNumber, totalPages, progress, geminiSemaphore, failures, effectiveToken).ConfigureAwait(false);
                    if (pageIr != null)
                    {
                        lock (pageResults)
                        {
                            pageResults.Add(pageIr);
                        }
                    }
                }
                finally
                {
                    semaphore.Release();
                    var done = Interlocked.Increment(ref completed);
                    progress.Report(new JobProgress
                    {
                        TotalPages = pageRange.Pages.Count,
                        CompletedPages = done,
                        CurrentPage = pageNumber,
                        Stage = JobStage.Finalize,
                        Message = "页面处理完成"
                    });
                }

                if (request.Options.Validation.FailFast && !failures.IsEmpty)
                {
                    failFastTriggered = true;
                    failFastCts?.Cancel();
                }
            }).ToList();

            try
            {
                await Task.WhenAll(tasks).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                result.Status = failFastTriggered ? JobStatus.Failed : JobStatus.Canceled;
                result.Failures = failures.ToList();
                result.Elapsed = DateTime.UtcNow - startedAt;
                return result;
            }

            var failureList = failures.ToList();

            if (pageResults.Count == 0)
            {
                return Fail(result, "E_PAGE_ALL_FAILED", "全部页面处理失败。", ErrorSeverity.Fatal, JobStage.Finalize, failureList);
            }

            if (!request.Options.Validation.AllowSkipFailedPages && failureList.Count > 0)
            {
                return Fail(result, "E_PAGE_FAILED", "存在失败页面，已中止输出。", ErrorSeverity.Fatal, JobStage.Finalize, failureList);
            }

            doc.Pages = pageResults.OrderBy(p => p.PageNumber).ToList();
            if (request.Options.Diagnostics.KeepTempFiles)
            {
                WriteJson(_tempStorage.GetIrPath("doc_ir.json"), doc);
            }

            try
            {
                await using var fs = File.Create(outputPath);
                var writeOptions = request.Options.Docx;
                writeOptions.Dpi = request.Options.Render.Dpi;
                writeOptions.PageSizeMode = request.Options.Layout.PageSizeMode;
                writeOptions.MarginLeftTwips = request.Options.Layout.MarginLeftTwips;
                writeOptions.MarginRightTwips = request.Options.Layout.MarginRightTwips;
                writeOptions.MarginTopTwips = request.Options.Layout.MarginTopTwips;
                writeOptions.MarginBottomTwips = request.Options.Layout.MarginBottomTwips;
                await _docxWriter.WriteAsync(doc, writeOptions, fs, ct).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to write docx.");
                return Fail(result, "E_DOCX_WRITE_FAILED", "无法写入 Word 文件。", ErrorSeverity.Fatal, JobStage.DocxWrite, failureList);
            }

            result.Status = failureList.Any(f => f.Severity == ErrorSeverity.Fatal) ? JobStatus.Failed : JobStatus.Succeeded;
            result.Failures = failureList;
            result.Elapsed = DateTime.UtcNow - startedAt;
            PublishLog(JobStage.Finalize, ErrorSeverity.Warning, null, null, "任务结束");
            if (request.Options.Diagnostics.KeepTempFiles && request.Options.Diagnostics.ExportZip)
            {
                TryExportZip(_tempStorage.JobRoot);
            }
            return result;
        }
        finally
        {
            if (!request.Options.Diagnostics.KeepTempFiles)
            {
                _tempStorage.Cleanup();
            }
        }
    }

    private async Task<PageIr?> ProcessPageAsync(ConvertJobRequest request, int pageNumber, int totalPages, IProgress<JobProgress> progress,
        SemaphoreSlim geminiSemaphore, System.Collections.Concurrent.ConcurrentBag<FailureInfo> failures, CancellationToken ct)
    {
        var validationEnabled = request.Options.Validation.Enable;
        progress.Report(new JobProgress { TotalPages = totalPages, CompletedPages = 0, CurrentPage = pageNumber, Stage = JobStage.PdfRender, Message = "渲染页面" });
        System.Drawing.Bitmap rendered;
        try
        {
            rendered = _renderer.RenderPage(request.PdfPath, pageNumber - 1, request.Options.Render.Dpi);
            PublishLog(JobStage.PdfRender, ErrorSeverity.Warning, pageNumber, null, "页面渲染完成");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Render page {Page}", pageNumber);
            failures.Add(new FailureInfo { PageNumber = pageNumber, ErrorCode = "E_PDF_RENDER_FAILED", Message = "页面渲染失败", Severity = ErrorSeverity.Recoverable, Stage = JobStage.PdfRender, Attempt = 1 });
            return null;
        }

        ct.ThrowIfCancellationRequested();

        progress.Report(new JobProgress { TotalPages = totalPages, CompletedPages = 0, CurrentPage = pageNumber, Stage = JobStage.Preprocess, Message = "图像预处理" });
        PageImageBundle bundle;
        try
        {
            bundle = _pipeline.Process(rendered, new CropOptions
            {
                Mode = request.Options.Layout.HeaderFooterMode,
                HeaderPercent = request.Options.Layout.HeaderPercent,
                FooterPercent = request.Options.Layout.FooterPercent
            }, request.Options.Preprocess, pageNumber);
            PublishLog(JobStage.Preprocess, ErrorSeverity.Warning, pageNumber, null, "图像预处理完成");
            if (request.Options.Diagnostics.KeepTempFiles)
            {
                SaveBitmap(bundle.OriginalColor, _tempStorage.GetPageImagePath(pageNumber, "original"));
                SaveBitmap(bundle.CroppedColor, _tempStorage.GetPageImagePath(pageNumber, "cropped"));
                SaveBitmap(bundle.BinaryForTable, _tempStorage.GetPageImagePath(pageNumber, "binary"));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Preprocess page {Page}", pageNumber);
            failures.Add(new FailureInfo { PageNumber = pageNumber, ErrorCode = "E_IMG_PREPROCESS_FAILED", Message = "图像预处理失败", Severity = ErrorSeverity.Recoverable, Stage = JobStage.Preprocess, Attempt = 1 });
            return null;
        }

        try
        {
            var tableBlocks = new List<TableBlockIr>();
            var fallbackParagraphs = new List<ParagraphDto>();

        IReadOnlyList<TableDetection> tables = Array.Empty<TableDetection>();
        if (request.Options.TableDetect.Enable)
        {
            progress.Report(new JobProgress { TotalPages = totalPages, CurrentPage = pageNumber, Stage = JobStage.TableDetect, Message = "表格检测" });
            try
            {
                tables = _tableEngine.DetectTables(bundle, request.Options.TableDetect);
                PublishLog(JobStage.TableDetect, ErrorSeverity.Warning, pageNumber, null, $"检测到表格 {tables.Count} 个");
                if (request.Options.Diagnostics.KeepTempFiles)
                {
                    var tableIdx = 0;
                    foreach (var table in tables)
                    {
                        SaveBitmap(table.TableImageColor, _tempStorage.GetTableImagePath(pageNumber, tableIdx, "color"));
                        SaveBitmap(table.TableBinary, _tempStorage.GetTableImagePath(pageNumber, tableIdx, "binary"));
                        tableIdx++;
                    }

                    if (tables.Count > 0)
                    {
                        var snapshot = tables.Select(t => new
                        {
                            t.DebugId,
                            t.TableBBoxInPage,
                            t.NCols,
                            t.NRows,
                            t.ColLinesX,
                            t.RowLinesY,
                            Cells = t.Cells.Select(c => new
                            {
                                c.Id,
                                c.Row,
                                c.Col,
                                c.Rowspan,
                                c.Colspan,
                                c.BBoxInTable
                            })
                        });
                        WriteJson(_tempStorage.GetIrPath($"p{pageNumber:000}_tables.json"), snapshot);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Table detection failed on page {Page}", pageNumber);
                failures.Add(new FailureInfo { PageNumber = pageNumber, ErrorCode = "E_TABLE_DETECT_FAILED", Message = "表格检测失败", Severity = ErrorSeverity.Recoverable, Stage = JobStage.TableDetect, Attempt = 1 });
            }
        }

        var tableIndex = 0;
        foreach (var table in tables)
        {
            progress.Report(new JobProgress { TotalPages = totalPages, CurrentPage = pageNumber, Stage = JobStage.TableGrid, Message = "表格结构" });
            var tableIr = TableIrBuilder.Build(table);
            var validation = validationEnabled ? TableStructureValidator.Validate(tableIr) : new TableStructureValidationResult { IsValid = true };
            if (validationEnabled && !validation.IsValid)
            {
                failures.Add(new FailureInfo { PageNumber = pageNumber, TableIndex = tableIndex, ErrorCode = validation.ErrorCode ?? "E_TABLE_GRID_CONFLICT", Message = validation.Message ?? "表格结构冲突", Severity = ErrorSeverity.Recoverable, Stage = JobStage.TableGrid, Attempt = 1 });
                if (request.Options.Validation.TableStructureBad == TableFallbackPolicy.FallbackTextTable && request.Options.TableDetect.EnableTextTableFallback)
                {
                    var lines = await GeminiTableAsLinesAsync(request, table, geminiSemaphore, failures, pageNumber, tableIndex, ct).ConfigureAwait(false);
                    if (lines.Count > 0)
                    {
                        fallbackParagraphs.Add(new ParagraphDto { Role = "body", Text = string.Join("\n", lines) });
                    }
                }

                tableIndex++;
                continue;
            }

            progress.Report(new JobProgress { TotalPages = totalPages, CurrentPage = pageNumber, Stage = JobStage.GeminiTableOcr, Message = "表格文字识别" });
            var ocrTexts = await GeminiTableCellsAsync(request, table, geminiSemaphore, failures, pageNumber, tableIndex, ct).ConfigureAwait(false);
            if (ocrTexts == null)
            {
                if (request.Options.Validation.TableStructureOkTextBad == TableFallbackPolicy.FallbackTextTable && request.Options.TableDetect.EnableTextTableFallback)
                {
                    var lines = await GeminiTableAsLinesAsync(request, table, geminiSemaphore, failures, pageNumber, tableIndex, ct).ConfigureAwait(false);
                    if (lines.Count > 0)
                    {
                        fallbackParagraphs.Add(new ParagraphDto { Role = "body", Text = string.Join("\n", lines) });
                        tableIndex++;
                        continue;
                    }
                }

                tableIr = TableIrBuilder.Build(table, null);
            }
            else
            {
                tableIr = TableIrBuilder.Build(table, ocrTexts);
            }

            tableBlocks.Add(tableIr);
            tableIndex++;
        }

        DisposeTableImages(tables);

        progress.Report(new JobProgress { TotalPages = totalPages, CurrentPage = pageNumber, Stage = JobStage.GeminiPageOcr, Message = "正文识别" });
        var masked = _tableEngine.MaskTables(bundle.ColorForGemini, tables.Select(t => t.TableBBoxInPage));
        if (request.Options.Diagnostics.KeepTempFiles)
        {
            SaveBitmap(masked, _tempStorage.GetPageImagePath(pageNumber, "masked"));
        }
        var paragraphs = await GeminiPageParagraphsAsync(request, masked, geminiSemaphore, failures, pageNumber, ct).ConfigureAwait(false);
        masked.Dispose();
        if (fallbackParagraphs.Count > 0)
        {
            paragraphs.InsertRange(0, fallbackParagraphs);
        }

        var sanitizedParagraphs = paragraphs.Select(p => new ParagraphDto
        {
            Role = p.Role,
            Text = TextCleaner.RemoveControlChars(p.Text)
        }).ToList();

        var pageIr = IrBuilder.BuildPage(pageNumber, bundle.OriginalSizePx, bundle.CroppedSizePx, bundle.CropInfo, sanitizedParagraphs, tableBlocks, "GeminiText");
        if (pageIr.Blocks.Count == 0 && validationEnabled)
        {
            failures.Add(new FailureInfo
            {
                PageNumber = pageNumber,
                ErrorCode = "E_PAGE_EMPTY",
                Message = "页面未识别到任何内容",
                Severity = ErrorSeverity.Recoverable,
                Stage = JobStage.GeminiPageOcr,
                Attempt = 1
            });
            return null;
        }

        if (request.Options.Diagnostics.KeepTempFiles)
        {
            WriteJson(_tempStorage.GetIrPath($"p{pageNumber:000}_ir.json"), pageIr);
        }
        return pageIr;
        }
        finally
        {
            DisposeBundle(bundle);
        }
    }

    private async Task<Dictionary<string, string>?> GeminiTableCellsAsync(ConvertJobRequest request, TableDetection table, SemaphoreSlim geminiSemaphore, System.Collections.Concurrent.ConcurrentBag<FailureInfo> failures, int pageNumber, int tableIndex, CancellationToken ct)
    {
        var cells = table.Cells.Select(c => new CellBoxForOcr
        {
            Id = c.Id,
            X = c.BBoxInTable.X,
            Y = c.BBoxInTable.Y,
            W = c.BBoxInTable.W,
            H = c.BBoxInTable.H
        }).ToList();

        if (cells.Count == 0)
        {
            return new Dictionary<string, string>();
        }

        var maxRetries = Math.Max(0, request.Options.Gemini.MaxRetryCount);
        var validationEnabled = request.Options.Validation.Enable;
        for (var attempt = 0; attempt <= maxRetries; attempt++)
        {
            await geminiSemaphore.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var strict = attempt > 0;
                var result = await _gemini.RecognizeTableCellsAsync(table.TableImageColor, cells, ct, strict, attempt + 1).ConfigureAwait(false);
                WriteRawJsonIfNeeded(request, _tempStorage.GetIrPath($"p{pageNumber:000}_t{tableIndex:00}_cells_a{attempt + 1}.json"), result.RawJson);
                var cleaned = new Dictionary<string, string>();
                foreach (var kvp in result.TextById)
                {
                    cleaned[kvp.Key] = TextCleaner.RemoveControlChars(kvp.Value ?? string.Empty);
                }
                if (!validationEnabled)
                {
                    return cleaned;
                }

                var missingRate = MissingRate(cells.Select(c => c.Id), result.TextById.Keys);
                var emptyCount = cleaned.Values.Count(v => string.IsNullOrWhiteSpace(v));
                var emptyRate = cells.Count == 0 ? 0 : (double)emptyCount / cells.Count;
                if (missingRate <= request.Options.Gemini.TableMissingIdThreshold && emptyRate <= request.Options.Gemini.TableEmptyTextRateThreshold)
                {
                    return cleaned;
                }

                var errorCode = missingRate > request.Options.Gemini.TableMissingIdThreshold
                    ? "E_GEMINI_OCR_MISSING_IDS"
                    : "E_GEMINI_OCR_EMPTY_TEXT";
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    TableIndex = tableIndex,
                    ErrorCode = errorCode,
                    Message = "表格识别结果不完整",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiTableOcr,
                    Attempt = attempt + 1
                });
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Gemini table OCR returned invalid JSON");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    TableIndex = tableIndex,
                    ErrorCode = "E_GEMINI_JSON_INVALID",
                    Message = "表格识别返回非 JSON",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiTableOcr,
                    Attempt = attempt + 1
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini table OCR failed");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    TableIndex = tableIndex,
                    ErrorCode = "E_GEMINI_TIMEOUT",
                    Message = "表格识别失败",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiTableOcr,
                    Attempt = attempt + 1
                });
            }
            finally
            {
                geminiSemaphore.Release();
            }

            if (attempt < maxRetries)
            {
                await DelayWithBackoffAsync(request, attempt, ct).ConfigureAwait(false);
            }
        }

        return null;
    }

    private async Task<List<string>> GeminiTableAsLinesAsync(ConvertJobRequest request, TableDetection table, SemaphoreSlim geminiSemaphore, System.Collections.Concurrent.ConcurrentBag<FailureInfo> failures, int pageNumber, int tableIndex, CancellationToken ct)
    {
        var maxRetries = Math.Max(0, request.Options.Gemini.MaxRetryCount);
        for (var attempt = 0; attempt <= maxRetries; attempt++)
        {
            await geminiSemaphore.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var result = await _gemini.RecognizeTableAsLinesAsync(table.TableImageColor, ct, attempt + 1).ConfigureAwait(false);
                WriteRawJsonIfNeeded(request, _tempStorage.GetIrPath($"p{pageNumber:000}_t{tableIndex:00}_lines_a{attempt + 1}.json"), result.RawJson);
                if (result.Lines.Count > 0)
                {
                    return result.Lines.Select(TextCleaner.RemoveControlChars).ToList();
                }
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Gemini table fallback returned invalid JSON");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    TableIndex = tableIndex,
                    ErrorCode = "E_GEMINI_JSON_INVALID",
                    Message = "表格降级返回非 JSON",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiTableOcr,
                    Attempt = attempt + 1
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini table fallback failed");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    TableIndex = tableIndex,
                    ErrorCode = "E_GEMINI_TIMEOUT",
                    Message = "表格降级识别失败",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiTableOcr,
                    Attempt = attempt + 1
                });
            }
            finally
            {
                geminiSemaphore.Release();
            }

            if (attempt < maxRetries)
            {
                await DelayWithBackoffAsync(request, attempt, ct).ConfigureAwait(false);
            }
        }

        return new List<string>();
    }

    private async Task<List<ParagraphDto>> GeminiPageParagraphsAsync(ConvertJobRequest request, System.Drawing.Bitmap image, SemaphoreSlim geminiSemaphore, System.Collections.Concurrent.ConcurrentBag<FailureInfo> failures, int pageNumber, CancellationToken ct)
    {
        var maxRetries = Math.Max(0, request.Options.Gemini.MaxRetryCount);
        var validationEnabled = request.Options.Validation.Enable;
        for (var attempt = 0; attempt <= maxRetries; attempt++)
        {
            await geminiSemaphore.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var strict = attempt > 0;
                var result = await _gemini.RecognizePageParagraphsAsync(image, ct, strict, attempt + 1).ConfigureAwait(false);
                WriteRawJsonIfNeeded(request, _tempStorage.GetIrPath($"p{pageNumber:000}_page_a{attempt + 1}.json"), result.RawJson);
                if (!validationEnabled)
                {
                    return result.Paragraphs;
                }

                var totalChars = result.Paragraphs.Sum(p => p.Text?.Length ?? 0);
                if (result.Paragraphs.Count > 0 && totalChars >= request.Options.Gemini.MinPageTextCharCount)
                {
                    return result.Paragraphs;
                }

                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    ErrorCode = "E_GEMINI_OCR_EMPTY_TEXT",
                    Message = "正文识别为空",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiPageOcr,
                    Attempt = attempt + 1
                });
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Gemini page OCR returned invalid JSON");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    ErrorCode = "E_GEMINI_JSON_INVALID",
                    Message = "正文识别返回非 JSON",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiPageOcr,
                    Attempt = attempt + 1
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini page OCR failed");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    ErrorCode = "E_GEMINI_TIMEOUT",
                    Message = "正文识别失败",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiPageOcr,
                    Attempt = attempt + 1
                });
            }
            finally
            {
                geminiSemaphore.Release();
            }

            if (attempt < maxRetries)
            {
                await DelayWithBackoffAsync(request, attempt, ct).ConfigureAwait(false);
            }
        }

        if (validationEnabled && request.Options.Validation.TextEmpty == TextFallbackPolicy.SingleParagraph)
        {
            return await GeminiPageSingleParagraphAsync(request, image, geminiSemaphore, failures, pageNumber, ct).ConfigureAwait(false);
        }

        return new List<ParagraphDto>();
    }

    private async Task<List<ParagraphDto>> GeminiPageSingleParagraphAsync(ConvertJobRequest request, System.Drawing.Bitmap image, SemaphoreSlim geminiSemaphore, System.Collections.Concurrent.ConcurrentBag<FailureInfo> failures, int pageNumber, CancellationToken ct)
    {
        var maxRetries = Math.Max(0, request.Options.Gemini.MaxRetryCount);
        for (var attempt = 0; attempt <= maxRetries; attempt++)
        {
            await geminiSemaphore.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                var result = await _gemini.RecognizePageAsSingleTextAsync(image, ct, attempt + 1).ConfigureAwait(false);
                WriteRawJsonIfNeeded(request, _tempStorage.GetIrPath($"p{pageNumber:000}_page_fallback_a{attempt + 1}.json"), result.RawJson);
                if (!string.IsNullOrWhiteSpace(result.Text))
                {
                    return new List<ParagraphDto> { new ParagraphDto { Role = "body", Text = TextCleaner.RemoveControlChars(result.Text) } };
                }
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, "Gemini page fallback returned invalid JSON");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    ErrorCode = "E_GEMINI_JSON_INVALID",
                    Message = "正文降级返回非 JSON",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiPageOcr,
                    Attempt = attempt + 1
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Gemini page fallback failed");
                failures.Add(new FailureInfo
                {
                    PageNumber = pageNumber,
                    ErrorCode = "E_GEMINI_TIMEOUT",
                    Message = "正文降级识别失败",
                    Severity = ErrorSeverity.Recoverable,
                    Stage = JobStage.GeminiPageOcr,
                    Attempt = attempt + 1
                });
            }
            finally
            {
                geminiSemaphore.Release();
            }

            if (attempt < maxRetries)
            {
                await DelayWithBackoffAsync(request, attempt, ct).ConfigureAwait(false);
            }
        }

        return new List<ParagraphDto>();
    }

    private static double MissingRate(IEnumerable<string> expected, IEnumerable<string> actual)
    {
        var expectedSet = new HashSet<string>(expected);
        var expectedCount = expectedSet.Count;
        var actualSet = new HashSet<string>(actual);
        if (expectedCount == 0)
        {
            return 0;
        }

        expectedSet.ExceptWith(actualSet);
        return (double)expectedSet.Count / expectedCount;
    }

    private ConvertResult Fail(ConvertResult result, string code, string message, ErrorSeverity severity, JobStage stage, List<FailureInfo>? failures = null)
    {
        result.Status = JobStatus.Failed;
        var info = new FailureInfo { ErrorCode = code, Message = message, Severity = severity, Stage = stage, Attempt = 1 };
        result.Failures = failures ?? new List<FailureInfo>();
        result.Failures.Add(info);
        PublishLog(stage, severity, info.PageNumber, info.TableIndex, message, code);
        return result;
    }

    private void PublishLog(JobStage stage, ErrorSeverity severity, int? pageNumber, int? tableIndex, string message, string? errorCode = null, int attempt = 1)
    {
        _logSink?.Publish(new LogEntry
        {
            Stage = stage,
            Severity = severity,
            PageNumber = pageNumber,
            TableIndex = tableIndex,
            Message = message,
            ErrorCode = errorCode ?? string.Empty,
            Attempt = attempt
        });
    }

    private static void DisposeTableImages(IEnumerable<TableDetection> tables)
    {
        foreach (var table in tables)
        {
            table.TableImageColor.Dispose();
            table.TableBinary.Dispose();
        }
    }

    private static void DisposeBundle(PageImageBundle bundle)
    {
        bundle.OriginalColor.Dispose();
        bundle.CroppedColor.Dispose();
        bundle.Gray.Dispose();
        bundle.ColorForGemini.Dispose();
        if (!ReferenceEquals(bundle.Binary, bundle.BinaryForTable))
        {
            bundle.BinaryForTable.Dispose();
        }
        bundle.Binary.Dispose();
    }

    private static Task DelayWithBackoffAsync(ConvertJobRequest request, int attempt, CancellationToken ct)
    {
        var baseMs = Math.Max(0, request.Options.Gemini.BackoffBaseMs);
        if (baseMs == 0)
        {
            return Task.CompletedTask;
        }

        var delay = baseMs * Math.Pow(2, attempt);
        delay = Math.Min(delay, 5000);
        return Task.Delay((int)delay, ct);
    }

    private void WriteRawJsonIfNeeded(ConvertJobRequest request, string path, string? rawJson)
    {
        if (!request.Options.Diagnostics.SaveRawGeminiJson || string.IsNullOrWhiteSpace(rawJson))
        {
            return;
        }

        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);
        File.WriteAllText(path, rawJson);
    }

    private static void SaveBitmap(System.Drawing.Bitmap bitmap, string path)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);
        bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
    }

    private static void WriteJson<T>(string path, T payload)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? string.Empty);
        var json = JsonSerializer.Serialize(payload, JsonOptions);
        File.WriteAllText(path, json);
    }

    private static void TryExportZip(string jobRoot)
    {
        try
        {
            var zipPath = jobRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + ".zip";
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }
            ZipFile.CreateFromDirectory(jobRoot, zipPath, CompressionLevel.Optimal, false);
        }
        catch
        {
            // zip export is best-effort
        }
    }
}
