using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Pdf2Word.App.Models;
using Pdf2Word.App.Services;
using Pdf2Word.Core.Logging;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Models.Ir;
using Pdf2Word.Core.Options;
using CoreRenderOptions = Pdf2Word.Core.Options.RenderOptions;
using Pdf2Word.Core.Services;

namespace Pdf2Word.App.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private const double MaxCropPercentPerEdge = 0.2;
    private static readonly Brush CropOverlayBrush = new SolidColorBrush(Color.FromArgb(90, 46, 110, 247));
    private static readonly Brush TableStrokeBrush = new SolidColorBrush(Color.FromArgb(220, 255, 77, 79));

    private readonly IConversionService _conversionService;
    private readonly IPdfRenderer _pdfRenderer;
    private readonly IPageImagePipeline _imagePipeline;
    private readonly ITableEngine _tableEngine;
    private readonly AppOptions _defaults;
    private readonly IApiKeyStore _apiKeyStore;

    private CancellationTokenSource? _cts;
    private CancellationTokenSource? _previewCts;

    private ImageSource? _previewOriginalImage;
    private ImageSource? _previewCroppedImage;
    private (int W, int H) _previewOriginalSize;
    private (int W, int H) _previewCroppedSize;
    private CropInfo? _previewCropInfo;
    private List<BBox> _previewTableBoxes = new();

    public ObservableCollection<int> DpiOptions { get; } = new() { 200, 300, 400 };
    public ObservableCollection<int> ConcurrencyOptions { get; } = new() { 1, 2, 4 };
    public ObservableCollection<HeaderFooterRemoveMode> HeaderFooterModes { get; } = new(Enum.GetValues<HeaderFooterRemoveMode>());
    public ObservableCollection<PageSizeMode> PageSizeModes { get; } = new(Enum.GetValues<PageSizeMode>());

    public ObservableCollection<string> Logs { get; } = new();
    public ObservableCollection<string> Failures { get; } = new();
    public ObservableCollection<PreviewOverlay> PreviewOverlays { get; } = new();

    [ObservableProperty]
    private string _pdfPath = string.Empty;

    [ObservableProperty]
    private string _pageRangeText = string.Empty;

    [ObservableProperty]
    private HeaderFooterRemoveMode _headerFooterMode;

    [ObservableProperty]
    private double _headerPercent;

    [ObservableProperty]
    private double _footerPercent;

    [ObservableProperty]
    private PageSizeMode _pageSizeMode;

    [ObservableProperty]
    private int _dpi;

    [ObservableProperty]
    private int _concurrency;

    [ObservableProperty]
    private bool _deskewEnabled;

    [ObservableProperty]
    private string _outputPath = string.Empty;

    [ObservableProperty]
    private bool _keepDiagnostics;

    [ObservableProperty]
    private string _geminiEndpointUrl = string.Empty;

    [ObservableProperty]
    private string _geminiApiKey = string.Empty;

    [ObservableProperty]
    private string _geminiEndpointError = string.Empty;

    [ObservableProperty]
    private bool _hasGeminiEndpointError;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressPercent;

    [ObservableProperty]
    private string _progressText = string.Empty;

    [ObservableProperty]
    private int _pdfPageCount;

    [ObservableProperty]
    private string _pdfInfoText = string.Empty;

    [ObservableProperty]
    private string _pageRangeSummary = string.Empty;

    [ObservableProperty]
    private string _pageRangeError = string.Empty;

    [ObservableProperty]
    private bool _hasPageRangeError;

    [ObservableProperty]
    private string _cropError = string.Empty;

    [ObservableProperty]
    private bool _hasCropError;

    [ObservableProperty]
    private bool _isHeaderEnabled;

    [ObservableProperty]
    private bool _isFooterEnabled;

    [ObservableProperty]
    private bool _canOpenOutput;

    [ObservableProperty]
    private bool _canOpenOutputFolder;

    [ObservableProperty]
    private bool _isPreviewBusy;

    [ObservableProperty]
    private int _previewPageNumber = 1;

    [ObservableProperty]
    private string _previewStatus = string.Empty;

    [ObservableProperty]
    private ImageSource? _previewImage;

    [ObservableProperty]
    private int _previewImageWidth = 1;

    [ObservableProperty]
    private int _previewImageHeight = 1;

    [ObservableProperty]
    private bool _previewShowCropped;

    [ObservableProperty]
    private bool _previewShowTables;

    [ObservableProperty]
    private bool _previewShowCropOverlay = true;

    public bool HasInputErrors => HasPageRangeError || HasCropError || HasGeminiEndpointError;

    public bool CanStart => !IsBusy && !string.IsNullOrWhiteSpace(PdfPath) && !HasInputErrors;

    public MainViewModel(IConversionService conversionService,
        IPdfRenderer pdfRenderer,
        IPageImagePipeline imagePipeline,
        ITableEngine tableEngine,
        AppOptions defaults,
        IApiKeyStore apiKeyStore,
        UiLogSink logSink)
    {
        _conversionService = conversionService;
        _pdfRenderer = pdfRenderer;
        _imagePipeline = imagePipeline;
        _tableEngine = tableEngine;
        _defaults = defaults;
        _apiKeyStore = apiKeyStore;

        HeaderFooterMode = defaults.Layout.HeaderFooterMode;
        HeaderPercent = defaults.Layout.HeaderPercent;
        FooterPercent = defaults.Layout.FooterPercent;
        PageSizeMode = defaults.Layout.PageSizeMode;
        Dpi = defaults.Render.Dpi;
        Concurrency = defaults.Runtime.PageConcurrency;
        DeskewEnabled = defaults.Preprocess.EnableDeskew;
        KeepDiagnostics = defaults.Diagnostics.KeepTempFiles;
        GeminiEndpointUrl = defaults.Gemini.Endpoint;
        GeminiApiKey = apiKeyStore.GetApiKey() ?? string.Empty;
        PreviewShowCropped = false;
        PreviewShowTables = false;

        UpdateHeaderFooterEnabled();
        ValidateCrop();
        ValidateGeminiEndpoint();

        logSink.LogPublished += OnLogPublished;
    }

    partial void OnPdfPathChanged(string value)
    {
        OnPropertyChanged(nameof(CanStart));
        _ = LoadPdfInfoAsync();
        UpdateOutputAvailability();
    }

    partial void OnPageRangeTextChanged(string value) => UpdatePageRangeSummary();

    partial void OnHeaderFooterModeChanged(HeaderFooterRemoveMode value) => UpdateHeaderFooterEnabled();

    partial void OnHeaderPercentChanged(double value) => ValidateCrop();

    partial void OnFooterPercentChanged(double value) => ValidateCrop();

    partial void OnIsBusyChanged(bool value) => OnPropertyChanged(nameof(CanStart));

    partial void OnHasPageRangeErrorChanged(bool value) => OnPropertyChanged(nameof(CanStart));

    partial void OnHasCropErrorChanged(bool value) => OnPropertyChanged(nameof(CanStart));

    partial void OnGeminiEndpointUrlChanged(string value) => ValidateGeminiEndpoint();

    partial void OnHasGeminiEndpointErrorChanged(bool value) => OnPropertyChanged(nameof(CanStart));

    partial void OnOutputPathChanged(string value) => UpdateOutputAvailability();

    partial void OnPreviewShowCroppedChanged(bool value) => UpdatePreviewImage();

    partial void OnPreviewShowTablesChanged(bool value) => UpdatePreviewOverlays();

    partial void OnPreviewShowCropOverlayChanged(bool value) => UpdatePreviewOverlays();

    [RelayCommand]
    private void BrowsePdf()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "PDF Files (*.pdf)|*.pdf",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            PdfPath = dialog.FileName;
            if (string.IsNullOrWhiteSpace(OutputPath))
            {
                var dir = Path.GetDirectoryName(PdfPath) ?? string.Empty;
                var name = Path.GetFileNameWithoutExtension(PdfPath) + "_converted.docx";
                OutputPath = Path.Combine(dir, name);
            }
        }
    }

    [RelayCommand]
    private void BrowseOutput()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "Word Document (*.docx)|*.docx",
            FileName = string.IsNullOrWhiteSpace(OutputPath) ? "output.docx" : Path.GetFileName(OutputPath)
        };

        if (dialog.ShowDialog() == true)
        {
            OutputPath = dialog.FileName;
        }
    }

    [RelayCommand]
    private async Task StartAsync()
    {
        if (!CanStart)
        {
            return;
        }

        IsBusy = true;
        Logs.Clear();
        Failures.Clear();
        ProgressPercent = 0;
        ProgressText = "准备中";
        CanOpenOutput = false;
        CanOpenOutputFolder = false;

        _apiKeyStore.SaveApiKey(GeminiApiKey);

        var options = BuildOptions();
        var request = new ConvertJobRequest
        {
            PdfPath = PdfPath,
            OutputPath = OutputPath,
            PageRangeText = PageRangeText,
            Options = options,
            GeminiApiKey = GeminiApiKey
        };

        _cts = new CancellationTokenSource();
        var progress = new Progress<JobProgress>(UpdateProgress);

        try
        {
            var result = await _conversionService.ConvertAsync(request, progress, _cts.Token);
            if (!string.IsNullOrWhiteSpace(result.OutputPath))
            {
                OutputPath = result.OutputPath;
            }

            foreach (var failure in result.Failures)
            {
                var pageLabel = failure.PageNumber.HasValue ? $"P{failure.PageNumber}" : "P?";
                var attemptLabel = failure.Attempt > 0 ? $"A{failure.Attempt}" : "";
                Failures.Add($"{pageLabel} {attemptLabel} {failure.ErrorCode} {failure.Message}".Trim());
            }
            if (result.Status == JobStatus.Succeeded)
            {
                ProgressText = "转换完成";
            }
            else if (result.Status == JobStatus.Canceled)
            {
                ProgressText = "已取消";
            }
            else
            {
                ProgressText = "转换失败";
            }

            UpdateOutputAvailability();
        }
        catch (Exception ex)
        {
            Logs.Add("转换失败: " + ex.Message);
        }
        finally
        {
            IsBusy = false;
        }
    }

    [RelayCommand]
    private void Cancel()
    {
        _cts?.Cancel();
    }

    [RelayCommand]
    private void OpenOutput()
    {
        if (!File.Exists(OutputPath))
        {
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = OutputPath,
            UseShellExecute = true
        });
    }

    [RelayCommand]
    private void OpenOutputFolder()
    {
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            return;
        }

        var folder = Path.GetDirectoryName(OutputPath) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(folder))
        {
            return;
        }

        if (OperatingSystem.IsWindows())
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = $"/select,\"{OutputPath}\"",
                UseShellExecute = true
            });
        }
        else if (OperatingSystem.IsLinux())
        {
            Process.Start("xdg-open", folder);
        }
        else if (OperatingSystem.IsMacOS())
        {
            Process.Start("open", folder);
        }
    }

    [RelayCommand]
    private async Task RenderPreviewAsync()
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            PreviewStatus = "请先选择 PDF";
            return;
        }

        if (PdfPageCount <= 0)
        {
            await LoadPdfInfoAsync();
            if (PdfPageCount <= 0)
            {
                PreviewStatus = "无法读取 PDF 页数";
                return;
            }
        }

        var pageNumber = Math.Clamp(PreviewPageNumber, 1, PdfPageCount);
        PreviewPageNumber = pageNumber;

        _previewCts?.Cancel();
        _previewCts = new CancellationTokenSource();

        IsPreviewBusy = true;
        PreviewStatus = "预览渲染中...";

        try
        {
            var preview = await Task.Run(() => BuildPreview(pageNumber), _previewCts.Token);
            ApplyPreview(preview);
            PreviewStatus = "预览就绪";
        }
        catch (OperationCanceledException)
        {
            PreviewStatus = "预览已取消";
        }
        catch (Exception ex)
        {
            PreviewStatus = "预览失败: " + ex.Message;
        }
        finally
        {
            IsPreviewBusy = false;
        }
    }

    [RelayCommand]
    private async Task PreviewPrevAsync()
    {
        if (PdfPageCount <= 0)
        {
            return;
        }

        PreviewPageNumber = Math.Max(1, PreviewPageNumber - 1);
        await RenderPreviewAsync();
    }

    [RelayCommand]
    private async Task PreviewNextAsync()
    {
        if (PdfPageCount <= 0)
        {
            return;
        }

        PreviewPageNumber = Math.Min(PdfPageCount, PreviewPageNumber + 1);
        await RenderPreviewAsync();
    }

    private void UpdateProgress(JobProgress progress)
    {
        if (progress.TotalPages > 0)
        {
            ProgressPercent = progress.CompletedPages * 100.0 / progress.TotalPages;
        }
        ProgressText = $"{progress.Stage} - 第 {progress.CurrentPage ?? 0} 页";
    }

    private void OnLogPublished(LogEntry entry)
    {
        Application.Current.Dispatcher.Invoke(() =>
        {
            Logs.Add($"[{entry.TimestampUtc:HH:mm:ss}] {entry.Stage} {entry.Message}");
        });
    }

    private AppOptions BuildOptions()
    {
        return new AppOptions
        {
            Render = new CoreRenderOptions { Dpi = Dpi, ColorMode = _defaults.Render.ColorMode, MaxPreviewDpi = _defaults.Render.MaxPreviewDpi },
            Layout = new LayoutOptions
            {
                HeaderFooterMode = HeaderFooterMode,
                HeaderPercent = HeaderPercent,
                FooterPercent = FooterPercent,
                MaxCropTotalPercent = _defaults.Layout.MaxCropTotalPercent,
                PageSizeMode = PageSizeMode,
                MarginLeftTwips = _defaults.Layout.MarginLeftTwips,
                MarginRightTwips = _defaults.Layout.MarginRightTwips,
                MarginTopTwips = _defaults.Layout.MarginTopTwips,
                MarginBottomTwips = _defaults.Layout.MarginBottomTwips
            },
            Runtime = new RuntimeOptions { PageConcurrency = Concurrency, GeminiConcurrency = _defaults.Runtime.GeminiConcurrency },
            Preprocess = BuildPreprocessOptions(),
            TableDetect = new TableDetectOptions
            {
                Enable = _defaults.TableDetect.Enable,
                MinAreaRatio = _defaults.TableDetect.MinAreaRatio,
                MinWidthPx = _defaults.TableDetect.MinWidthPx,
                MinHeightPx = _defaults.TableDetect.MinHeightPx,
                MergeGapPx = _defaults.TableDetect.MergeGapPx,
                KernelDivisor = _defaults.TableDetect.KernelDivisor,
                MinKernelPx = _defaults.TableDetect.MinKernelPx,
                DilateKernelPx = _defaults.TableDetect.DilateKernelPx,
                ClusterEpsPx = _defaults.TableDetect.ClusterEpsPx,
                MinCellSizeRatio = _defaults.TableDetect.MinCellSizeRatio,
                OwnerConflictPolicy = _defaults.TableDetect.OwnerConflictPolicy,
                EnableTextTableFallback = _defaults.TableDetect.EnableTextTableFallback,
                EnableImageTableFallback = _defaults.TableDetect.EnableImageTableFallback
            },
            Gemini = new GeminiOptions
            {
                Endpoint = ResolveGeminiEndpoint(),
                ApiKeyStorage = _defaults.Gemini.ApiKeyStorage,
                TimeoutSeconds = new GeminiTimeouts { Table = _defaults.Gemini.TimeoutSeconds.Table, Page = _defaults.Gemini.TimeoutSeconds.Page },
                MaxRetryCount = _defaults.Gemini.MaxRetryCount,
                BackoffBaseMs = _defaults.Gemini.BackoffBaseMs,
                ExpectStrictJson = _defaults.Gemini.ExpectStrictJson,
                TableMissingIdThreshold = _defaults.Gemini.TableMissingIdThreshold,
                TableEmptyTextRateThreshold = _defaults.Gemini.TableEmptyTextRateThreshold,
                MinPageTextCharCount = _defaults.Gemini.MinPageTextCharCount,
                Image = new GeminiImageOptions { MaxLongEdgePx = _defaults.Gemini.Image.MaxLongEdgePx, JpegQuality = _defaults.Gemini.Image.JpegQuality }
            },
            Docx = new DocxWriteOptions
            {
                FontEastAsia = _defaults.Docx.FontEastAsia,
                FontAscii = _defaults.Docx.FontAscii,
                DefaultFontSizeHalfPoints = _defaults.Docx.DefaultFontSizeHalfPoints,
                UseStylesPart = _defaults.Docx.UseStylesPart,
                Table = new DocxTableOptions { WidthMode = _defaults.Docx.Table.WidthMode, SetBorders = _defaults.Docx.Table.SetBorders },
                Paragraph = new DocxParagraphOptions { KeepLineBreaks = _defaults.Docx.Paragraph.KeepLineBreaks },
                PageBreak = new DocxPageBreakOptions { ModeA4 = _defaults.Docx.PageBreak.ModeA4, ModeFollowPdf = _defaults.Docx.PageBreak.ModeFollowPdf }
            },
            Diagnostics = new DiagnosticsOptions { KeepTempFiles = KeepDiagnostics, ExportZip = _defaults.Diagnostics.ExportZip, SaveRawGeminiJson = _defaults.Diagnostics.SaveRawGeminiJson },
            Validation = new ValidationOptions
            {
                Enable = _defaults.Validation.Enable,
                FailFast = _defaults.Validation.FailFast,
                AllowSkipFailedPages = _defaults.Validation.AllowSkipFailedPages,
                TableStructureOkTextBad = _defaults.Validation.TableStructureOkTextBad,
                TableStructureBad = _defaults.Validation.TableStructureBad,
                TextEmpty = _defaults.Validation.TextEmpty
            },
            Output = new OutputOptions
            {
                OutputDirectory = _defaults.Output.OutputDirectory,
                FileNameMode = _defaults.Output.FileNameMode,
                Overwrite = _defaults.Output.Overwrite
            }
        };
    }

    private PreprocessOptions BuildPreprocessOptions()
    {
        return new PreprocessOptions
        {
            EnableDeskew = DeskewEnabled,
            ContrastEnhance = _defaults.Preprocess.ContrastEnhance,
            Clahe = new ClaheOptions { ClipLimit = _defaults.Preprocess.Clahe.ClipLimit, TileGridSize = _defaults.Preprocess.Clahe.TileGridSize },
            Denoise = _defaults.Preprocess.Denoise,
            Binarize = _defaults.Preprocess.Binarize,
            Adaptive = new AdaptiveThresholdOptions { BlockSize = _defaults.Preprocess.Adaptive.BlockSize, C = _defaults.Preprocess.Adaptive.C },
            Deskew = new DeskewOptions { MinAngleDeg = _defaults.Preprocess.Deskew.MinAngleDeg, MaxAngleDeg = _defaults.Preprocess.Deskew.MaxAngleDeg }
        };
    }

    private void UpdateHeaderFooterEnabled()
    {
        IsHeaderEnabled = HeaderFooterMode is HeaderFooterRemoveMode.RemoveHeader or HeaderFooterRemoveMode.RemoveBoth;
        IsFooterEnabled = HeaderFooterMode is HeaderFooterRemoveMode.RemoveFooter or HeaderFooterRemoveMode.RemoveBoth;
        ValidateCrop();
    }

    private void ValidateCrop()
    {
        var effectiveHeader = IsHeaderEnabled ? HeaderPercent : 0;
        var effectiveFooter = IsFooterEnabled ? FooterPercent : 0;

        if (effectiveHeader < 0 || effectiveFooter < 0)
        {
            CropError = "裁切比例不能为负数。";
            HasCropError = true;
            return;
        }

        if (effectiveHeader > MaxCropPercentPerEdge || effectiveFooter > MaxCropPercentPerEdge)
        {
            CropError = "裁切比例建议不超过 0.20。";
            HasCropError = true;
            return;
        }

        if (effectiveHeader + effectiveFooter > _defaults.Layout.MaxCropTotalPercent)
        {
            CropError = "页眉/页脚裁切总比例过大。";
            HasCropError = true;
            return;
        }

        CropError = string.Empty;
        HasCropError = false;
    }

    private void ValidateGeminiEndpoint()
    {
        var value = GeminiEndpointUrl?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value))
        {
            GeminiEndpointError = "请填写 Gemini URL。";
            HasGeminiEndpointError = true;
            return;
        }

        if (!value.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            GeminiEndpointError = "Gemini URL 必须以 https:// 开头。";
            HasGeminiEndpointError = true;
            return;
        }

        if (!value.EndsWith("generateContent", StringComparison.OrdinalIgnoreCase))
        {
            GeminiEndpointError = "Gemini URL 必须以 generateContent 结束。";
            HasGeminiEndpointError = true;
            return;
        }

        GeminiEndpointError = string.Empty;
        HasGeminiEndpointError = false;
    }

    private string ResolveGeminiEndpoint()
    {
        var value = GeminiEndpointUrl?.Trim() ?? string.Empty;
        return string.IsNullOrWhiteSpace(value) ? _defaults.Gemini.Endpoint : value;
    }

    private async Task LoadPdfInfoAsync()
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            PdfPageCount = 0;
            PdfInfoText = string.Empty;
            return;
        }

        try
        {
            var count = await Task.Run(() => _pdfRenderer.GetPageCount(PdfPath));
            PdfPageCount = count;
            PdfInfoText = count > 0 ? $"共 {count} 页" : "未读取到页数";
            if (PreviewPageNumber <= 0)
            {
                PreviewPageNumber = 1;
            }
            if (count > 0 && PreviewPageNumber > count)
            {
                PreviewPageNumber = count;
            }
            UpdatePageRangeSummary();
        }
        catch (Exception ex)
        {
            PdfPageCount = 0;
            PdfInfoText = "读取页数失败";
            Logs.Add("读取页数失败: " + ex.Message);
        }
    }

    private void UpdatePageRangeSummary()
    {
        if (PdfPageCount <= 0)
        {
            PageRangeSummary = string.Empty;
            PageRangeError = string.Empty;
            HasPageRangeError = false;
            return;
        }

        var result = PageRangeParser.Parse(PageRangeText, PdfPageCount);
        if (result.HasError)
        {
            PageRangeSummary = string.Empty;
            PageRangeError = result.ErrorMessage ?? "页码范围格式不正确";
            HasPageRangeError = true;
            return;
        }

        if (result.Pages.Count == 0)
        {
            PageRangeSummary = "页码范围为空";
            PageRangeError = "页码范围为空";
            HasPageRangeError = true;
            return;
        }

        HasPageRangeError = false;
        PageRangeError = string.Empty;
        PageRangeSummary = $"将转换 {result.Pages.Count} 页";
    }

    private void UpdateOutputAvailability()
    {
        CanOpenOutput = !string.IsNullOrWhiteSpace(OutputPath) && File.Exists(OutputPath);
        var folder = string.IsNullOrWhiteSpace(OutputPath) ? string.Empty : Path.GetDirectoryName(OutputPath) ?? string.Empty;
        CanOpenOutputFolder = !string.IsNullOrWhiteSpace(folder) && Directory.Exists(folder);
    }

    private PreviewPayload BuildPreview(int pageNumber)
    {
        var previewDpi = Math.Min(Dpi, _defaults.Render.MaxPreviewDpi);
        using var rendered = _pdfRenderer.RenderPage(PdfPath, pageNumber - 1, previewDpi);
        var bundle = _imagePipeline.Process(rendered, new CropOptions
        {
            Mode = HeaderFooterMode,
            HeaderPercent = HeaderPercent,
            FooterPercent = FooterPercent
        }, BuildPreprocessOptions(), pageNumber);

        var original = BitmapSourceHelper.ToBitmapSource(bundle.OriginalColor);
        var cropped = BitmapSourceHelper.ToBitmapSource(bundle.CroppedColor);

        var boxes = new List<BBox>();
        if (_defaults.TableDetect.Enable)
        {
            var tables = _tableEngine.DetectTables(bundle, _defaults.TableDetect);
            boxes = tables.Select(t => t.TableBBoxInPage).ToList();
        }

        var payload = new PreviewPayload(original, cropped, bundle.OriginalSizePx, bundle.CroppedSizePx, bundle.CropInfo, boxes);
        DisposeBundle(bundle);
        return payload;
    }

    private void ApplyPreview(PreviewPayload payload)
    {
        _previewOriginalImage = payload.OriginalImage;
        _previewCroppedImage = payload.CroppedImage;
        _previewOriginalSize = payload.OriginalSize;
        _previewCroppedSize = payload.CroppedSize;
        _previewCropInfo = payload.CropInfo;
        _previewTableBoxes = payload.TableBoxes.ToList();
        UpdatePreviewImage();
    }

    private void UpdatePreviewImage()
    {
        PreviewImage = PreviewShowCropped ? _previewCroppedImage : _previewOriginalImage;
        if (PreviewShowCropped)
        {
            PreviewImageWidth = Math.Max(1, _previewCroppedSize.W);
            PreviewImageHeight = Math.Max(1, _previewCroppedSize.H);
        }
        else
        {
            PreviewImageWidth = Math.Max(1, _previewOriginalSize.W);
            PreviewImageHeight = Math.Max(1, _previewOriginalSize.H);
        }

        UpdatePreviewOverlays();
    }

    private void UpdatePreviewOverlays()
    {
        PreviewOverlays.Clear();
        if (PreviewShowCropped)
        {
            if (PreviewShowTables)
            {
                foreach (var box in _previewTableBoxes)
                {
                    PreviewOverlays.Add(new PreviewOverlay
                    {
                        X = box.X,
                        Y = box.Y,
                        Width = box.W,
                        Height = box.H,
                        Stroke = TableStrokeBrush,
                        StrokeThickness = 2,
                        Fill = Brushes.Transparent
                    });
                }
            }
            return;
        }

        if (PreviewShowCropOverlay && _previewCropInfo is { } cropInfo)
        {
            if (cropInfo.CropTopPx > 0)
            {
                PreviewOverlays.Add(new PreviewOverlay
                {
                    X = 0,
                    Y = 0,
                    Width = PreviewImageWidth,
                    Height = cropInfo.CropTopPx,
                    Stroke = Brushes.Transparent,
                    Fill = CropOverlayBrush
                });
            }

            if (cropInfo.CropBottomPx > 0)
            {
                PreviewOverlays.Add(new PreviewOverlay
                {
                    X = 0,
                    Y = Math.Max(0, PreviewImageHeight - cropInfo.CropBottomPx),
                    Width = PreviewImageWidth,
                    Height = cropInfo.CropBottomPx,
                    Stroke = Brushes.Transparent,
                    Fill = CropOverlayBrush
                });
            }
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

    private readonly record struct PreviewPayload(
        ImageSource OriginalImage,
        ImageSource CroppedImage,
        (int W, int H) OriginalSize,
        (int W, int H) CroppedSize,
        CropInfo CropInfo,
        IReadOnlyList<BBox> TableBoxes);
}
