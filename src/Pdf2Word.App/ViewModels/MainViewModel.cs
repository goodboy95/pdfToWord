using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using Pdf2Word.App.Services;
using Pdf2Word.Core.Logging;
using Pdf2Word.Core.Models;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;

namespace Pdf2Word.App.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IConversionService _conversionService;
    private readonly AppOptions _defaults;
    private readonly IApiKeyStore _apiKeyStore;

    private CancellationTokenSource? _cts;

    public ObservableCollection<int> DpiOptions { get; } = new() { 200, 300, 400 };
    public ObservableCollection<int> ConcurrencyOptions { get; } = new() { 1, 2, 4 };
    public ObservableCollection<HeaderFooterRemoveMode> HeaderFooterModes { get; } = new(Enum.GetValues<HeaderFooterRemoveMode>());
    public ObservableCollection<PageSizeMode> PageSizeModes { get; } = new(Enum.GetValues<PageSizeMode>());

    public ObservableCollection<string> Logs { get; } = new();
    public ObservableCollection<string> Failures { get; } = new();

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
    private string _geminiApiKey = string.Empty;

    [ObservableProperty]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressPercent;

    [ObservableProperty]
    private string _progressText = string.Empty;

    public bool CanStart => !IsBusy && !string.IsNullOrWhiteSpace(PdfPath);

    public MainViewModel(IConversionService conversionService, AppOptions defaults, IApiKeyStore apiKeyStore, UiLogSink logSink)
    {
        _conversionService = conversionService;
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
        GeminiApiKey = apiKeyStore.GetApiKey() ?? string.Empty;

        logSink.LogPublished += OnLogPublished;
    }

    partial void OnPdfPathChanged(string value) => OnPropertyChanged(nameof(CanStart));
    partial void OnIsBusyChanged(bool value) => OnPropertyChanged(nameof(CanStart));

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
            foreach (var failure in result.Failures)
            {
                Failures.Add($"P{failure.PageNumber ?? 0} {failure.ErrorCode} {failure.Message}");
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
            Render = new RenderOptions { Dpi = Dpi, ColorMode = _defaults.Render.ColorMode, MaxPreviewDpi = _defaults.Render.MaxPreviewDpi },
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
            Preprocess = new PreprocessOptions
            {
                EnableDeskew = DeskewEnabled,
                ContrastEnhance = _defaults.Preprocess.ContrastEnhance,
                Clahe = new ClaheOptions { ClipLimit = _defaults.Preprocess.Clahe.ClipLimit, TileGridSize = _defaults.Preprocess.Clahe.TileGridSize },
                Denoise = _defaults.Preprocess.Denoise,
                Binarize = _defaults.Preprocess.Binarize,
                Adaptive = new AdaptiveThresholdOptions { BlockSize = _defaults.Preprocess.Adaptive.BlockSize, C = _defaults.Preprocess.Adaptive.C },
                Deskew = new DeskewOptions { MinAngleDeg = _defaults.Preprocess.Deskew.MinAngleDeg, MaxAngleDeg = _defaults.Preprocess.Deskew.MaxAngleDeg }
            },
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
                Endpoint = _defaults.Gemini.Endpoint,
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
}
