using System.Net.Http;
using System.Windows;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Pdf2Word.Core.Logging;
using Pdf2Word.Core.Options;
using Pdf2Word.Core.Services;
using Pdf2Word.Infrastructure.Docx;
using Pdf2Word.Infrastructure.Gemini;
using Pdf2Word.Infrastructure.ImageProcessing;
using Pdf2Word.Infrastructure.Logging;
using Pdf2Word.Infrastructure.Pdf;
using Pdf2Word.Infrastructure.Storage;
using Pdf2Word.Infrastructure.Table;
using Pdf2Word.App.Services;
using Pdf2Word.App.ViewModels;

namespace Pdf2Word.App;

public partial class App : Application
{
    private IHost? _host;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        _host = Host.CreateDefaultBuilder()
            .ConfigureAppConfiguration(builder =>
            {
                builder.SetBasePath(AppContext.BaseDirectory);
                builder.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            })
            .ConfigureServices((context, services) =>
            {
                var options = context.Configuration.Get<AppOptions>() ?? new AppOptions();
                services.AddSingleton(options);

                services.AddSingleton<IPdfRenderer, PdfiumRenderer>();
                services.AddSingleton<IPageImagePipeline, OpenCvPageImagePipeline>();
                services.AddSingleton<ITableEngine, OpenCvTableEngine>();
                services.AddSingleton<IDocxWriter, OpenXmlDocxWriter>();
                services.AddSingleton<ITempStorage, TempStorage>();
                services.AddSingleton<IApiKeyStore, DpapiApiKeyStore>();
                services.AddSingleton<UiLogSink>();
                services.AddSingleton<FileLogSink>();
                services.AddSingleton<InstallLogSink>();
                services.AddSingleton<ILogSink>(sp => new CompositeLogSink(
                    sp.GetRequiredService<UiLogSink>(),
                    sp.GetRequiredService<FileLogSink>(),
                    sp.GetRequiredService<InstallLogSink>()));
                services.AddHttpClient();
                services.AddSingleton<IGeminiClient>(sp =>
                {
                    var http = sp.GetRequiredService<IHttpClientFactory>().CreateClient();
                    var opt = sp.GetRequiredService<AppOptions>().Gemini;
                    var keyStore = sp.GetRequiredService<IApiKeyStore>();
                    return new GeminiClient(http, opt, keyStore.GetApiKey);
                });
                services.AddSingleton<IConversionService, ConversionService>();
                services.AddSingleton<MainViewModel>();
                services.AddSingleton<MainWindow>();
            })
            .Build();

        _host.Start();
        var mainWindow = _host.Services.GetRequiredService<MainWindow>();
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _host?.Dispose();
        base.OnExit(e);
    }
}
