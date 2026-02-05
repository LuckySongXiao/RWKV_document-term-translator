using System;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using DocumentTranslator.Services;

namespace DocumentTranslator
{
    public partial class App : Application
    {
        private ServiceProvider _serviceProvider;
        private ILogger<App> _logger;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            ConfigureServices();
            
            _logger = _serviceProvider.GetRequiredService<ILogger<App>>();
            _logger.LogInformation("应用程序启动");

            InitializeServicesAsync();

            this.DispatcherUnhandledException += (sender, args) =>
            {
                _logger?.LogError(args.Exception, "应用程序发生未处理的异常");
                MessageBox.Show($"应用程序发生未处理的异常：\n{args.Exception.Message}", 
                              "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true;
            };
        }

        private void ConfigureServices()
        {
            var services = new ServiceCollection();

            services.AddLogging(configure =>
            {
                configure.AddConsole();
                configure.AddDebug();
                configure.SetMinimumLevel(LogLevel.Information);
            });

            services.AddSingleton<OcrInitializationService>();
            services.AddSingleton<AcrobatDetectionService>();

            _serviceProvider = services.BuildServiceProvider();
        }

        private async void InitializeServicesAsync()
        {
            try
            {
                var ocrService = _serviceProvider.GetRequiredService<OcrInitializationService>();
                var acrobatService = _serviceProvider.GetRequiredService<AcrobatDetectionService>();

                _logger.LogInformation("开始初始化服务");

                var ocrReady = await ocrService.InitializeOcrAsync();
                if (ocrReady)
                {
                    _logger.LogInformation("OCR初始化成功");
                }
                else
                {
                    _logger.LogWarning("OCR初始化失败，OCR功能将不可用");
                }

                acrobatService.ShowAcrobatAvailableInfo();

                _logger.LogInformation("服务初始化完成");

                var mainWindow = new MainWindow();
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "服务初始化失败");
                MessageBox.Show($"服务初始化失败：\n{ex.Message}", 
                              "初始化错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            _logger?.LogInformation("应用程序退出");
            _serviceProvider?.Dispose();
            base.OnExit(e);
        }
    }
}
