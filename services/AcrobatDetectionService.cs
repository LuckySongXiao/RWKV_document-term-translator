using System;
using System.IO;
using System.Windows;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services
{
    public class AcrobatDetectionService
    {
        private readonly ILogger<AcrobatDetectionService> _logger;

        private static readonly string[] ADOBE_ACROBAT_PATHS = new[]
        {
            @"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            @"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            @"C:\Program Files\Adobe\Acrobat 2020\Acrobat\Acrobat.exe",
            @"C:\Program Files (x86)\Adobe\Acrobat 2020\Acrobat\Acrobat.exe",
            @"C:\Program Files\Adobe\Acrobat 2017\Acrobat\Acrobat.exe",
            @"C:\Program Files (x86)\Adobe\Acrobat 2017\Acrobat\Acrobat.exe",
            @"C:\Program Files\Adobe\Acrobat 2015\Acrobat\Acrobat.exe",
            @"C:\Program Files (x86)\Adobe\Acrobat 2015\Acrobat\Acrobat.exe"
        };

        private const string ADOBE_DOWNLOAD_URL = "https://get.adobe.com/reader/";

        public AcrobatDetectionService(ILogger<AcrobatDetectionService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string FindAdobeAcrobat()
        {
            foreach (var path in ADOBE_ACROBAT_PATHS)
            {
                if (File.Exists(path))
                {
                    _logger.LogInformation($"找到Adobe Acrobat: {path}");
                    return path;
                }
            }

            _logger.LogWarning("未找到Adobe Acrobat");
            return null;
        }

        public bool IsAdobeAcrobatAvailable()
        {
            return FindAdobeAcrobat() != null;
        }

        public void ShowAcrobatNotAvailableWarning()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                var message = "未检测到Adobe Acrobat。\n\n" +
                             "Adobe Acrobat预处理功能将不可用。\n\n" +
                             "如需使用此功能，请安装Adobe Acrobat（非Reader版本）。\n\n" +
                             "注意：OCR功能仍然可用，但识别准确率可能较低。";

                var result = MessageBox.Show(
                    message,
                    "Adobe Acrobat未安装",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Warning,
                    MessageBoxResult.OK,
                    MessageBoxOptions.DefaultDesktopOnly);

                if (result == MessageBoxResult.OK)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = ADOBE_DOWNLOAD_URL,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "无法打开Adobe下载页面");
                        MessageBox.Show(
                            $"无法自动打开下载页面，请手动访问:\n{ADOBE_DOWNLOAD_URL}",
                            "无法打开浏览器",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error,
                            MessageBoxResult.OK,
                            MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
            });
        }

        public void ShowAcrobatAvailableInfo()
        {
            var acrobatPath = FindAdobeAcrobat();
            if (acrobatPath != null)
            {
                _logger.LogInformation($"Adobe Acrobat已安装: {acrobatPath}");
            }
        }
    }
}
