using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services
{
    public class OcrInitializationService
    {
        private readonly ILogger<OcrInitializationService> _logger;
        private readonly HttpClient _httpClient;
        
        private const string TESSDATA_DIR = "tessdata";
        private const string CHI_SIM_FILE = "chi_sim.traineddata";
        private const string CHI_SIM_URL = "https://github.com/tesseract-ocr/tessdata/raw/main/chi_sim.traineddata";
        private const string CHI_SIM_BACKUP_URL = "https://raw.githubusercontent.com/tesseract-ocr/tessdata/main/chi_sim.traineddata";
        private const string DOWNLOAD_PAGE_URL = "https://github.com/tesseract-ocr/tessdata";

        public OcrInitializationService(ILogger<OcrInitializationService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _httpClient = new HttpClient
            {
                Timeout = TimeSpan.FromMinutes(5)
            };
        }

        public async Task<bool> InitializeOcrAsync()
        {
            try
            {
                _logger.LogInformation("开始初始化OCR环境");

                var tessdataPath = GetTessdataPath();
                
                if (!Directory.Exists(tessdataPath))
                {
                    Directory.CreateDirectory(tessdataPath);
                    _logger.LogInformation($"创建tessdata目录: {tessdataPath}");
                }

                var chiSimPath = Path.Combine(tessdataPath, CHI_SIM_FILE);

                if (File.Exists(chiSimPath))
                {
                    var fileInfo = new FileInfo(chiSimPath);
                    if (fileInfo.Length > 1000)
                    {
                        _logger.LogInformation($"中文OCR数据文件已存在: {chiSimPath} ({fileInfo.Length} bytes)");
                        return true;
                    }
                    else
                    {
                        _logger.LogWarning($"中文OCR数据文件存在但大小异常: {fileInfo.Length} bytes，将重新下载");
                        File.Delete(chiSimPath);
                    }
                }

                _logger.LogInformation("中文OCR数据文件不存在，开始下载");

                var downloadSuccess = await DownloadChiSimDataAsync(chiSimPath);

                if (downloadSuccess)
                {
                    _logger.LogInformation("中文OCR数据文件下载成功");
                    return true;
                }
                else
                {
                    _logger.LogWarning("自动下载失败，提示用户手动下载");
                    return await PromptUserManualDownloadAsync(tessdataPath);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "OCR初始化失败");
                return await PromptUserManualDownloadAsync(GetTessdataPath());
            }
        }

        private string GetTessdataPath()
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var tessdataPath = Path.Combine(baseDir, TESSDATA_DIR);
            
            if (!Directory.Exists(tessdataPath))
            {
                var assemblyLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
                var assemblyDir = Path.GetDirectoryName(assemblyLocation);
                tessdataPath = Path.Combine(assemblyDir, TESSDATA_DIR);
            }
            
            return tessdataPath;
        }

        private async Task<bool> DownloadChiSimDataAsync(string targetPath)
        {
            string[] urls = { CHI_SIM_URL, CHI_SIM_BACKUP_URL };

            foreach (var url in urls)
            {
                try
                {
                    _logger.LogInformation($"尝试从 {url} 下载");

                    var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
                    response.EnsureSuccessStatusCode();

                    var totalBytes = response.Content.Headers.ContentLength ?? 0;
                    _logger.LogInformation($"文件大小: {totalBytes} bytes");

                    var tempPath = targetPath + ".tmp";

                    using (var contentStream = await response.Content.ReadAsStreamAsync())
                    using (var fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        var buffer = new byte[8192];
                        var bytesRead = 0;
                        var totalRead = 0L;

                        while ((bytesRead = await contentStream.ReadAsync(buffer, 0, buffer.Length)) > 0)
                        {
                            await fileStream.WriteAsync(buffer, 0, bytesRead);
                            totalRead += bytesRead;

                            if (totalBytes > 0)
                            {
                                var progress = (double)totalRead / totalBytes * 100;
                                if (totalRead % (1024 * 1024) == 0)
                                {
                                    _logger.LogInformation($"下载进度: {progress:F1}% ({totalRead}/{totalBytes} bytes)");
                                }
                            }
                        }
                    }

                    if (File.Exists(targetPath))
                    {
                        File.Delete(targetPath);
                    }

                    File.Move(tempPath, targetPath);

                    var fileInfo = new FileInfo(targetPath);
                    if (fileInfo.Length > 1000)
                    {
                        _logger.LogInformation($"下载完成: {fileInfo.Length} bytes");
                        return true;
                    }
                    else
                    {
                        _logger.LogWarning($"下载的文件大小异常: {fileInfo.Length} bytes");
                        File.Delete(targetPath);
                        return false;
                    }
                }
                catch (HttpRequestException ex)
                {
                    _logger.LogWarning(ex, $"从 {url} 下载失败");
                    continue;
                }
                catch (TaskCanceledException ex)
                {
                    _logger.LogWarning(ex, $"从 {url} 下载超时");
                    continue;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"从 {url} 下载时发生错误");
                    continue;
                }
            }

            return false;
        }

        private async Task<bool> PromptUserManualDownloadAsync(string tessdataPath)
        {
            return await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var message = $"OCR功能需要中文语言数据文件，但自动下载失败。\n\n" +
                             $"请按以下步骤手动下载：\n" +
                             $"1. 点击下方按钮打开下载页面\n" +
                             $"2. 下载文件: {CHI_SIM_FILE}\n" +
                             $"3. 将文件保存到以下目录:\n" +
                             $"   {tessdataPath}\n\n" +
                             $"注意: OCR功能将无法使用，直到文件下载完成。";

                var result = MessageBox.Show(
                    message,
                    "OCR初始化失败",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning,
                    MessageBoxResult.Yes,
                    MessageBoxOptions.DefaultDesktopOnly);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = DOWNLOAD_PAGE_URL,
                            UseShellExecute = true
                        });

                        MessageBox.Show(
                            $"请将下载的 {CHI_SIM_FILE} 文件保存到:\n\n{tessdataPath}\n\n" +
                            $"保存后，请重新启动应用程序以启用OCR功能。",
                            "下载说明",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information,
                            MessageBoxResult.OK,
                            MessageBoxOptions.DefaultDesktopOnly);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "无法打开下载页面");
                        MessageBox.Show(
                            $"无法自动打开下载页面，请手动访问:\n{DOWNLOAD_PAGE_URL}\n\n" +
                            $"请将 {CHI_SIM_FILE} 保存到:\n{tessdataPath}",
                            "无法打开浏览器",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error,
                            MessageBoxResult.OK,
                            MessageBoxOptions.DefaultDesktopOnly);
                    }
                }

                return false;
            });
        }

        public bool IsOcrReady()
        {
            var chiSimPath = Path.Combine(GetTessdataPath(), CHI_SIM_FILE);
            return File.Exists(chiSimPath) && new FileInfo(chiSimPath).Length > 1000;
        }
    }
}
