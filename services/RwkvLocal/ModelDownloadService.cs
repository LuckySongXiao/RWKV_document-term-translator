using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// ModelScope 模型下载服务
    /// </summary>
    public class ModelDownloadService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly HttpClient _httpClient;
        
        // ModelScope 配置（可通过SetModelScopeConfig动态修改）
        private const string ModelScopeDownloadBase = "https://www.modelscope.cn/models";
        private string _modelOwner = "AlicLi";
        private string _modelName = "rwkv7-g1-libtorch-st";
        private bool _isGgufRepo = false;  // 标记当前仓库是否为GGUF模型仓库
        
        // 下载进度事件
        public event EventHandler<DownloadProgressEventArgs>? ProgressChanged;
        public event EventHandler<DownloadCompletedEventArgs>? DownloadCompleted;
        public event EventHandler<string>? StatusChanged;

        public ModelDownloadService(ILogger logger)
        {
            _logger = logger;
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromHours(2);
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36");
            _httpClient.DefaultRequestHeaders.Add("Accept", "*/*");
            UpdateRefererHeader();
        }

        /// <summary>
        /// 从ModelScope URL解析并设置Owner和ModelName
        /// 支持格式: https://www.modelscope.cn/models/{Owner}/{ModelName}/files 或
        ///           https://www.modelscope.cn/models/{Owner}/{ModelName}
        /// </summary>
        public bool SetModelScopeConfigFromUrl(string url)
        {
            try
            {
                // 去掉末尾的 /files 或 /files/
                var cleanUrl = url.TrimEnd('/');
                if (cleanUrl.EndsWith("/files", StringComparison.OrdinalIgnoreCase))
                {
                    cleanUrl = cleanUrl.Substring(0, cleanUrl.Length - "/files".Length);
                }

                // 解析路径部分: /models/{Owner}/{ModelName}
                var uri = new Uri(cleanUrl);
                var segments = uri.AbsolutePath.Trim('/').Split('/');
                
                // 期望格式: models/{Owner}/{ModelName}
                if (segments.Length >= 3 && segments[0].Equals("models", StringComparison.OrdinalIgnoreCase))
                {
                    _modelOwner = segments[1];
                    _modelName = segments[2];
                    
                    // 自动检测是否为GGUF仓库（仓库名包含GGUF或gguf）
                    _isGgufRepo = _modelName.IndexOf("GGUF", StringComparison.OrdinalIgnoreCase) >= 0;
                    
                    UpdateRefererHeader();
                    _logger.LogInformation("ModelScope配置已更新: Owner={Owner}, ModelName={ModelName}, IsGgufRepo={IsGguf}", _modelOwner, _modelName, _isGgufRepo);
                    return true;
                }
                
                _logger.LogWarning("无法从URL解析ModelScope配置: {Url}", url);
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "解析ModelScope URL失败: {Url}", url);
                return false;
            }
        }

        /// <summary>
        /// 获取当前ModelScope文件浏览页面URL
        /// </summary>
        public string GetModelScopeFilesUrl()
        {
            return $"{ModelScopeDownloadBase}/{_modelOwner}/{_modelName}/files";
        }

        private void UpdateRefererHeader()
        {
            // 移除旧的Referer头
            _httpClient.DefaultRequestHeaders.Remove("Referer");
            _httpClient.DefaultRequestHeaders.Add("Referer", $"{ModelScopeDownloadBase}/{_modelOwner}/{_modelName}/files");
        }

        /// <summary>
        /// 构建 ModelScope 下载 URL
        /// 格式: https://www.modelscope.cn/models/{user}/{model}/resolve/master/{filename}
        /// </summary>
        private string BuildDownloadUrl(string fileName)
        {
            // ModelScope 直接下载链接格式
            return $"{ModelScopeDownloadBase}/{_modelOwner}/{_modelName}/resolve/master/{Uri.EscapeDataString(fileName)}";
        }

        /// <summary>
        /// 获取默认模型文件列表
        /// 仅对已知的safetensors仓库提供硬编码默认列表
        /// GGUF仓库不提供默认列表，强制通过API动态获取
        /// </summary>
        public List<ModelScopeFile> GetDefaultModelFiles()
        {
            // GGUF仓库不提供硬编码默认文件列表，需要通过API动态获取
            if (_isGgufRepo)
            {
                _logger.LogInformation("当前为GGUF仓库，不使用硬编码默认文件列表，将通过API动态获取");
                return new List<ModelScopeFile>();
            }
            
            var files = new List<ModelScopeFile>
            {
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1d-0.1b-20260129-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1D-0.1B (382MB) - 轻量版",
                    Size = "382 MB",
                    SizeBytes = 382303080,
                    Description = "轻量级翻译模型，适合显存较小的GPU",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1d-0.1b-20260129-ctx8192.safetensors")
                },
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1d-0.4b-20260210-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1D-0.4B (901MB) - 入门版",
                    Size = "901 MB",
                    SizeBytes = 901872600,
                    Description = "入门级翻译模型，性价比高",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1d-0.4b-20260210-ctx8192.safetensors")
                },
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1e-1.5b-20260309-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1E-1.5B (3GB) - 标准版",
                    Size = "3 GB",
                    SizeBytes = 3055673472,
                    Description = "标准翻译模型，推荐大多数用户使用",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1e-1.5b-20260309-ctx8192.safetensors")
                },
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1e-2.9b-20260312-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1E-2.9B (5.8GB) - 增强版",
                    Size = "5.8 GB",
                    SizeBytes = 5896557136,
                    Description = "增强翻译模型，效果更好",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1e-2.9b-20260312-ctx8192.safetensors")
                },
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1e-7.2b-20260301-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1E-7.2B (14.4GB) - 专业版",
                    Size = "14.4 GB",
                    SizeBytes = 14400488936,
                    Description = "专业级翻译模型，需要较大显存",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1e-7.2b-20260301-ctx8192.safetensors")
                },
                new ModelScopeFile
                {
                    FileName = "rwkv7-g1e-13.3b-20260309-ctx8192.safetensors",
                    DisplayName = "📦 RWKV7-G1E-13.3B (26.5GB) - 旗舰版",
                    Size = "26.5 GB",
                    SizeBytes = 26541837560,
                    Description = "旗舰翻译模型，效果最佳，需要大容量显存",
                    DownloadUrl = BuildDownloadUrl("rwkv7-g1e-13.3b-20260309-ctx8192.safetensors")
                }
            };
            
            return files;
        }

        /// <summary>
        /// 获取 ModelScope 模型文件列表
        /// </summary>
        public async Task<List<ModelScopeFile>> GetModelFilesAsync(CancellationToken cancellationToken = default)
        {
            var files = new List<ModelScopeFile>();
            
            try
            {
                _logger.LogInformation("正在获取 ModelScope 文件列表");
                StatusChanged?.Invoke(this, "正在获取文件列表...");
                
                // 使用 ModelScope API 获取文件列表
                // 正确的API端点是 /repo/files（而非 /repo?Revision=master，后者会返回400）
                var apiUrls = new[]
                {
                    $"https://www.modelscope.cn/api/v1/models/{_modelOwner}/{_modelName}/repo/files",
                    $"https://www.modelscope.cn/api/v1/models/{_modelOwner}/{_modelName}/repo?Revision=master"
                };
                
                foreach (var apiUrl in apiUrls)
                {
                    _logger.LogInformation("尝试 API URL: {Url}", apiUrl);
                    
                    try
                    {
                        var response = await _httpClient.GetAsync(apiUrl, cancellationToken);
                        
                        if (response.IsSuccessStatusCode)
                        {
                            var json = await response.Content.ReadAsStringAsync(cancellationToken);
                            files = ParseModelScopeApiResponse(json);
                            if (files.Count > 0)
                            {
                                break;
                            }
                        }
                        else
                        {
                            _logger.LogWarning("API 返回错误状态码: {StatusCode} for {Url}", response.StatusCode, apiUrl);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "请求 API 失败: {Url}", apiUrl);
                    }
                }
                
                if (files.Count == 0)
                {
                    files = GetDefaultModelFiles();
                    if (files.Count == 0)
                    {
                        _logger.LogWarning("无法从API获取文件列表且无默认列表，请检查ModelScope仓库地址是否正确");
                        StatusChanged?.Invoke(this, "无法获取文件列表，请检查仓库地址");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "获取 ModelScope 文件列表失败");
                files = GetDefaultModelFiles();
                if (files.Count == 0)
                {
                    _logger.LogWarning("获取文件列表失败且无默认列表，请检查ModelScope仓库地址和网络连接");
                    StatusChanged?.Invoke(this, "获取文件列表失败，请检查网络和仓库地址");
                }
            }
            
            _logger.LogInformation("获取到 {Count} 个文件", files.Count);
            StatusChanged?.Invoke(this, $"获取到 {files.Count} 个模型文件");
            
            return files;
        }

        /// <summary>
        /// 解析 ModelScope API JSON 响应
        /// </summary>
        private List<ModelScopeFile> ParseModelScopeApiResponse(string json)
        {
            var files = new List<ModelScopeFile>();
            
            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;
                
                // ModelScope /repo/files API 返回格式: {"Code":200,"Data":{"Files":[...]}}
                // 尝试多种路径定位文件数组
                JsonElement fileArray = default;
                bool foundArray = false;
                
                // 路径1: Data.Files（/repo/files 端点格式）
                if (!foundArray && root.TryGetProperty("Data", out var dataElem))
                {
                    if (dataElem.TryGetProperty("Files", out fileArray))
                    {
                        foundArray = true;
                    }
                    // 路径2: Data 本身是数组
                    else if (dataElem.ValueKind == JsonValueKind.Array)
                    {
                        fileArray = dataElem;
                        foundArray = true;
                    }
                }
                
                // 路径3: data 或 files（旧格式兼容）
                if (!foundArray)
                {
                    if (root.TryGetProperty("data", out fileArray) || root.TryGetProperty("files", out fileArray))
                    {
                        foundArray = true;
                    }
                }
                
                if (foundArray && fileArray.ValueKind == JsonValueKind.Array)
                {
                    foreach (var item in fileArray.EnumerateArray())
                    {
                        var fileName = item.TryGetProperty("Name", out var name) ? name.GetString() : 
                                       item.TryGetProperty("name", out name) ? name.GetString() : 
                                       item.TryGetProperty("Path", out name) ? name.GetString() :
                                       item.TryGetProperty("path", out name) ? name.GetString() : null;
                        var size = item.TryGetProperty("Size", out var sz) ? sz.GetInt64() : 
                                   item.TryGetProperty("size", out sz) ? sz.GetInt64() : 0L;
                        
                        if (!string.IsNullOrEmpty(fileName) && IsModelFile(fileName))
                        {
                            files.Add(new ModelScopeFile
                            {
                                FileName = fileName,
                                DisplayName = fileName,
                                Size = FormatFileSize(size),
                                SizeBytes = size,
                                DownloadUrl = BuildDownloadUrl(fileName)
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "解析 JSON 响应失败");
            }
            
            if (files.Count == 0)
            {
                var defaultFiles = GetDefaultModelFiles();
                if (defaultFiles.Count > 0)
                {
                    files = defaultFiles;
                }
                // GGUF仓库不回退到硬编码列表，返回空列表让调用者处理
            }
            
            return files;
        }

        /// <summary>
        /// 判断是否为模型文件
        /// </summary>
        private bool IsModelFile(string fileName)
        {
            var modelExtensions = new[] { ".safetensors", ".pth", ".bin", ".pt", ".ckpt", ".onnx", ".gguf" };
            var lowerName = fileName.ToLower();
            return modelExtensions.Any(ext => lowerName.EndsWith(ext));
        }

        /// <summary>
        /// 格式化文件大小
        /// </summary>
        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            double size = bytes;
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            return $"{size:0.##} {sizes[order]}";
        }

        /// <summary>
        /// 下载模型文件
        /// </summary>
        public async Task<string> DownloadModelAsync(
            ModelScopeFile file,
            string saveDirectory,
            IProgress<DownloadProgressEventArgs>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var fullPath = Path.Combine(saveDirectory, file.FileName);
            
            try
            {
                _logger.LogInformation("开始下载: {FileName} -> {Path}", file.FileName, fullPath);
                _logger.LogInformation("下载 URL: {Url}", file.DownloadUrl);
                StatusChanged?.Invoke(this, $"正在下载: {file.FileName}");
                
                Directory.CreateDirectory(saveDirectory);
                
                // 检查文件是否已存在
                if (File.Exists(fullPath))
                {
                    var existingInfo = new FileInfo(fullPath);
                    if (file.SizeBytes > 0 && existingInfo.Length == file.SizeBytes)
                    {
                        _logger.LogInformation("文件已存在且完整: {Path} ({Size} bytes)", fullPath, existingInfo.Length);
                        StatusChanged?.Invoke(this, $"文件已存在，跳过下载: {file.FileName}");
                        
                        progress?.Report(new DownloadProgressEventArgs
                        {
                            FileName = file.FileName,
                            Progress = 100,
                            BytesDownloaded = existingInfo.Length,
                            TotalBytes = file.SizeBytes,
                            Status = "已完成（已存在）"
                        });
                        
                        return fullPath;
                    }
                    else
                    {
                        _logger.LogInformation("文件存在但大小不匹配，重新下载");
                        File.Delete(fullPath);
                    }
                }
                
                // 确保下载 URL 有效
                if (string.IsNullOrEmpty(file.DownloadUrl) || !Uri.IsWellFormedUriString(file.DownloadUrl, UriKind.Absolute))
                {
                    throw new InvalidOperationException($"无效的下载 URL: {file.DownloadUrl}");
                }
                
                // 使用 HttpClient 下载
                using var request = new HttpRequestMessage(HttpMethod.Get, file.DownloadUrl);
                request.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36");
                request.Headers.Add("Referer", $"{ModelScopeDownloadBase}/{_modelOwner}/{_modelName}/files");
                
                using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
                response.EnsureSuccessStatusCode();
                
                var totalBytes = response.Content.Headers.ContentLength ?? file.SizeBytes;
                var canReportProgress = totalBytes > 0;
                
                await using var contentStream = await response.Content.ReadAsStreamAsync(cancellationToken);
                await using var fileStream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true);
                
                var buffer = new byte[8192 * 4];
                long totalBytesRead = 0;
                int bytesRead;
                var watch = System.Diagnostics.Stopwatch.StartNew();
                
                while ((bytesRead = await contentStream.ReadAsync(buffer, cancellationToken)) > 0)
                {
                    await fileStream.WriteAsync(buffer.AsMemory(0, bytesRead), cancellationToken);
                    totalBytesRead += bytesRead;
                    
                    if (canReportProgress)
                    {
                        var progressPercent = (int)((totalBytesRead * 100) / totalBytes);
                        var speed = totalBytesRead / Math.Max(1, watch.Elapsed.TotalSeconds);
                        
                        var args = new DownloadProgressEventArgs
                        {
                            FileName = file.FileName,
                            Progress = progressPercent,
                            BytesDownloaded = totalBytesRead,
                            TotalBytes = totalBytes,
                            Speed = speed,
                            Status = $"下载中 {FormatFileSize(totalBytesRead)} / {FormatFileSize(totalBytes)} ({FormatFileSize((long)speed)}/s)"
                        };
                        
                        progress?.Report(args);
                        ProgressChanged?.Invoke(this, args);
                    }
                }
                
                watch.Stop();
                
                _logger.LogInformation("下载完成: {Path}", fullPath);
                StatusChanged?.Invoke(this, $"下载完成: {file.FileName}");
                
                DownloadCompleted?.Invoke(this, new DownloadCompletedEventArgs
                {
                    FileName = file.FileName,
                    FilePath = fullPath,
                    Success = true
                });
                
                return fullPath;
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("下载已取消: {FileName}", file.FileName);
                StatusChanged?.Invoke(this, "下载已取消");
                
                if (File.Exists(fullPath))
                {
                    try { File.Delete(fullPath); } catch { }
                }
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "下载失败: {FileName}", file.FileName);
                StatusChanged?.Invoke(this, $"下载失败: {ex.Message}");
                
                DownloadCompleted?.Invoke(this, new DownloadCompletedEventArgs
                {
                    FileName = file.FileName,
                    FilePath = fullPath,
                    Success = false,
                    ErrorMessage = ex.Message
                });
                
                if (File.Exists(fullPath))
                {
                    try { File.Delete(fullPath); } catch { }
                }
                throw;
            }
        }

        /// <summary>
        /// 下载多个模型文件
        /// </summary>
        public async Task<List<string>> DownloadModelFilesAsync(
            string saveDirectory,
            IProgress<DownloadProgressEventArgs>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var downloadedFiles = new List<string>();
            
            try
            {
                var files = await GetModelFilesAsync(cancellationToken);
                
                if (files.Count == 0)
                {
                    _logger.LogWarning("未找到可下载的模型文件");
                    StatusChanged?.Invoke(this, "未找到可下载的模型文件");
                    return downloadedFiles;
                }
                
                _logger.LogInformation("开始下载 {Count} 个文件", files.Count);
                StatusChanged?.Invoke(this, $"准备下载 {files.Count} 个文件...");
                
                long totalSize = files.Sum(f => f.SizeBytes);
                long downloadedSize = 0;
                var fileIndex = 0;
                
                foreach (var file in files)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    fileIndex++;
                    _logger.LogInformation("下载文件 {Index}/{Count}: {FileName}", fileIndex, files.Count, file.FileName);
                    StatusChanged?.Invoke(this, $"[{fileIndex}/{files.Count}] 正在下载: {file.FileName}");
                    
                    try
                    {
                        var fileProgress = new Progress<DownloadProgressEventArgs>(args =>
                        {
                            var overallProgress = totalSize > 0
                                ? (int)(((downloadedSize + args.BytesDownloaded) * 100) / totalSize)
                                : args.Progress;
                                
                            args.OverallFileIndex = fileIndex;
                            args.TotalFiles = files.Count;
                            args.OverallProgress = overallProgress;
                            args.Status = $"[{fileIndex}/{files.Count}] {args.Status}";
                            progress?.Report(args);
                        });
                        
                        var path = await DownloadModelAsync(file, saveDirectory, fileProgress, cancellationToken);
                        downloadedFiles.Add(path);
                        
                        if (File.Exists(path))
                        {
                            downloadedSize += new FileInfo(path).Length;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "下载文件失败: {FileName}", file.FileName);
                        StatusChanged?.Invoke(this, $"下载失败: {file.FileName} - {ex.Message}");
                        continue;
                    }
                }
                
                StatusChanged?.Invoke(this, $"下载完成! 共 {downloadedFiles.Count}/{files.Count} 个文件");
                _logger.LogInformation("模型下载完成: {Count} 个文件", downloadedFiles.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "下载模型时发生错误");
                StatusChanged?.Invoke(this, $"下载错误: {ex.Message}");
                throw;
            }
            
            return downloadedFiles;
        }

        /// <summary>
        /// 下载指定的模型文件列表（用户选择后）
        /// </summary>
        public async Task<List<string>> DownloadModelFilesAsync(
            List<ModelScopeFile> filesToDownload,
            string saveDirectory,
            IProgress<DownloadProgressEventArgs>? progress = null,
            CancellationToken cancellationToken = default)
        {
            var downloadedFiles = new List<string>();
            
            try
            {
                if (filesToDownload == null || filesToDownload.Count == 0)
                {
                    _logger.LogWarning("未指定要下载的模型文件");
                    StatusChanged?.Invoke(this, "未指定要下载的模型文件");
                    return downloadedFiles;
                }
                
                _logger.LogInformation("开始下载 {Count} 个文件", filesToDownload.Count);
                StatusChanged?.Invoke(this, $"准备下载 {filesToDownload.Count} 个文件...");
                
                long totalSize = filesToDownload.Sum(f => f.SizeBytes);
                long downloadedSize = 0;
                var fileIndex = 0;
                
                foreach (var file in filesToDownload)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    fileIndex++;
                    _logger.LogInformation("下载文件 {Index}/{Count}: {FileName}", fileIndex, filesToDownload.Count, file.FileName);
                    StatusChanged?.Invoke(this, $"[{fileIndex}/{filesToDownload.Count}] 正在下载: {file.FileName}");
                    
                    try
                    {
                        var fileProgress = new Progress<DownloadProgressEventArgs>(args =>
                        {
                            var overallProgress = totalSize > 0
                                ? (int)(((downloadedSize + args.BytesDownloaded) * 100) / totalSize)
                                : args.Progress;
                            
                            args.OverallFileIndex = fileIndex;
                            args.TotalFiles = filesToDownload.Count;
                            args.OverallProgress = overallProgress;
                            args.Status = $"[{fileIndex}/{filesToDownload.Count}] {args.Status}";
                            progress?.Report(args);
                        });
                        
                        var path = await DownloadModelAsync(file, saveDirectory, fileProgress, cancellationToken);
                        downloadedFiles.Add(path);
                        
                        if (File.Exists(path))
                        {
                            downloadedSize += new FileInfo(path).Length;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "下载文件失败: {FileName}", file.FileName);
                        StatusChanged?.Invoke(this, $"下载失败: {file.FileName} - {ex.Message}");
                        continue;
                    }
                }
                
                StatusChanged?.Invoke(this, $"下载完成! 共 {downloadedFiles.Count}/{filesToDownload.Count} 个文件");
                _logger.LogInformation("模型下载完成: {Count} 个文件", downloadedFiles.Count);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "下载模型时发生错误");
                StatusChanged?.Invoke(this, $"下载错误: {ex.Message}");
                throw;
            }
            
            return downloadedFiles;
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }

    /// <summary>
    /// ModelScope 文件信息
    /// </summary>
    public class ModelScopeFile
    {
        public string FileName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Size { get; set; } = string.Empty;
        public long SizeBytes { get; set; }
        public string DownloadUrl { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// 下载进度事件参数
    /// </summary>
    public class DownloadProgressEventArgs : EventArgs
    {
        public string FileName { get; set; } = string.Empty;
        public int Progress { get; set; }
        public int OverallProgress { get; set; }
        public long BytesDownloaded { get; set; }
        public long TotalBytes { get; set; }
        public double Speed { get; set; }
        public string Status { get; set; } = string.Empty;
        public int OverallFileIndex { get; set; }
        public int TotalFiles { get; set; }
    }

    /// <summary>
    /// 下载完成事件参数
    /// </summary>
    public class DownloadCompletedEventArgs : EventArgs
    {
        public string FileName { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
    }
}