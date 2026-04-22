using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// llama.cpp 推理进程管理服务
    /// 负责启动、停止、监控 llama-server 进程
    /// llama-server 提供 OpenAI 兼容的 HTTP API
    /// </summary>
    public class LlamaCppProcessService : IDisposable
    {
        private static readonly Encoding OutputEncoding;

        static LlamaCppProcessService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            OutputEncoding = Encoding.UTF8;
        }

        private readonly ILogger _logger;
        private readonly ModelManagementService _modelService;
        private readonly GpuResourceService _gpuService;
        
        private Process? _process;
        private readonly HttpClient _httpClient;
        private readonly CancellationTokenSource _monitorCts;
        private Task? _monitorTask;
        private bool _disposed;

        public RwkvServiceStatus Status { get; }
        
        public event EventHandler<ServiceStatusEventArgs>? StatusChanged;

        public LlamaCppProcessService(
            ILogger logger,
            ModelManagementService modelService,
            GpuResourceService gpuService)
        {
            _logger = logger;
            _modelService = modelService;
            _gpuService = gpuService;
            
            Status = new RwkvServiceStatus();
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            _monitorCts = new CancellationTokenSource();
        }

        /// <summary>
        /// 启动 llama-server 推理服务
        /// </summary>
        /// <param name="model">模型信息（必须为GGUF格式）</param>
        /// <param name="gpu">GPU信息</param>
        /// <param name="toolName">工具名称（llama-cuda, llama-cpu, llama-vulkan）</param>
        /// <param name="preferredPort">首选端口</param>
        /// <param name="chatTemplate">对话模板名称（如 rwkv-world、chatml、llama3），默认 rwkv-world</param>
        public async Task<bool> StartAsync(ModelInfo model, GpuInfo? gpu, string toolName = "llama-cuda", int? preferredPort = null, string chatTemplate = null)
        {
            if (Status.State == ServiceState.Running || Status.State == ServiceState.Starting)
            {
                _logger.LogWarning("llama-server 服务已在运行或正在启动中");
                return false;
            }

            try
            {
                UpdateStatus(ServiceState.Starting);
                
                // 检查推理工具
                var toolPath = _modelService.GetToolPath(toolName);
                if (!_modelService.ToolExists(toolName))
                {
                    Status.ErrorMessage = $"推理工具不存在: {toolPath}";
                    _logger.LogError(Status.ErrorMessage);
                    UpdateStatus(ServiceState.Error);
                    return false;
                }

                // 检查模型格式
                if (model.Format != ModelFormat.GGUF)
                {
                    Status.ErrorMessage = $"llama.cpp 仅支持 GGUF 格式模型，当前模型格式: {model.Format}";
                    _logger.LogError(Status.ErrorMessage);
                    UpdateStatus(ServiceState.Error);
                    return false;
                }

                // 获取模型路径
                var modelPath = _modelService.GetModelPath(model);
                if (!File.Exists(modelPath))
                {
                    Status.ErrorMessage = $"模型文件不存在: {modelPath}";
                    _logger.LogError(Status.ErrorMessage);
                    UpdateStatus(ServiceState.Error);
                    return false;
                }

                // llama-server.exe 不支持中文路径
                // 将模型文件硬链接到纯ASCII临时目录，避免路径编码问题
                var safeModelPath = PrepareAsciiModelPath(modelPath);
                _logger.LogInformation("模型路径: {Original} -> {Safe}", modelPath, safeModelPath);

                // 查找可用端口
                Status.Port = preferredPort ?? await FindAvailablePortAsync(8080);
                Status.SelectedGpu = gpu;
                Status.LoadedModel = model;

                // 构建 llama-server 启动参数
                // llama-server 提供 OpenAI 兼容 API
                // --chat-template: 对话模板（如 rwkv-world、chatml、llama3）
                var effectiveChatTemplate = string.IsNullOrWhiteSpace(chatTemplate) ? "rwkv-world" : chatTemplate;
                var arguments = $"-m \"{safeModelPath}\" " +
                               $"--port {Status.Port} " +
                               $"--host 127.0.0.1 " +
                               $"-c 8192 " +                       // 上下文长度
                               $"-t 4 " +                          // 线程数
                               $"--parallel 4 " +                  // 并行请求数
                               $"--chat-template {effectiveChatTemplate}";  // 对话模板

                // GPU卸载层数设置
                if ((toolName == "llama-cuda" || toolName == "llama-sycl" || toolName == "llama-vulkan") && gpu != null)
                {
                    arguments += " -ngl all";  // 全部层放到GPU
                }
                else
                {
                    // CPU 模式，不使用 GPU 层
                    arguments += " -ngl 0";
                }

                _logger.LogInformation("启动 llama-server: {ToolPath} {Arguments}", toolPath, arguments);

                // 获取工作目录（llama.cpp的DLL在同一目录下）
                var workingDir = _modelService.GetLlamaCppWorkingDirectory(toolName);
                
                // 工作目录和工具路径也转换为短路径，避免中文路径问题
                var safeWorkingDir = GetShortPath(workingDir);
                var safeToolPath = GetShortPath(toolPath);
                
                // 如果短路径仍含非ASCII字符，使用临时ASCII目录
                if (!IsAsciiOnly(safeWorkingDir))
                {
                    safeWorkingDir = PrepareAsciiDirectory(safeWorkingDir, "llama_work");
                }
                if (!IsAsciiOnly(safeToolPath))
                {
                    safeToolPath = PrepareAsciiDirectory(safeToolPath, "llama_tool");
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = safeToolPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = OutputEncoding,
                    StandardErrorEncoding = OutputEncoding,
                    CreateNoWindow = true,
                    WorkingDirectory = safeWorkingDir
                };

                // 设置CUDA_VISIBLE_DEVICES环境变量
                if (gpu != null && toolName == "llama-cuda")
                {
                    startInfo.EnvironmentVariables["CUDA_VISIBLE_DEVICES"] = gpu.Index.ToString();
                }

                // 创建并启动进程
                _process = new Process { StartInfo = startInfo };
                
                _process.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        _logger.LogDebug("[llama-server] {Data}", e.Data);
                    }
                };

                _process.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        _logger.LogWarning("[llama-server-ERR] {Data}", e.Data);
                    }
                };

                _process.Exited += (s, e) =>
                {
                    _logger.LogInformation("llama-server 进程已退出，退出码: {ExitCode}", _process?.ExitCode);
                    
                    if (Status.State == ServiceState.Running)
                    {
                        UpdateStatus(ServiceState.Stopped);
                    }
                };

                _process.EnableRaisingEvents = true;
                _process.Start();
                _process.BeginOutputReadLine();
                _process.BeginErrorReadLine();

                Status.ProcessId = _process.Id;
                Status.StartTime = DateTime.Now;

                // 等待服务就绪
                var ready = await WaitForServiceReadyAsync(TimeSpan.FromSeconds(120));

                if (ready)
                {
                    UpdateStatus(ServiceState.Running);
                    
                    // 设置默认并发数
                    Status.MaxConcurrency = 4;
                    
                    // 启动监控任务
                    _monitorTask = MonitorServiceAsync(_monitorCts.Token);
                    
                    _logger.LogInformation("llama-server 启动成功，端口: {Port}", Status.Port);
                    return true;
                }
                else
                {
                    Status.ErrorMessage = "llama-server 启动超时";
                    _logger.LogError(Status.ErrorMessage);
                    
                    try { _process?.Kill(); } catch { }
                    _process = null;
                    
                    UpdateStatus(ServiceState.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Status.ErrorMessage = $"启动 llama-server 时发生异常: {ex.Message}";
                _logger.LogError(ex, "启动 llama-server 时发生异常");
                UpdateStatus(ServiceState.Error);
                return false;
            }
        }

        /// <summary>
        /// 停止推理服务
        /// </summary>
        public async Task StopAsync()
        {
            if (Status.State == ServiceState.Stopped || Status.State == ServiceState.Stopping)
            {
                return;
            }

            try
            {
                UpdateStatus(ServiceState.Stopping);

                _monitorCts.Cancel();

                if (_process != null && !_process.HasExited)
                {
                    _logger.LogInformation("正在停止 llama-server...");

                    _process.CloseMainWindow();
                    
                    if (!_process.WaitForExit(5000))
                    {
                        _logger.LogWarning("进程未响应，强制终止");
                        _process.Kill();
                    }
                }

                _process?.Dispose();
                _process = null;
                
                Status.Reset();
                UpdateStatus(ServiceState.Stopped);
                
                _logger.LogInformation("llama-server 已停止");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停止 llama-server 时发生异常");
                UpdateStatus(ServiceState.Error);
            }
        }

        /// <summary>
        /// 检查服务是否就绪（llama-server 的健康检查端点）
        /// </summary>
        public async Task<bool> CheckHealthAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{Status.ApiUrl}/health");
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 等待服务就绪
        /// </summary>
        private async Task<bool> WaitForServiceReadyAsync(TimeSpan timeout)
        {
            var startTime = DateTime.Now;
            var checkInterval = TimeSpan.FromMilliseconds(1000);

            while (DateTime.Now - startTime < timeout)
            {
                try
                {
                    var response = await _httpClient.GetAsync($"{Status.ApiUrl}/health");
                    
                    if (response.IsSuccessStatusCode)
                    {
                        return true;
                    }
                }
                catch (HttpRequestException)
                {
                }
                catch (TaskCanceledException)
                {
                }

                await Task.Delay(checkInterval);
            }

            return false;
        }

        /// <summary>
        /// 监控服务状态
        /// </summary>
        private async Task MonitorServiceAsync(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested && Status.State == ServiceState.Running)
            {
                try
                {
                    var healthOk = await CheckHealthAsync();
                    Status.LastHeartbeat = DateTime.Now;

                    if (Status.SelectedGpu != null)
                    {
                        Status.CurrentGpuStatus = await _gpuService.RefreshGpuStatusAsync(Status.SelectedGpu.Index);
                    }

                    StatusChanged?.Invoke(this, new ServiceStatusEventArgs { Status = Status });

                    await Task.Delay(TimeSpan.FromSeconds(10), cancellationToken);
                }
                catch (TaskCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "监控 llama-server 状态时发生异常");
                    await Task.Delay(TimeSpan.FromSeconds(5), cancellationToken);
                }
            }
        }

        /// <summary>
        /// 查找可用端口
        /// </summary>
        private async Task<int> FindAvailablePortAsync(int startPort)
        {
            return await Task.Run(() =>
            {
                for (int port = startPort; port < startPort + 100; port++)
                {
                    try
                    {
                        var listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, port);
                        listener.Start();
                        listener.Stop();
                        return port;
                    }
                    catch
                    {
                    }
                }
                return startPort;
            });
        }

        private void UpdateStatus(ServiceState state)
        {
            Status.State = state;
            StatusChanged?.Invoke(this, new ServiceStatusEventArgs { Status = Status });
        }

        /// <summary>
        /// 将模型文件硬链接到纯ASCII临时目录，解决 llama-server 不支持中文路径的问题
        /// 优先在模型所在磁盘创建临时目录（支持硬链接，不占额外空间）
        /// </summary>
        private string PrepareAsciiModelPath(string originalPath)
        {
            try
            {
                // 如果路径已经是纯ASCII，直接返回
                if (IsAsciiOnly(originalPath))
                {
                    return originalPath;
                }

                var fileName = Path.GetFileName(originalPath);
                var modelDir = Path.GetDirectoryName(originalPath) ?? "";

                // 优先在模型所在磁盘的根目录下创建临时目录（支持硬链接）
                var modelDrive = Path.GetPathRoot(modelDir);
                var tempDir = Path.Combine(modelDrive, "llama_models_tmp");
                
                try
                {
                    Directory.CreateDirectory(tempDir);
                }
                catch
                {
                    // 如果无法在模型磁盘创建，回退到系统临时目录
                    tempDir = Path.Combine(Path.GetTempPath(), "llama_models");
                    Directory.CreateDirectory(tempDir);
                }

                var targetPath = Path.Combine(tempDir, fileName);

                // 如果目标已存在且是同一文件（大小和修改时间匹配），直接使用
                if (File.Exists(targetPath))
                {
                    var origInfo = new FileInfo(originalPath);
                    var targetInfo = new FileInfo(targetPath);
                    if (origInfo.Length == targetInfo.Length && 
                        origInfo.LastWriteTimeUtc == targetInfo.LastWriteTimeUtc)
                    {
                        return targetPath;
                    }
                    // 文件不匹配，删除旧的
                    File.Delete(targetPath);
                }

                // 尝试创建硬链接（不占用额外磁盘空间）
                if (CreateHardLink(targetPath, originalPath, IntPtr.Zero))
                {
                    _logger.LogInformation("已创建硬链接: {Target} -> {Original}", targetPath, originalPath);
                    return targetPath;
                }

                // 硬链接失败（可能跨磁盘），复制文件
                _logger.LogWarning("硬链接创建失败，复制文件到临时目录（可能需要一些时间）");
                File.Copy(originalPath, targetPath, overwrite: true);
                return targetPath;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "准备ASCII模型路径失败，使用原始路径");
                return originalPath;
            }
        }

        /// <summary>
        /// 将目录内容符号链接到纯ASCII临时目录
        /// </summary>
        private string PrepareAsciiDirectory(string originalDir, string prefix)
        {
            try
            {
                if (IsAsciiOnly(originalDir))
                {
                    return originalDir;
                }

                var tempDir = Path.Combine(Path.GetTempPath(), "llama_work", prefix);
                
                // 如果已存在，直接使用
                if (Directory.Exists(tempDir))
                {
                    return tempDir;
                }

                // 创建目录符号链接
                Directory.CreateDirectory(tempDir);
                
                // 复制必要的DLL文件（llama.cpp工作目录需要DLL）
                if (Directory.Exists(originalDir))
                {
                    foreach (var file in Directory.GetFiles(originalDir))
                    {
                        var destFile = Path.Combine(tempDir, Path.GetFileName(file));
                        if (!File.Exists(destFile))
                        {
                            File.Copy(file, destFile, overwrite: false);
                        }
                    }
                }

                return tempDir;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "准备ASCII工作目录失败，使用原始路径");
                return originalDir;
            }
        }

        /// <summary>
        /// 检查字符串是否只包含ASCII字符
        /// </summary>
        private static bool IsAsciiOnly(string value)
        {
            if (string.IsNullOrEmpty(value)) return true;
            foreach (char c in value)
            {
                if (c > 127) return false;
            }
            return true;
        }

        /// <summary>
        /// 获取 Windows 短路径（8.3格式）
        /// </summary>
        private static string GetShortPath(string longPath)
        {
            try
            {
                var pathExists = File.Exists(longPath) || Directory.Exists(longPath);
                if (!pathExists) return longPath;

                var buffer = new char[512];
                var result = GetShortPathName(longPath, buffer, 512);
                if (result > 0 && result < 512)
                {
                    var shortPath = new string(buffer, 0, (int)result);
                    if (!string.IsNullOrEmpty(shortPath))
                    {
                        return shortPath;
                    }
                }
            }
            catch
            {
            }
            return longPath;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetShortPathName(string path, char[] buffer, int bufferLength);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool CreateHardLink(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                _monitorCts.Cancel();
                
                try { _process?.Kill(); } catch { }
                _process?.Dispose();
                
                _httpClient.Dispose();
                _monitorCts.Dispose();
            }
        }
    }
}
