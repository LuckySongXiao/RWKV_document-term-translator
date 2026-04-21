using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// RWKV推理进程管理服务
    /// 负责启动、停止、监控rwkv_lightning进程
    /// </summary>
    public class RwkvProcessService : IDisposable
    {
        private static readonly Encoding OutputEncoding;

        static RwkvProcessService()
        {
            // 注册编码提供程序以支持各种编码
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            // benchmark.exe 输出 UTF-8 编码
            OutputEncoding = Encoding.UTF8;
        }
        private readonly ILogger _logger;
        private readonly ModelManagementService _modelService;
        private readonly GpuResourceService _gpuService;
        private readonly ConcurrencyCalculator _concurrencyCalculator;
        
        private Process? _process;
        private readonly HttpClient _httpClient;
        private readonly CancellationTokenSource _monitorCts;
        private Task? _monitorTask;
        private bool _disposed;
        private bool _isBenchmarkMode;
        private readonly System.Text.StringBuilder _benchmarkOutput = new System.Text.StringBuilder();

        public RwkvServiceStatus Status { get; }
        
        public event EventHandler<ServiceStatusEventArgs>? StatusChanged;
        
        /// <summary>
        /// Benchmark测试结果输出事件
        /// </summary>
        public event EventHandler<string>? BenchmarkOutputReceived;

        public RwkvProcessService(
            ILogger logger,
            ModelManagementService modelService,
            GpuResourceService gpuService,
            ConcurrencyCalculator concurrencyCalculator)
        {
            _logger = logger;
            _modelService = modelService;
            _gpuService = gpuService;
            _concurrencyCalculator = concurrencyCalculator;
            
            Status = new RwkvServiceStatus();
            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            _monitorCts = new CancellationTokenSource();
        }

        /// <summary>
        /// 启动推理服务
        /// </summary>
        /// <param name="model">模型信息</param>
        /// <param name="gpu">GPU信息</param>
        /// <param name="preferredPort">首选端口</param>
        /// <param name="toolName">推理工具名称（rwkv_lightning 或 benchmark）</param>
        /// <param name="enableBszTest">是否启用 BSZ 上限测试</param>
        /// <param name="bszIncrement">BSZ 测试每次递增的数量</param>
        public async Task<bool> StartAsync(ModelInfo model, GpuInfo gpu, int? preferredPort = null, string toolName = "rwkv_lightning", bool enableBszTest = false, int bszIncrement = 10)
        {
            if (Status.State == ServiceState.Running || Status.State == ServiceState.Starting)
            {
                _logger.LogWarning("服务已在运行或正在启动中");
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

                // 检查词汇表文件
                if (!_modelService.VocabFileExists)
                {
                    Status.ErrorMessage = $"词汇表文件不存在: {_modelService.VocabFilePath}";
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

                // 查找可用端口
                Status.Port = preferredPort ?? await FindAvailablePortAsync(8000);

                // 保存参数用于后续并发测试
                Status.SelectedGpu = gpu;
                Status.LoadedModel = model;

                // 构建启动参数（根据工具类型使用不同的参数）
                string arguments;
                bool providesApi = true; // 是否提供API服务

                if (toolName == "benchmark")
                {
                    // benchmark.exe 使用不同的参数，不提供API服务
                    _isBenchmarkMode = true;
                    _benchmarkOutput.Clear();
                    
                    arguments = $"--weights \"{modelPath}\" " +
                               $"--vocab \"{_modelService.VocabFilePath}\" " +
                               $"--decode-prompt \"User: \" " +
                               $"--decode-steps 10 " +
                               $"--decode-temp 1.0";
                    providesApi = false;
                    _logger.LogInformation($"启动RWKV性能测试: {toolPath} {arguments}");
                }
                else
                {
                    // rwkv_lightning.exe 提供HTTP API服务
                    _isBenchmarkMode = false;
                    
                    // rwkv_lightning.exe 提供HTTP API服务
                    arguments = $"--model-path \"{modelPath}\" " +
                               $"--vocab-path \"{_modelService.VocabFilePath}\" " +
                               $"--port {Status.Port}";
                    _logger.LogInformation($"启动RWKV推理服务: {toolPath} {arguments}");
                }

                // 设置环境变量指定GPU
                var startInfo = new ProcessStartInfo
                {
                    FileName = toolPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    // 使用 UTF-8 编码读取 C++ 程序输出
                    StandardOutputEncoding = OutputEncoding,
                    StandardErrorEncoding = OutputEncoding,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(toolPath)
                };

                // 设置CUDA_VISIBLE_DEVICES环境变量
                startInfo.EnvironmentVariables["CUDA_VISIBLE_DEVICES"] = gpu.Index.ToString();

                // 创建并启动进程
                _process = new Process { StartInfo = startInfo };
                
                _process.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        _logger.LogDebug($"[RWKV] {e.Data}");
                        
                        // benchmark模式下收集输出
                        if (_isBenchmarkMode)
                        {
                            _benchmarkOutput.AppendLine(e.Data);
                        }
                    }
                };

                _process.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        _logger.LogWarning($"[RWKV-ERR] {e.Data}");
                        
                        // benchmark模式下也收集错误输出
                        if (_isBenchmarkMode)
                        {
                            _benchmarkOutput.AppendLine(e.Data);
                        }
                    }
                };

                _process.Exited += (s, e) =>
                {
                    _logger.LogInformation($"RWKV进程已退出，退出码: {_process?.ExitCode}");
                    
                    // benchmark模式下触发输出事件
                    if (_isBenchmarkMode && _benchmarkOutput.Length > 0)
                    {
                        var output = _benchmarkOutput.ToString();
                        BenchmarkOutputReceived?.Invoke(this, output);
                        _benchmarkOutput.Clear();
                    }
                    
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

                // benchmark.exe 不提供API服务，启动后立即返回
                if (!providesApi)
                {
                    UpdateStatus(ServiceState.Running);
                    _logger.LogInformation($"RWKV性能测试进程启动成功，PID: {Status.ProcessId}");
                    return true;
                }

                // 等待服务就绪（仅适用于 rwkv_lightning.exe）
                var ready = await WaitForServiceReadyAsync(TimeSpan.FromSeconds(60));

                if (ready)
                {
                    UpdateStatus(ServiceState.Running);
                    
                    // 启动监控任务
                    _monitorTask = MonitorServiceAsync(_monitorCts.Token);
                    
                    // 服务启动后进行并发测试
                    if (toolName != "benchmark" && enableBszTest)
                    {
                        var apiUrl = $"http://127.0.0.1:{Status.Port}/translate/v1/batch-translate";
                        Status.MaxConcurrency = await _concurrencyCalculator.CalculateMaxConcurrencyAsync(
                            gpu, model, apiUrl, enableBszTest, bszIncrement);
                        
                        _logger.LogInformation($"RWKV推理服务启动成功，端口: {Status.Port}, 并发数: {Status.MaxConcurrency}");
                    }
                    else
                    {
                        _logger.LogInformation($"RWKV推理服务启动成功，端口: {Status.Port}");
                    }
                    
                    return true;
                }
                else
                {
                    Status.ErrorMessage = "服务启动超时";
                    _logger.LogError(Status.ErrorMessage);
                    
                    try { _process?.Kill(); } catch { }
                    _process = null;
                    
                    UpdateStatus(ServiceState.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Status.ErrorMessage = $"启动服务时发生异常: {ex.Message}";
                _logger.LogError(ex, "启动RWKV推理服务时发生异常");
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

                // 停止监控
                _monitorCts.Cancel();

                if (_process != null && !_process.HasExited)
                {
                    _logger.LogInformation("正在停止RWKV推理服务...");

                    // 尝试优雅关闭
                    _process.CloseMainWindow();
                    
                    // 等待进程退出
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
                
                _logger.LogInformation("RWKV推理服务已停止");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停止RWKV推理服务时发生异常");
                UpdateStatus(ServiceState.Error);
            }
        }

        /// <summary>
        /// 检查服务是否就绪
        /// </summary>
        public async Task<bool> CheckHealthAsync()
        {
            try
            {
                var response = await _httpClient.GetAsync($"{Status.ApiUrl}/v1/models");
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
            var checkInterval = TimeSpan.FromMilliseconds(500);

            while (DateTime.Now - startTime < timeout)
            {
                try
                {
                    var response = await _httpClient.GetAsync($"{Status.ApiUrl}/v1/models");
                    
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
                    // 定期检查心跳
                    var healthOk = await CheckHealthAsync();
                    Status.LastHeartbeat = DateTime.Now;

                    // 更新GPU状态
                    if (Status.SelectedGpu != null)
                    {
                        Status.CurrentGpuStatus = await _gpuService.RefreshGpuStatusAsync(Status.SelectedGpu.Index);
                    }

                    // 触发状态更新事件
                    StatusChanged?.Invoke(this, new ServiceStatusEventArgs { Status = Status });

                    await Task.Delay(TimeSpan.FromSeconds(10), cancellationToken);
                }
                catch (TaskCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "监控服务状态时发生异常");
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
                        // 端口被占用，继续尝试下一个
                    }
                }
                return startPort;
            });
        }

        /// <summary>
        /// 更新服务状态
        /// </summary>
        private void UpdateStatus(ServiceState state)
        {
            Status.State = state;
            StatusChanged?.Invoke(this, new ServiceStatusEventArgs { Status = Status });
        }

        /// <summary>
        /// 翻译文本（通过 HTTP API）
        /// </summary>
        public async Task<string> TranslateAsync(string text, string sourceLang = "zh", string targetLang = "en")
        {
            if (!Status.IsRunning)
            {
                throw new InvalidOperationException("服务未运行");
            }

            try
            {
                var requestBody = new
                {
                    source_lang = sourceLang,
                    target_lang = targetLang,
                    text_list = new[] { text }
                };

                var json = System.Text.Json.JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(
                    $"{Status.ApiUrl}/translate/v1/batch-translate",
                    content
                );

                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync();
                var result = System.Text.Json.JsonSerializer.Deserialize<TranslateResponse>(responseJson);

                return result?.translations?.FirstOrDefault() ?? "";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "翻译请求失败");
                throw;
            }
        }

        /// <summary>
        /// 批量翻译
        /// </summary>
        public async Task<string[]> BatchTranslateAsync(string[] texts, string sourceLang = "zh", string targetLang = "en")
        {
            if (!Status.IsRunning)
            {
                throw new InvalidOperationException("服务未运行");
            }

            try
            {
                var requestBody = new
                {
                    source_lang = sourceLang,
                    target_lang = targetLang,
                    text_list = texts
                };

                var json = System.Text.Json.JsonSerializer.Serialize(requestBody);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");

                var response = await _httpClient.PostAsync(
                    $"{Status.ApiUrl}/translate/v1/batch-translate",
                    content
                );

                response.EnsureSuccessStatusCode();

                var responseJson = await response.Content.ReadAsStringAsync();
                var result = System.Text.Json.JsonSerializer.Deserialize<TranslateResponse>(responseJson);

                return result?.translations ?? Array.Empty<string>();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "批量翻译请求失败");
                throw;
            }
        }

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

    /// <summary>
    /// 服务状态变更事件参数
    /// </summary>
    public class ServiceStatusEventArgs : EventArgs
    {
        public RwkvServiceStatus Status { get; set; } = new();
    }

    /// <summary>
    /// 翻译响应
    /// </summary>
    internal class TranslateResponse
    {
        public string[]? translations { get; set; }
    }
}
