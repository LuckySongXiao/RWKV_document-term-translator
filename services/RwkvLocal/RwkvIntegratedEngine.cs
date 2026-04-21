using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// RWKV 集成引擎 - 统一接口，自动选择原生模式或进程模式
    /// </summary>
    public class RwkvIntegratedEngine : IDisposable
    {
        public enum EngineMode
        {
            Auto,           // 自动选择（优先原生）
            Native,         // 原生 DLL 模式
            Process         // 独立进程模式
        }

        private readonly ILogger _logger;
        private EngineMode _mode;
        private readonly string _modelPath;
        private readonly string _vocabPath;
        private readonly int _gpuId;
        private readonly int _port;

        // 两种模式的引擎实例
        private RwkvNativeEngine? _nativeEngine;
        private RwkvProcessService? _processService;
        private ModelManagementService? _modelService;

        private bool _isRunning = false;
        private bool _disposed = false;

        /// <summary>
        /// 当前运行模式
        /// </summary>
        public EngineMode CurrentMode => _mode;

        /// <summary>
        /// 是否正在运行
        /// </summary>
        public bool IsRunning => _isRunning;

        /// <summary>
        /// 模型名称
        /// </summary>
        public string ModelName { get; private set; } = "";

        /// <summary>
        /// 创建集成引擎实例
        /// </summary>
        public RwkvIntegratedEngine(
            ILogger logger,
            string modelPath,
            string vocabPath,
            int gpuId = 0,
            int port = 8000,
            EngineMode mode = EngineMode.Auto)
        {
            _logger = logger;
            _modelPath = modelPath;
            _vocabPath = vocabPath;
            _gpuId = gpuId;
            _port = port;
            _mode = mode;
        }

        /// <summary>
        /// 启动引擎
        /// </summary>
        public async Task<bool> StartAsync()
        {
            if (_isRunning)
            {
                _logger.LogWarning("引擎已在运行");
                return true;
            }

            // 自动选择模式
            if (_mode == EngineMode.Auto)
            {
                _mode = DetermineBestMode();
            }

            _logger.LogInformation($"使用 {_mode} 模式启动 RWKV 引擎");

            bool success = false;

            if (_mode == EngineMode.Native)
            {
                success = await StartNativeModeAsync();
                if (!success)
                {
                    _logger.LogWarning("原生模式启动失败，尝试切换到进程模式");
                    _mode = EngineMode.Process;
                    success = await StartProcessModeAsync();
                }
            }
            else
            {
                success = await StartProcessModeAsync();
            }

            _isRunning = success;
            return success;
        }

        /// <summary>
        /// 确定最佳运行模式
        /// </summary>
        private EngineMode DetermineBestMode()
        {
            // 检查原生 DLL 是否存在
            var dllPath = FindNativeDll();
            if (!string.IsNullOrEmpty(dllPath))
            {
                _logger.LogInformation($"找到原生 DLL: {dllPath}");
                
                // 检查 GPU 是否可用
                if (RwkvNativeEngine.IsGpuAvailable)
                {
                    _logger.LogInformation("检测到可用 GPU，优先使用原生模式");
                    return EngineMode.Native;
                }
            }

            _logger.LogInformation("使用进程模式");
            return EngineMode.Process;
        }

        /// <summary>
        /// 查找原生 DLL
        /// </summary>
        private string? FindNativeDll()
        {
            var baseDir = PathHelper.GetSafeBaseDirectory();
            var projectRoot = PathHelper.FindProjectRoot(baseDir);
            var possiblePaths = new[]
            {
                Path.Combine(baseDir, "rwkv_native.dll"),
                Path.Combine(baseDir, "rwkv_lightning_libtorch", "rwkv_native.dll"),
                Path.Combine(baseDir, "rwkv_lightning_libtorch", "dist_native", "bin", "rwkv_native.dll"),
                Path.Combine(baseDir, "rwkv_lightning_libtorch", "build_native", "Release", "rwkv_native.dll"),
                Path.Combine(baseDir, "rwkv_lightning_libtorch", "build", "Release", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch", "dist_native", "bin", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch", "build_native", "Release", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch", "build", "Release", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch", "build", "Debug", "rwkv_native.dll"),
                Path.Combine(projectRoot, "transplantation", "dist_native", "bin", "rwkv_native.dll"),
                Path.Combine(projectRoot, "transplantation", "rwkv_lightning_libtorch", "dist_native", "bin", "rwkv_native.dll"),
                Path.Combine(projectRoot, "transplantation", "rwkv_lightning_libtorch", "build_native", "Release", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch_win", "rwkv_native.dll"),
                Path.Combine(projectRoot, "rwkv_lightning_libtorch_win", "rwkv_native.dll"),
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            return null;
        }

        /// <summary>
        /// 启动原生模式
        /// </summary>
        private async Task<bool> StartNativeModeAsync()
        {
            try
            {
                _nativeEngine = new RwkvNativeEngine(_logger);
                
                if (!_nativeEngine.LoadModel(_modelPath, _vocabPath, _gpuId))
                {
                    _logger.LogError("原生模式：模型加载失败");
                    return false;
                }

                ModelName = _nativeEngine.ModelName;
                _logger.LogInformation($"原生模式启动成功: {ModelName}");
                return true;
            }
            catch (DllNotFoundException ex)
            {
                _logger.LogWarning($"原生 DLL 未找到: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "原生模式启动失败");
                return false;
            }
        }

        /// <summary>
        /// 启动进程模式
        /// </summary>
        private async Task<bool> StartProcessModeAsync()
        {
            try
            {
                var gpuService = new GpuResourceService(_logger);
                _modelService = new ModelManagementService(_logger);
                var benchmarkCache = new ModelBenchmarkCache(_logger);
                var concurrencyCalc = new ConcurrencyCalculator(_logger, gpuService, _modelService, benchmarkCache);
                _processService = new RwkvProcessService(_logger, _modelService, gpuService, concurrencyCalc);

                var modelInfo = new ModelInfo
                {
                    FilePath = _modelPath
                };

                // 获取 GPU 信息
                var gpus = await gpuService.GetAllGpusAsync();
                var gpu = gpus.FirstOrDefault(g => g.Index == _gpuId) ?? gpus.FirstOrDefault();
                
                if (gpu == null)
                {
                    _logger.LogError("未找到可用的 GPU");
                    return false;
                }

                var success = await _processService.StartAsync(modelInfo, gpu, _port);
                
                if (success)
                {
                    ModelName = modelInfo.FileName;
                    _logger.LogInformation($"进程模式启动成功: {ModelName}, 端口: {_processService.Status.Port}");
                    return true;
                }

                _logger.LogError("进程模式启动失败");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "进程模式启动失败");
                return false;
            }
        }

        /// <summary>
        /// 翻译文本
        /// </summary>
        public async Task<string> TranslateAsync(string prompt, RwkvGenerateOptions? options = null)
        {
            if (!_isRunning)
            {
                throw new InvalidOperationException("引擎未启动");
            }

            if (_mode == EngineMode.Native && _nativeEngine != null)
            {
                // 原生模式：直接调用
                return await Task.Run(() => _nativeEngine.Translate(prompt, options));
            }
            else if (_mode == EngineMode.Process && _processService != null)
            {
                // 进程模式：通过 HTTP API 调用
                return await _processService.TranslateAsync(prompt);
            }

            throw new InvalidOperationException("引擎状态异常");
        }

        /// <summary>
        /// 批量翻译
        /// </summary>
        public async Task<string[]> BatchTranslateAsync(string[] prompts, RwkvGenerateOptions? options = null)
        {
            if (!_isRunning)
            {
                throw new InvalidOperationException("引擎未启动");
            }

            if (_mode == EngineMode.Native && _nativeEngine != null)
            {
                return await Task.Run(() => _nativeEngine.BatchTranslate(prompts, options));
            }
            else if (_mode == EngineMode.Process && _processService != null)
            {
                // 进程模式使用顺序调用（也可以实现批量 API）
                var results = new string[prompts.Length];
                for (int i = 0; i < prompts.Length; i++)
                {
                    results[i] = await _processService.TranslateAsync(prompts[i]);
                }
                return results;
            }

            throw new InvalidOperationException("引擎状态异常");
        }

        /// <summary>
        /// 重置对话状态
        /// </summary>
        public void ResetState()
        {
            if (_mode == EngineMode.Native && _nativeEngine != null)
            {
                _nativeEngine.ResetState();
            }
            // 进程模式通过新的 session ID 实现
        }

        /// <summary>
        /// 停止引擎
        /// </summary>
        public async Task StopAsync()
        {
            if (!_isRunning) return;

            if (_nativeEngine != null)
            {
                _nativeEngine.Dispose();
                _nativeEngine = null;
            }

            if (_processService != null)
            {
                await _processService.StopAsync();
            }

            _isRunning = false;
            _logger.LogInformation("RWKV 引擎已停止");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _nativeEngine?.Dispose();
                _processService?.Dispose();
                _disposed = true;
            }
        }
    }
}
