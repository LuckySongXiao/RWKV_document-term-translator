using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// RWKV 原生推理引擎 - P/Invoke 封装
    /// 直接调用 C++ DLL 进行推理，无需启动独立进程
    /// </summary>
    public class RwkvNativeEngine : IDisposable
    {
        #region P/Invoke 声明

        private const string DLL_NAME = "rwkv_native";

        // 错误码
        private const int RWKV_SUCCESS = 0;
        private const int RWKV_ERROR_INVALID_HANDLE = -1;
        private const int RWKV_ERROR_MODEL_LOAD_FAILED = -2;
        private const int RWKV_ERROR_INFERENCE_FAILED = -4;

        // 句柄类型
        private IntPtr _modelHandle = IntPtr.Zero;
        private IntPtr _stateHandle = IntPtr.Zero;

        // 回调委托类型
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void LogCallback([MarshalAs(UnmanagedType.LPStr)] string message);

        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate void StreamCallback([MarshalAs(UnmanagedType.LPStr)] string token, IntPtr userData);

        // P/Invoke 函数声明
        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern void rwkv_set_log_callback(LogCallback callback);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_is_gpu_available();

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_get_gpu_info(StringBuilder gpu_info, int buffer_size);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr rwkv_load_model(
            [MarshalAs(UnmanagedType.LPStr)] string model_path,
            [MarshalAs(UnmanagedType.LPStr)] string vocab_path,
            int gpu_id);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern void rwkv_unload_model(IntPtr handle);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_get_model_info(
            IntPtr handle,
            StringBuilder model_name,
            int buffer_size);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr rwkv_create_state(IntPtr model_handle);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern void rwkv_free_state(IntPtr state_handle);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_batch_generate(
            IntPtr model_handle,
            string[] prompts,
            int prompt_count,
            ref RwkvGenerateOptionsNative options,
            [Out] StringBuilder[] outputs,
            [Out] int[] output_sizes);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_generate_with_state(
            IntPtr model_handle,
            IntPtr state_handle,
            [MarshalAs(UnmanagedType.LPStr)] string prompt,
            ref RwkvGenerateOptionsNative options,
            StringBuilder output,
            int output_size);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern int rwkv_count_tokens(
            IntPtr model_handle,
            [MarshalAs(UnmanagedType.LPStr)] string text);

        [DllImport(DLL_NAME, CallingConvention = CallingConvention.Cdecl)]
        private static extern IntPtr rwkv_get_last_error();

        #endregion

        #region 结构体定义

        [StructLayout(LayoutKind.Sequential)]
        private struct RwkvGenerateOptionsNative
        {
            public int max_tokens;
            public double temperature;
            public int top_k;
            public double top_p;
            public double alpha_presence;
            public double alpha_frequency;
            public double alpha_decay;
            public int pad_zero;

            public static RwkvGenerateOptionsNative FromManaged(RwkvGenerateOptions opts)
            {
                return new RwkvGenerateOptionsNative
                {
                    max_tokens = opts.MaxTokens,
                    temperature = opts.Temperature,
                    top_k = opts.TopK,
                    top_p = opts.TopP,
                    alpha_presence = opts.AlphaPresence,
                    alpha_frequency = opts.AlphaFrequency,
                    alpha_decay = opts.AlphaDecay,
                    pad_zero = opts.PadZero ? 1 : 0
                };
            }
        }

        #endregion

        private readonly ILogger _logger;
        private bool _disposed = false;
        private static bool _logCallbackSet = false;
        private static readonly object _lock = new object();

        /// <summary>
        /// 模型是否已加载
        /// </summary>
        public bool IsLoaded => _modelHandle != IntPtr.Zero;

        /// <summary>
        /// 模型名称
        /// </summary>
        public string ModelName { get; private set; } = "";

        /// <summary>
        /// GPU 是否可用
        /// </summary>
        public static bool IsGpuAvailable => rwkv_is_gpu_available() == 1;

        public RwkvNativeEngine(ILogger logger)
        {
            _logger = logger;
            
            // 设置日志回调（只设置一次）
            if (!_logCallbackSet)
            {
                lock (_lock)
                {
                    if (!_logCallbackSet)
                    {
                        try
                        {
                            rwkv_set_log_callback(OnNativeLog);
                            _logCallbackSet = true;
                        }
                        catch (DllNotFoundException ex)
                        {
                            _logger.LogWarning($"rwkv_native.dll 未找到: {ex.Message}");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 日志回调
        /// </summary>
        private void OnNativeLog(string message)
        {
            _logger?.LogInformation($"[RWKV Native] {message}");
        }

        /// <summary>
        /// 获取 GPU 信息
        /// </summary>
        public static string GetGpuInfo()
        {
            try
            {
                var sb = new StringBuilder(1024);
                rwkv_get_gpu_info(sb, sb.Capacity);
                return sb.ToString();
            }
            catch
            {
                return "无法获取 GPU 信息";
            }
        }

        /// <summary>
        /// 加载模型
        /// </summary>
        public bool LoadModel(string modelPath, string vocabPath, int gpuId = 0)
        {
            if (_modelHandle != IntPtr.Zero)
            {
                _logger.LogWarning("模型已加载，请先卸载");
                return false;
            }

            try
            {
                _logger.LogInformation($"正在加载模型: {modelPath}");
                
                _modelHandle = rwkv_load_model(modelPath, vocabPath, gpuId);
                
                if (_modelHandle == IntPtr.Zero)
                {
                    var error = GetLastNativeError();
                    _logger.LogError($"模型加载失败: {error}");
                    return false;
                }

                // 获取模型信息
                var nameSb = new StringBuilder(256);
                rwkv_get_model_info(_modelHandle, nameSb, nameSb.Capacity);
                ModelName = nameSb.ToString();

                // 创建状态
                _stateHandle = rwkv_create_state(_modelHandle);
                
                _logger.LogInformation($"模型加载成功: {ModelName}");
                return true;
            }
            catch (DllNotFoundException ex)
            {
                _logger.LogError($"rwkv_native.dll 未找到: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载模型失败");
                return false;
            }
        }

        /// <summary>
        /// 卸载模型
        /// </summary>
        public void UnloadModel()
        {
            if (_stateHandle != IntPtr.Zero)
            {
                rwkv_free_state(_stateHandle);
                _stateHandle = IntPtr.Zero;
            }

            if (_modelHandle != IntPtr.Zero)
            {
                rwkv_unload_model(_modelHandle);
                _modelHandle = IntPtr.Zero;
                ModelName = "";
                _logger.LogInformation("模型已卸载");
            }
        }

        /// <summary>
        /// 批量翻译
        /// </summary>
        public string[] BatchTranslate(string[] prompts, RwkvGenerateOptions? options = null)
        {
            if (_modelHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("模型未加载");
            }

            var opts = options ?? new RwkvGenerateOptions();
            var nativeOpts = RwkvGenerateOptionsNative.FromManaged(opts);

            var outputs = new StringBuilder[prompts.Length];
            var outputSizes = new int[prompts.Length];
            
            for (int i = 0; i < prompts.Length; i++)
            {
                outputs[i] = new StringBuilder(8192);
                outputSizes[i] = 8192;
            }

            int result = rwkv_batch_generate(
                _modelHandle,
                prompts,
                prompts.Length,
                ref nativeOpts,
                outputs,
                outputSizes
            );

            if (result != RWKV_SUCCESS)
            {
                var error = GetLastNativeError();
                throw new Exception($"推理失败: {error}");
            }

            var results = new string[prompts.Length];
            for (int i = 0; i < prompts.Length; i++)
            {
                results[i] = outputs[i].ToString();
            }

            return results;
        }

        /// <summary>
        /// 单条翻译（带状态，支持连续对话）
        /// </summary>
        public string Translate(string prompt, RwkvGenerateOptions? options = null)
        {
            if (_modelHandle == IntPtr.Zero)
            {
                throw new InvalidOperationException("模型未加载");
            }

            var opts = options ?? new RwkvGenerateOptions();
            var nativeOpts = RwkvGenerateOptionsNative.FromManaged(opts);

            var output = new StringBuilder(8192);

            int result = rwkv_generate_with_state(
                _modelHandle,
                _stateHandle,
                prompt,
                ref nativeOpts,
                output,
                output.Capacity
            );

            if (result != RWKV_SUCCESS)
            {
                var error = GetLastNativeError();
                throw new Exception($"推理失败: {error}");
            }

            return output.ToString();
        }

        /// <summary>
        /// 统计 token 数量
        /// </summary>
        public int CountTokens(string text)
        {
            if (_modelHandle == IntPtr.Zero)
            {
                return -1;
            }

            return rwkv_count_tokens(_modelHandle, text);
        }

        /// <summary>
        /// 重置对话状态
        /// </summary>
        public void ResetState()
        {
            if (_stateHandle != IntPtr.Zero)
            {
                rwkv_free_state(_stateHandle);
                _stateHandle = rwkv_create_state(_modelHandle);
                _logger.LogDebug("对话状态已重置");
            }
        }

        private string GetLastNativeError()
        {
            try
            {
                var errorPtr = rwkv_get_last_error();
                if (errorPtr != IntPtr.Zero)
                {
                    return Marshal.PtrToStringAnsi(errorPtr) ?? "Unknown error";
                }
            }
            catch { }
            return "Unknown error";
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                UnloadModel();
                _disposed = true;
            }
        }
    }

    /// <summary>
    /// 推理生成选项
    /// </summary>
    public class RwkvGenerateOptions
    {
        public int MaxTokens { get; set; } = 512;
        public double Temperature { get; set; } = 1.0;
        public int TopK { get; set; } = 50;
        public double TopP { get; set; } = 0.6;
        public double AlphaPresence { get; set; } = 1.0;
        public double AlphaFrequency { get; set; } = 0.1;
        public double AlphaDecay { get; set; } = 0.996;
        public bool PadZero { get; set; } = true;
    }
}
