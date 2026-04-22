using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
        /// 资源使用情况
        /// </summary>
        public class ResourceUsage
        {
            public double GpuUtilization { get; set; }
            public double GpuMemoryUtilization { get; set; }
            public double CpuUtilization { get; set; }
            public double CpuMemoryUtilization { get; set; }
            public bool UseGpu { get; set; } = true;
            /// <summary>
            /// 峰值显存占用（字节）
            /// </summary>
            public long PeakGpuMemoryBytes { get; set; }
            /// <summary>
            /// 是否检测到共享显存被占用
            /// </summary>
            public bool SharedMemoryUsed { get; set; }
            /// <summary>
            /// 共享显存占用字节数
            /// </summary>
            public long SharedMemoryBytes { get; set; }
            /// <summary>
            /// 是否应中止当前测试批次（共享显存被占用或专属显存>97%）
            /// </summary>
            public bool ShouldAbort { get; set; }
            /// <summary>
            /// 中止原因
            /// </summary>
            public string AbortReason { get; set; } = "";
        }

    /// <summary>
    /// Benchmark 测试结果
    /// </summary>
    public class BenchmarkResult
    {
        public int ShortPromptMaxBatchSize { get; set; }
        public int MediumPromptMaxBatchSize { get; set; }
        public int LongPromptMaxBatchSize { get; set; }
        public double ShortPromptMaxTokPerSec { get; set; }
        public double MediumPromptMaxTokPerSec { get; set; }
        public double LongPromptMaxTokPerSec { get; set; }
        public string ShortPromptName { get; set; } = "短 prompt (~50 tokens)";
        public string MediumPromptName { get; set; } = "中 prompt (~500 tokens)";
        public string LongPromptName { get; set; } = "长 prompt (~2000 tokens)";
    }

    /// <summary>
    /// 并发计算服务
    /// 综合GPU显存和CPU核心数计算最大并发数
    /// </summary>
    public class ConcurrencyCalculator
    {
        private static readonly Encoding OutputEncoding;

        static ConcurrencyCalculator()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            OutputEncoding = Encoding.UTF8;
        }

        private readonly ILogger _logger;
        private readonly GpuResourceService _gpuService;
        private readonly ModelManagementService _modelService;
        private readonly ModelBenchmarkCache _benchmarkCache;

        // 安全预留比例（5%，目标使用95%显存）
        private const double SafetyMargin = 0.05;

        // 不同参数量模型的显存消耗系数（GB/并发）
        private static readonly double[] ModelMemoryCoefficients = { 0.1, 0.2, 0.5, 1.0 };

        /// <summary>
        /// 每批次并发请求上限（GPU资源超警戒时使用）
        /// 默认50，可通过UI设置修改
        /// </summary>
        public static int BatchConcurrencyLimit { get; set; } = 50;

        public ConcurrencyCalculator(ILogger logger, GpuResourceService gpuService, ModelManagementService modelService, ModelBenchmarkCache benchmarkCache)
        {
            _logger = logger;
            _gpuService = gpuService;
            _modelService = modelService;
            _benchmarkCache = benchmarkCache;
        }

        /// <summary>
        /// 计算最大并发数（使用 HTTP 并发请求实际测试）
        /// </summary>
        /// <param name="gpu">目标GPU</param>
        /// <param name="model">要加载的模型</param>
        /// <param name="apiUrl">API 地址</param>
        /// <param name="enableBszTest">是否启用并发上限测试</param>
        /// <param name="bszIncrement">并发测试每次递增的数量</param>
        /// <param name="engineType">引擎类型: "rwkv" 或 "llama_cpp"</param>
        /// <returns>计算得出的最大并发数</returns>
        public async Task<int> CalculateMaxConcurrencyAsync(GpuInfo? gpu, ModelInfo model, string? apiUrl = null, bool enableBszTest = false, int bszIncrement = 10, string engineType = "rwkv")
        {
            try
            {
                // 刷新GPU状态获取最新显存信息（仅GPU模式）
                if (gpu != null)
                {
                    var currentGpu = await _gpuService.RefreshGpuStatusAsync(gpu.Index);
                    if (currentGpu != null)
                    {
                        gpu = currentGpu;
                    }
                }

                // 如果有 API 地址，先检查缓存
                if (!string.IsNullOrEmpty(apiUrl))
                {
                    // 检查是否有缓存的并发数数据
                    var cacheKey = $"{apiUrl}_{engineType}";
                    var gpuIndex = gpu?.Index ?? -1;
                    var cachedConcurrency = _benchmarkCache.GetMaxBatchSize(cacheKey, gpuIndex);
                    if (cachedConcurrency.HasValue)
                    {
                        _logger.LogInformation($"使用缓存的并发数数据: {cachedConcurrency.Value}");
                        return cachedConcurrency.Value;
                    }

                    // 只有启用测试时才进行实际测试
                    if (enableBszTest)
                    {
                        _logger.LogInformation($"开始 HTTP 并发上限测试（引擎: {engineType}），增量值: {bszIncrement}");
                        
                        BenchmarkResult benchmarkResult;
                        if (engineType.StartsWith("llama-"))
                        {
                            // llama_cpp 各工具各自适配测试方案
                            benchmarkResult = await TestAllScenariosLlamaCppAsync(apiUrl, gpu, bszIncrement, engineType);
                        }
                        else
                        {
                            benchmarkResult = await TestAllScenariosAsync(apiUrl, gpu!, bszIncrement);
                        }
                        
                        if (benchmarkResult.MediumPromptMaxBatchSize > 0)
                        {
                            // 保存测试结果到缓存
                            _benchmarkCache.SaveMaxBatchSize(
                                cacheKey, gpuIndex,
                                benchmarkResult.ShortPromptMaxBatchSize,
                                benchmarkResult.MediumPromptMaxBatchSize,
                                benchmarkResult.LongPromptMaxBatchSize);
                            
                            _logger.LogInformation(
                                $"HTTP 并发测试完成: " +
                                $"短prompt 并发={benchmarkResult.ShortPromptMaxBatchSize}, " +
                                $"中prompt 并发={benchmarkResult.MediumPromptMaxBatchSize}, " +
                                $"长prompt 并发={benchmarkResult.LongPromptMaxBatchSize}");
                            
                            // 返回中 prompt 场景的并发数作为默认值
                            return benchmarkResult.MediumPromptMaxBatchSize;
                        }
                    }
                    else
                    {
                        _logger.LogInformation("未启用并发测试，使用显存估算");
                    }
                }

                // 回退到基于资源的估算
                int cpuCores = Environment.ProcessorCount;
                int cpuBasedLimit = (int)Math.Ceiling(cpuCores * 0.7);
                
                if (gpu != null)
                {
                    int gpuBasedConcurrency = CalculateGpuBasedConcurrency(gpu, model);
                    int maxConcurrency = Math.Min(gpuBasedConcurrency, cpuBasedLimit);
                    maxConcurrency = Math.Max(1, maxConcurrency);

                    _logger.LogInformation(
                        $"并发计算完成(估算): GPU显存={gpu.UsedMemoryGB:F1}/{gpu.DedicatedMemoryGB:F1}GB, " +
                        $"模型估算={model.EstimatedMemoryGB:F1}GB, " +
                        $"GPU并发上限={gpuBasedConcurrency}, " +
                        $"CPU并发上限={cpuBasedLimit}(70%), " +
                        $"最终并发数={maxConcurrency}");

                    return maxConcurrency;
                }
                else
                {
                    // CPU 模式：仅基于 CPU 核心数估算
                    int maxConcurrency = Math.Max(1, cpuBasedLimit);
                    _logger.LogInformation(
                        $"并发计算完成(CPU估算): CPU核心数={cpuCores}, " +
                        $"CPU并发上限={cpuBasedLimit}(70%), " +
                        $"最终并发数={maxConcurrency}");
                    return maxConcurrency;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "计算并发数时发生异常");
                return 4; // 返回默认值
            }
        }

        /// <summary>
        /// 测试三种场景的最大 HTTP 并发请求数
        /// </summary>
        private async Task<BenchmarkResult> TestAllScenariosAsync(string apiUrl, GpuInfo gpu, int concurrencyIncrement)
        {
            var result = new BenchmarkResult();

            // 短 prompt 场景（约 50 tokens）
            _logger.LogInformation("开始测试短 prompt 场景的 HTTP 并发请求数...");
            var (shortConc, shortToks) = await TestHttpConcurrencyWithPromptAsync(
                apiUrl, gpu, "你好", concurrencyIncrement);
            result.ShortPromptMaxBatchSize = shortConc;
            result.ShortPromptMaxTokPerSec = shortToks;

            // 中 prompt 场景（约 500 tokens）
            _logger.LogInformation("开始测试中 prompt 场景的 HTTP 并发请求数...");
            var (medConc, medToks) = await TestHttpConcurrencyWithPromptAsync(
                apiUrl, gpu,
                "请根据以下内容进行翻译：The quick brown fox jumps over the lazy dog. " +
                "This is a classic pangram that contains every letter of the English alphabet. " +
                "It has been used for decades to test typewriters, keyboards, and fonts. " +
                "The sentence is short enough to be memorable but long enough to be useful. " +
                "In addition, it demonstrates the concept of efficient information transfer. " +
                "This makes it an ideal test case for translation systems. " +
                "Now please translate this text into Chinese.", concurrencyIncrement);
            result.MediumPromptMaxBatchSize = medConc;
            result.MediumPromptMaxTokPerSec = medToks;

            // 长 prompt 场景（约 2000 tokens）
            _logger.LogInformation("开始测试长 prompt 场景的 HTTP 并发请求数...");
            var (longConc, longToks) = await TestHttpConcurrencyWithPromptAsync(
                apiUrl, gpu,
                "请根据以下长文档内容进行翻译：\n\n" +
                "Artificial intelligence (AI) is intelligence demonstrated by machines, as opposed to natural intelligence displayed by animals and humans. " +
                "AI research has been defined as the field of study of intelligent agents, which refers to any system that perceives its environment and takes actions that maximize its chance of achieving its goals. " +
                "The term \"artificial intelligence\" had previously been used to describe machines that mimic and display \"human\" cognitive skills that are associated with the human mind, such as \"learning\" and \"problem-solving\". " +
                "This definition has since been rejected by major AI researchers, who now describe AI in terms of rationality and acting rationally, which does not limit how intelligence can be articulated.\n\n" +
                "AI applications include advanced web search engines, recommendation systems, understanding human speech, self-driving cars, automated decision-making, and competing at the highest level in strategic game systems. " +
                "As machines become increasingly capable, tasks considered to require \"intelligence\" are often removed from the definition of AI, a phenomenon known as the AI effect. " +
                "For instance, optical character recognition is frequently excluded from things considered to be AI, having become a routine technology.\n\n" +
                "Artificial intelligence was founded as an academic discipline in 1956, and in the years since it has experienced several waves of optimism, followed by disappointment and the loss of funding, followed by new approaches, success, and renewed funding. " +
                "AI research has tried and discarded many different approaches, including simulating the brain, modeling human problem solving, formal logic, large databases of knowledge, and imitating animal behavior. " +
                "In the first decades of the 21st century, highly mathematical and statistical machine learning has dominated the field, and this technique has proved highly successful, helping to solve many challenging problems throughout industry and academia.\n\n" +
                "The various sub-fields of AI research are centered around particular goals and the use of particular tools. " +
                "The traditional goals of AI research include reasoning, knowledge representation, planning, learning, natural language processing, perception, and the ability to move and manipulate objects. " +
                "General intelligence (the ability to solve an arbitrary problem) is among the field's long-term goals. " +
                "To solve these problems, AI researchers have adapted and integrated a wide range of problem-solving techniques, including search and mathematical optimization, formal logic, artificial neural networks, and methods based on statistics, probability, and economics. " +
                "AI also draws upon computer science, psychology, linguistics, philosophy, and many other fields.\n\n" +
                "Now please translate this entire document into Chinese, maintaining the original meaning and tone.", concurrencyIncrement);
            result.LongPromptMaxBatchSize = longConc;
            result.LongPromptMaxTokPerSec = longToks;

            // 打印汇总报告
            _logger.LogInformation("");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation("📊 HTTP 并发测试汇总报告");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation($"  短 prompt 场景: 最大并发={result.ShortPromptMaxBatchSize}, 吞吐≈{result.ShortPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation($"  中 prompt 场景: 最大并发={result.MediumPromptMaxBatchSize}, 吞吐≈{result.MediumPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation($"  长 prompt 场景: 最大并发={result.LongPromptMaxBatchSize}, 吞吐≈{result.LongPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation("");

            return result;
        }

        /// <summary>
        /// 测试 llama_cpp 三种场景的最大并发请求数
        /// llama_cpp 使用 /v1/completions API，请求格式与 rwkv_lightning 不同
        /// 各推理工具（cuda/sycl/vulkan/cpu）使用差异化的测试参数
        /// </summary>
        private async Task<BenchmarkResult> TestAllScenariosLlamaCppAsync(string apiUrl, GpuInfo? gpu, int concurrencyIncrement, string toolName)
        {
            var result = new BenchmarkResult();
            
            // 根据推理工具确定差异化参数
            var toolConfig = GetLlamaCppToolConfig(toolName, concurrencyIncrement);
            var effectiveIncrement = toolConfig.ConcurrencyIncrement;
            var isCpuMode = toolConfig.IsCpuMode;
            
            _logger.LogInformation($"llama_cpp 工具 [{toolName}] 并发测试配置: " +
                $"增量={effectiveIncrement}, 模式={(isCpuMode ? "CPU" : "GPU")}, " +
                $"初始并发={toolConfig.InitialConcurrency}, 最大并发={toolConfig.MaxConcurrencyLimit}");

            // 短 prompt 场景（约 50 tokens）- 使用 RWKV 官方续写翻译格式
            _logger.LogInformation($"开始测试 [{toolName}] 短 prompt 场景的 HTTP 并发请求数...");
            var (shortConc, shortToks) = await TestLlamaCppConcurrencyWithPromptAsync(
                apiUrl, gpu, "Chinese:\n你好\n\nEnglish:", effectiveIncrement, isCpuMode, toolConfig.InitialConcurrency, toolConfig.MaxConcurrencyLimit);
            result.ShortPromptMaxBatchSize = shortConc;
            result.ShortPromptMaxTokPerSec = shortToks;

            // 中 prompt 场景（约 500 tokens）
            _logger.LogInformation($"开始测试 [{toolName}] 中 prompt 场景的 HTTP 并发请求数...");
            var (medConc, medToks) = await TestLlamaCppConcurrencyWithPromptAsync(
                apiUrl, gpu,
                "Chinese:\n请根据以下内容进行翻译：The quick brown fox jumps over the lazy dog. " +
                "This is a classic pangram that contains every letter of the English alphabet. " +
                "It has been used for decades to test typewriters, keyboards, and fonts. " +
                "The sentence is short enough to be memorable but long enough to be useful. " +
                "In addition, it demonstrates the concept of efficient information transfer. " +
                "This makes it an ideal test case for translation systems. " +
                "Now please translate this text into Chinese.\n\nEnglish:",
                effectiveIncrement, isCpuMode, toolConfig.InitialConcurrency, toolConfig.MaxConcurrencyLimit);
            result.MediumPromptMaxBatchSize = medConc;
            result.MediumPromptMaxTokPerSec = medToks;

            // 长 prompt 场景（约 2000 tokens）
            _logger.LogInformation($"开始测试 [{toolName}] 长 prompt 场景的 HTTP 并发请求数...");
            var (longConc, longToks) = await TestLlamaCppConcurrencyWithPromptAsync(
                apiUrl, gpu,
                "Chinese:\n请根据以下长文档内容进行翻译：\n\n" +
                "Artificial intelligence (AI) is intelligence demonstrated by machines, as opposed to natural intelligence displayed by animals and humans. " +
                "AI research has been defined as the field of study of intelligent agents, which refers to any system that perceives its environment and takes actions that maximize its chance of achieving its goals. " +
                "The term \"artificial intelligence\" had previously been used to describe machines that mimic and display \"human\" cognitive skills that are associated with the human mind, such as \"learning\" and \"problem-solving\". " +
                "This definition has since been rejected by major AI researchers, who now describe AI in terms of rationality and acting rationally, which does not limit how intelligence can be articulated.\n\n" +
                "AI applications include advanced web search engines, recommendation systems, understanding human speech, self-driving cars, automated decision-making, and competing at the highest level in strategic game systems. " +
                "As machines become increasingly capable, tasks considered to require \"intelligence\" are often removed from the definition of AI, a phenomenon known as the AI effect. " +
                "For instance, optical character recognition is frequently excluded from things considered to be AI, having become a routine technology.\n\n" +
                "Now please translate this entire document into English, maintaining the original meaning and tone.\n\nEnglish:",
                effectiveIncrement, isCpuMode, toolConfig.InitialConcurrency, toolConfig.MaxConcurrencyLimit);
            result.LongPromptMaxBatchSize = longConc;
            result.LongPromptMaxTokPerSec = longToks;

            // 打印汇总报告
            _logger.LogInformation("");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation($"📊 [{toolName}] HTTP 并发测试汇总报告");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation($"  短 prompt 场景: 最大并发={result.ShortPromptMaxBatchSize}, 吞吐≈{result.ShortPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation($"  中 prompt 场景: 最大并发={result.MediumPromptMaxBatchSize}, 吞吐≈{result.MediumPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation($"  长 prompt 场景: 最大并发={result.LongPromptMaxBatchSize}, 吞吐≈{result.LongPromptMaxTokPerSec:F0} tok/s");
            _logger.LogInformation("═══════════════════════════════════════════");
            _logger.LogInformation("");

            return result;
        }

        /// <summary>
        /// llama_cpp 各推理工具的并发测试配置
        /// </summary>
        private class LlamaCppToolTestConfig
        {
            public int ConcurrencyIncrement { get; set; } = 10;
            public int InitialConcurrency { get; set; } = 1;
            public int MaxConcurrencyLimit { get; set; } = 1000;
            public bool IsCpuMode { get; set; } = false;
        }

        /// <summary>
        /// 根据推理工具名获取差异化测试配置
        /// llama-cuda:  GPU加速，并发增量较大
        /// llama-sycl:  Intel GPU加速，并发增量适中
        /// llama-vulkan: 通用GPU加速，并发增量适中
        /// llama-cpu:   纯CPU推理，并发增量较小，上限较低
        /// </summary>
        private static LlamaCppToolTestConfig GetLlamaCppToolConfig(string toolName, int defaultIncrement)
        {
            return toolName switch
            {
                "llama-cuda" => new LlamaCppToolTestConfig
                {
                    ConcurrencyIncrement = Math.Max(defaultIncrement, 10),
                    InitialConcurrency = 1,
                    MaxConcurrencyLimit = 1000,
                    IsCpuMode = false
                },
                "llama-sycl" => new LlamaCppToolTestConfig
                {
                    ConcurrencyIncrement = Math.Max(defaultIncrement, 5),
                    InitialConcurrency = 1,
                    MaxConcurrencyLimit = 500,
                    IsCpuMode = false
                },
                "llama-vulkan" => new LlamaCppToolTestConfig
                {
                    ConcurrencyIncrement = Math.Max(defaultIncrement, 5),
                    InitialConcurrency = 1,
                    MaxConcurrencyLimit = 500,
                    IsCpuMode = false
                },
                "llama-cpu" => new LlamaCppToolTestConfig
                {
                    ConcurrencyIncrement = Math.Max(Math.Min(defaultIncrement, 2), 1),
                    InitialConcurrency = 1,
                    MaxConcurrencyLimit = Math.Min(Environment.ProcessorCount, 100),
                    IsCpuMode = true
                },
                _ => new LlamaCppToolTestConfig
                {
                    ConcurrencyIncrement = defaultIncrement,
                    InitialConcurrency = 1,
                    MaxConcurrencyLimit = 1000,
                    IsCpuMode = false
                }
            };
        }

        /// <summary>
        /// 使用指定 prompt 测试 llama_cpp 最大 HTTP 并发请求数
        /// llama_cpp 使用 /v1/completions API，每个请求独立发送
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task<(int maxConcurrency, double maxTokensPerSec)> TestLlamaCppConcurrencyWithPromptAsync(string apiUrl, GpuInfo? gpu, string promptText, int concurrencyIncrement, bool isCpuMode = false, int initialConcurrency = 1, int maxConcurrencyLimit = 1000)
        {
            int concurrency = initialConcurrency;
            int maxSuccessfulConcurrency = initialConcurrency;
            double maxTokensPerSec = 0;

            while (true)
            {
                // 运行 llama_cpp 并发测试，并在测试期间监控峰值资源
                var (success, peakResourceCheck) = await RunLlamaCppConcurrencyTestWithMonitoringAsync(apiUrl, gpu, concurrency, promptText, isCpuMode);
                
                // 同时运行计时版本获取 tok/s
                var (success2, elapsedMs, tokensPerSec) = await RunLlamaCppConcurrencyTestWithTimingAsync(apiUrl, concurrency, promptText);
                
                var resourceCheck = peakResourceCheck;
                
                // 检查是否需要中止测试
                if (resourceCheck.ShouldAbort)
                {
                    _logger.LogWarning($"[llama_cpp并发测试] 检测到需要中止: {resourceCheck.AbortReason}，停止测试，最大并发数={maxSuccessfulConcurrency}");
                    break;
                }
                
                // 判断使用模式
                bool reachedLimit;
                string limitReason;
                
                if (resourceCheck.UseGpu && !isCpuMode)
                {
                    reachedLimit = resourceCheck.GpuUtilization >= 95.0 && resourceCheck.GpuMemoryUtilization >= 95.0;
                    if (reachedLimit)
                        limitReason = $"GPU 使用率={resourceCheck.GpuUtilization:F1}% (≥95%) 且 专属显存使用率={resourceCheck.GpuMemoryUtilization:F1}% (≥95%)";
                    else
                        limitReason = $"GPU 使用率={resourceCheck.GpuUtilization:F1}%, 专属显存使用率={resourceCheck.GpuMemoryUtilization:F1}% (未同时达到95%上限)";
                }
                else
                {
                    reachedLimit = resourceCheck.CpuUtilization >= 95.0 && resourceCheck.CpuMemoryUtilization >= 90.0;
                    if (reachedLimit)
                        limitReason = $"CPU 使用率={resourceCheck.CpuUtilization:F1}% (≥95%) 且 内存使用率={resourceCheck.CpuMemoryUtilization:F1}% (≥90%)";
                    else
                        limitReason = $"CPU 使用率={resourceCheck.CpuUtilization:F1}%, 内存使用率={resourceCheck.CpuMemoryUtilization:F1}% (未同时达到上限)";
                }

                if (resourceCheck.UseGpu && !isCpuMode && gpu != null)
                {
                    var peakUsedMemGB = resourceCheck.PeakGpuMemoryBytes / (1024.0 * 1024.0 * 1024.0);
                    var dedicatedMemGB = gpu.DedicatedMemoryGB;
                    _logger.LogInformation(
                        $"[llama_cpp并发测试] 并发数={concurrency}, GPU 使用率={resourceCheck.GpuUtilization:F1}%, " +
                        $"GPU 显存峰值={peakUsedMemGB:F1}/{dedicatedMemGB:F1}GB ({resourceCheck.GpuMemoryUtilization:F1}%), " +
                        $"CPU 使用率={resourceCheck.CpuUtilization:F1}%, 内存使用率={resourceCheck.CpuMemoryUtilization:F1}%");
                }
                else
                {
                    _logger.LogInformation(
                        $"[llama_cpp并发测试] 并发数={concurrency}, CPU 使用率={resourceCheck.CpuUtilization:F1}%, " +
                        $"内存使用率={resourceCheck.CpuMemoryUtilization:F1}% ({(isCpuMode ? "CPU模式" : "GPU模式")})");
                }
                
                if (reachedLimit)
                {
                    _logger.LogInformation(
                        $"资源使用率已达上限: {limitReason}, " +
                        $"停止测试，最大并发数={maxSuccessfulConcurrency}");
                    break;
                }
                
                if (success && success2)
                {
                    maxSuccessfulConcurrency = concurrency;
                    if (tokensPerSec > maxTokensPerSec)
                        maxTokensPerSec = tokensPerSec;
                    _logger.LogInformation($"llama_cpp 并发数 {concurrency} 测试成功");
                    
                    concurrency += concurrencyIncrement;
                }
                else
                {
                    _logger.LogInformation($"llama_cpp 并发数 {concurrency} 测试失败，停止测试");
                    break;
                }

                if (concurrency > maxConcurrencyLimit)
                {
                    _logger.LogWarning($"并发数超过安全限制 ({maxConcurrencyLimit})，停止测试");
                    break;
                }
            }

            return (maxSuccessfulConcurrency, maxTokensPerSec);
        }

        /// <summary>
        /// 运行 llama_cpp HTTP 并发请求测试，并在测试期间持续监控 GPU 显存峰值
        /// llama_cpp 使用 /v1/completions API，并发发送多个独立请求
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task<(bool success, ResourceUsage peakUsage)> RunLlamaCppConcurrencyTestWithMonitoringAsync(string apiUrl, GpuInfo? gpu, int concurrency, string promptText, bool isCpuMode)
        {
            var peakUsage = new ResourceUsage { UseGpu = !isCpuMode, PeakGpuMemoryBytes = 0 };
            var monitorCts = new System.Threading.CancellationTokenSource();
            
            // 启动后台监控任务
            var monitorTask = Task.Run(async () =>
            {
                while (!monitorCts.Token.IsCancellationRequested)
                {
                    try
                    {
                        // GPU 模式下监控 GPU 资源
                        if (!isCpuMode && gpu != null)
                        {
                            var gpuInfo = await _gpuService.RefreshGpuStatusAsync(gpu.Index);
                            if (gpuInfo != null)
                            {
                                lock (peakUsage)
                                {
                                    if (gpuInfo.UsedMemoryBytes > peakUsage.PeakGpuMemoryBytes)
                                        peakUsage.PeakGpuMemoryBytes = gpuInfo.UsedMemoryBytes;
                                    if (gpuInfo.GpuUtilization > peakUsage.GpuUtilization)
                                        peakUsage.GpuUtilization = gpuInfo.GpuUtilization;
                                    if (gpuInfo.DedicatedMemoryBytes > 0)
                                    {
                                        var currentMemUtil = (double)gpuInfo.UsedMemoryBytes / gpuInfo.DedicatedMemoryBytes * 100.0;
                                        if (currentMemUtil > peakUsage.GpuMemoryUtilization)
                                            peakUsage.GpuMemoryUtilization = currentMemUtil;
                                        
                                        var sharedMemUsed = gpuInfo.UsedMemoryBytes > gpuInfo.DedicatedMemoryBytes;
                                        if (sharedMemUsed)
                                        {
                                            peakUsage.SharedMemoryUsed = true;
                                            peakUsage.SharedMemoryBytes = gpuInfo.UsedMemoryBytes - gpuInfo.DedicatedMemoryBytes;
                                        }
                                        if (currentMemUtil > 97.0 && !peakUsage.ShouldAbort)
                                        {
                                            peakUsage.ShouldAbort = true;
                                            peakUsage.AbortReason = $"专属显存占用率 {currentMemUtil:F1}% > 97%";
                                        }
                                        if (sharedMemUsed && !peakUsage.ShouldAbort)
                                        {
                                            peakUsage.ShouldAbort = true;
                                            peakUsage.AbortReason = $"共享显存被占用 ({peakUsage.SharedMemoryBytes / (1024.0*1024.0*1024.0):F1}GB)";
                                        }
                                    }
                                }
                            }
                        }
                        
                        // 始终监控 CPU 资源
                        try
                        {
                            using var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                            cpuCounter.NextValue();
                            await Task.Delay(300, monitorCts.Token).ConfigureAwait(false);
                            var cpuUtil = cpuCounter.NextValue();
                            lock (peakUsage)
                            {
                                if (cpuUtil > peakUsage.CpuUtilization)
                                    peakUsage.CpuUtilization = cpuUtil;
                            }
                        }
                        catch { }
                        
                        try
                        {
                            using var memCounter = new PerformanceCounter("Memory", "% Committed Bytes In Use");
                            var memUtil = memCounter.NextValue();
                            lock (peakUsage)
                            {
                                if (memUtil > peakUsage.CpuMemoryUtilization)
                                    peakUsage.CpuMemoryUtilization = memUtil;
                            }
                        }
                        catch { }
                    }
                    catch { }
                    await Task.Delay(200, monitorCts.Token).ConfigureAwait(false);
                }
            }, monitorCts.Token);

            try
            {
                var success = await RunLlamaCppConcurrencyTestAsync(apiUrl, concurrency, promptText);
                return (success, peakUsage);
            }
            finally
            {
                monitorCts.Cancel();
                try { await monitorTask; } catch { }
                monitorCts.Dispose();
            }
        }

        /// <summary>
        /// 运行 llama_cpp 并发请求测试（带计时）
        /// 使用 /v1/completions API，并发发送多个独立请求
        /// </summary>
        private async Task<(bool success, double elapsedMs, double tokensPerSec)> RunLlamaCppConcurrencyTestWithTimingAsync(string apiUrl, int concurrency, string promptText)
        {
            try
            {
                // 确定 completions API URL
                var baseUrl = apiUrl.TrimEnd('/');
                if (!baseUrl.Contains("/v1/completions"))
                {
                    var uri = new Uri(baseUrl);
                    baseUrl = $"{uri.Scheme}://{uri.Authority}/v1/completions";
                }

                var requestData = new
                {
                    prompt = promptText,
                    max_tokens = 10,
                    temperature = 1.0,
                    top_k = 1,
                    stream = false
                };

                var jsonContent = JsonSerializer.Serialize(requestData);

                using var httpClient = new HttpClient
                {
                    Timeout = TimeSpan.FromSeconds(120)
                };

                // 并发发送多个独立请求
                var tasks = new List<Task<HttpResponseMessage>>();
                for (int i = 0; i < concurrency; i++)
                {
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                    tasks.Add(httpClient.PostAsync(baseUrl, content));
                }

                var sw = System.Diagnostics.Stopwatch.StartNew();
                var responses = await Task.WhenAll(tasks);
                sw.Stop();

                var elapsedMs = sw.Elapsed.TotalMilliseconds;
                bool allSuccess = responses.All(r => r.IsSuccessStatusCode);

                if (!allSuccess)
                {
                    var failedCount = responses.Count(r => !r.IsSuccessStatusCode);
                    _logger.LogWarning($"llama_cpp 并发测试: {failedCount}/{concurrency} 个请求失败");
                    return (false, elapsedMs, 0);
                }

                // 估算 tokens
                int promptTokens = promptText.Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
                int totalOutputTokens = concurrency * (promptTokens + 10); // prompt + max_tokens
                double tokensPerSec = elapsedMs > 0 ? (totalOutputTokens / elapsedMs * 1000.0) : 0;

                return (true, elapsedMs, tokensPerSec);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"llama_cpp 并发测试失败，并发数: {concurrency}");
                return (false, 0, 0);
            }
        }

        /// <summary>
        /// 运行 llama_cpp 并发请求测试
        /// 使用 /v1/completions API，并发发送多个独立请求
        /// </summary>
        private async Task<bool> RunLlamaCppConcurrencyTestAsync(string apiUrl, int concurrency, string promptText)
        {
            try
            {
                var baseUrl = apiUrl.TrimEnd('/');
                if (!baseUrl.Contains("/v1/completions"))
                {
                    var uri = new Uri(baseUrl);
                    baseUrl = $"{uri.Scheme}://{uri.Authority}/v1/completions";
                }

                var requestData = new
                {
                    prompt = promptText,
                    max_tokens = 10,
                    temperature = 1.0,
                    top_k = 1,
                    stream = false
                };

                var jsonContent = JsonSerializer.Serialize(requestData);

                // 分批发送
                var effectiveBatchSize = BatchConcurrencyLimit > 0 ? Math.Min(BatchConcurrencyLimit, concurrency) : concurrency;
                var remaining = concurrency;
                var batchIndex = 0;

                while (remaining > 0)
                {
                    batchIndex++;
                    var currentBatch = Math.Min(remaining, effectiveBatchSize);

                    using var httpClient = new HttpClient
                    {
                        Timeout = TimeSpan.FromSeconds(120)
                    };

                    var tasks = new List<Task<HttpResponseMessage>>();
                    for (int i = 0; i < currentBatch; i++)
                    {
                        var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                        tasks.Add(httpClient.PostAsync(baseUrl, content));
                    }

                    var responses = await Task.WhenAll(tasks);
                    var failedInBatch = responses.Count(r => !r.IsSuccessStatusCode);

                    if (failedInBatch > 0)
                    {
                        _logger.LogWarning($"llama_cpp 并发测试: 批次{batchIndex}中 {failedInBatch}/{currentBatch} 个请求失败");
                        return false;
                    }

                    remaining -= currentBatch;
                    if (effectiveBatchSize < concurrency && remaining > 0)
                    {
                        _logger.LogDebug($"llama_cpp 批次{batchIndex}完成（{currentBatch}个请求），剩余{remaining}个");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"llama_cpp 并发测试失败，并发数: {concurrency}");
                return false;
            }
        }

        /// <summary>
        /// 使用指定 prompt 测试最大 HTTP 并发请求数
        /// 持续增加并发请求数直到资源使用率达到上限
        /// GPU 模式：GPU 使用率 >= 99% 且 GPU 显存使用率 >= 97%
        /// CPU 模式：CPU 使用率 >= 95% 且 内存使用率 >= 90%
        /// 返回 (最大并发数, 峰值 tok/s)
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task<(int maxConcurrency, double maxTokensPerSec)> TestHttpConcurrencyWithPromptAsync(string apiUrl, GpuInfo gpu, string promptText, int concurrencyIncrement)
        {
            int concurrency = 1;
            int maxSuccessfulConcurrency = 1;
            double maxTokensPerSec = 0;

            while (true)
            {
                // 运行 HTTP 并发测试，并在测试期间监控峰值显存
                var (success, peakResourceCheck) = await RunHttpConcurrencyTestWithMonitoringAsync(apiUrl, gpu.Index, concurrency, promptText, gpu);
                
                // 同时运行计时版本获取 tok/s
                var (success2, elapsedMs, tokensPerSec) = await RunHttpConcurrencyTestWithTimingAsync(apiUrl, gpu.Index, concurrency, promptText);
                
                var resourceCheck = peakResourceCheck;
                
                // 检查是否需要中止测试（共享显存被占用或专属显存>97%）
                if (resourceCheck.ShouldAbort)
                {
                    _logger.LogWarning($"[并发测试] 检测到需要中止: {resourceCheck.AbortReason}，停止测试，最大并发数={maxSuccessfulConcurrency}");
                    break;
                }
                
                // 判断使用模式：GPU 和 CPU 二选一，默认 GPU
                bool reachedLimit;
                string limitReason;
                
                if (resourceCheck.UseGpu)
                {
                    // 新规则：GPU占用率>=95% 且 专属显存占用率>=95% 时认为达到资源上限
                    reachedLimit = resourceCheck.GpuUtilization >= 95.0 && resourceCheck.GpuMemoryUtilization >= 95.0;
                    if (reachedLimit)
                        limitReason = $"GPU 使用率={resourceCheck.GpuUtilization:F1}% (≥95%) 且 专属显存使用率={resourceCheck.GpuMemoryUtilization:F1}% (≥95%)";
                    else
                        limitReason = $"GPU 使用率={resourceCheck.GpuUtilization:F1}%, 专属显存使用率={resourceCheck.GpuMemoryUtilization:F1}% (未同时达到95%上限)";
                }
                else
                {
                    reachedLimit = resourceCheck.CpuUtilization >= 95.0 && resourceCheck.CpuMemoryUtilization >= 90.0;
                    if (reachedLimit)
                        limitReason = $"CPU 使用率={resourceCheck.CpuUtilization:F1}% (≥95%) 且 内存使用率={resourceCheck.CpuMemoryUtilization:F1}% (≥90%)";
                    else
                        limitReason = $"CPU 使用率={resourceCheck.CpuUtilization:F1}%, 内存使用率={resourceCheck.CpuMemoryUtilization:F1}% (未同时达到上限)";
                }

                if (resourceCheck.UseGpu)
                {
                    var peakUsedMemGB = resourceCheck.PeakGpuMemoryBytes / (1024.0 * 1024.0 * 1024.0);
                    var dedicatedMemGB = gpu.DedicatedMemoryGB;
                    _logger.LogInformation(
                        $"[并发测试] 并发数={concurrency}, GPU 使用率={resourceCheck.GpuUtilization:F1}%, " +
                        $"GPU 显存峰值={peakUsedMemGB:F1}/{dedicatedMemGB:F1}GB ({resourceCheck.GpuMemoryUtilization:F1}%), " +
                        $"CPU 使用率={resourceCheck.CpuUtilization:F1}%, 内存使用率={resourceCheck.CpuMemoryUtilization:F1}%");
                }
                else
                {
                    _logger.LogInformation(
                        $"[并发测试] 并发数={concurrency}, CPU 使用率={resourceCheck.CpuUtilization:F1}%, " +
                        $"内存使用率={resourceCheck.CpuMemoryUtilization:F1}% (CPU 模式)");
                }
                
                if (reachedLimit)
                {
                    _logger.LogInformation(
                        $"资源使用率已达上限: {limitReason}, " +
                        $"停止测试，最大并发数={maxSuccessfulConcurrency}");
                    break;
                }
                
                if (success && success2)
                {
                    maxSuccessfulConcurrency = concurrency;
                    if (tokensPerSec > maxTokensPerSec)
                        maxTokensPerSec = tokensPerSec;
                    _logger.LogInformation($"并发数 {concurrency} 测试成功");
                    
                    concurrency += concurrencyIncrement;
                }
                else
                {
                    _logger.LogInformation($"并发数 {concurrency} 测试失败，停止测试");
                    break;
                }

                if (concurrency > 1000)
                {
                    _logger.LogWarning($"并发数超过安全限制 (1000)，停止测试");
                    break;
                }
            }

            return (maxSuccessfulConcurrency, maxTokensPerSec);
        }

        /// <summary>
        /// 运行 HTTP 并发请求测试，并在测试期间持续监控 GPU 显存峰值
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task<(bool success, ResourceUsage peakUsage)> RunHttpConcurrencyTestWithMonitoringAsync(string apiUrl, int gpuIndex, int concurrency, string promptText, GpuInfo gpu)
        {
            var peakUsage = new ResourceUsage { UseGpu = true, PeakGpuMemoryBytes = 0 };
            var monitorCts = new System.Threading.CancellationTokenSource();
            
            // 启动后台监控任务，持续采样显存峰值，并检测共享显存占用和专属显存>97%
            var monitorTask = Task.Run(async () =>
            {
                while (!monitorCts.Token.IsCancellationRequested)
                {
                    try
                    {
                        var gpuInfo = await _gpuService.RefreshGpuStatusAsync(gpuIndex);
                        if (gpuInfo != null)
                        {
                            lock (peakUsage)
                            {
                                if (gpuInfo.UsedMemoryBytes > peakUsage.PeakGpuMemoryBytes)
                                {
                                    peakUsage.PeakGpuMemoryBytes = gpuInfo.UsedMemoryBytes;
                                }
                                // 记录峰值 GPU 使用率
                                if (gpuInfo.GpuUtilization > peakUsage.GpuUtilization)
                                {
                                    peakUsage.GpuUtilization = gpuInfo.GpuUtilization;
                                }
                                // 计算峰值显存使用率（基于专属显存）
                                if (gpuInfo.DedicatedMemoryBytes > 0)
                                {
                                    var currentMemUtil = (double)gpuInfo.UsedMemoryBytes / gpuInfo.DedicatedMemoryBytes * 100.0;
                                    if (currentMemUtil > peakUsage.GpuMemoryUtilization)
                                    {
                                        peakUsage.GpuMemoryUtilization = currentMemUtil;
                                    }
                                    
                                    // 检测共享显存是否被占用
                                    // TotalMemoryBytes包含共享显存，DedicatedMemoryBytes仅含专属显存
                                    // 如果已用显存超过专属显存，说明使用了共享显存
                                    var sharedMemUsed = gpuInfo.UsedMemoryBytes > gpuInfo.DedicatedMemoryBytes;
                                    if (sharedMemUsed)
                                    {
                                        peakUsage.SharedMemoryUsed = true;
                                        peakUsage.SharedMemoryBytes = gpuInfo.UsedMemoryBytes - gpuInfo.DedicatedMemoryBytes;
                                    }
                                    
                                    // 检测专属显存占用率>97%，需要中止测试
                                    if (currentMemUtil > 97.0 && !peakUsage.ShouldAbort)
                                    {
                                        peakUsage.ShouldAbort = true;
                                        peakUsage.AbortReason = $"专属显存占用率 {currentMemUtil:F1}% > 97%";
                                    }
                                    
                                    // 检测共享显存被占用，需要中止测试
                                    if (sharedMemUsed && !peakUsage.ShouldAbort)
                                    {
                                        peakUsage.ShouldAbort = true;
                                        peakUsage.AbortReason = $"共享显存被占用 ({peakUsage.SharedMemoryBytes / (1024.0*1024.0*1024.0):F1}GB)";
                                    }
                                }
                            }
                        }
                    }
                    catch { }
                    
                    await Task.Delay(200, monitorCts.Token).ConfigureAwait(false);
                }
            }, monitorCts.Token);

            try
            {
                // 根据当前GPU资源状态决定批次大小
                // GPU占用率>=95% 且 专属显存占用率>=95% 时，使用用户设定的每批次并发上限
                // 否则不限制批次大小（一次性发完）
                int maxBatchSize = 0; // 0表示不限制
                lock (peakUsage)
                {
                    if (peakUsage.GpuUtilization >= 95.0 && peakUsage.GpuMemoryUtilization >= 95.0)
                    {
                        maxBatchSize = BatchConcurrencyLimit > 0 ? BatchConcurrencyLimit : 50;
                    }
                }
                
                // 执行 HTTP 请求（支持分批发送）
                var success = await RunHttpConcurrencyTestAsync(apiUrl, gpuIndex, concurrency, promptText, maxBatchSize);
                return (success, peakUsage);
            }
            finally
            {
                // 停止监控
                monitorCts.Cancel();
                try { await monitorTask; } catch { }
                monitorCts.Dispose();
            }
        }

        /// <summary>
        /// 运行 HTTP 并发请求测试（带计时）
        /// 返回 (成功, 耗时ms, 吞吐量tokens/s)
        /// </summary>
        private async Task<(bool success, double elapsedMs, double tokensPerSec)> RunHttpConcurrencyTestWithTimingAsync(string apiUrl, int gpuIndex, int concurrency, string promptText)
        {
            try
            {
                var baseUrl = apiUrl.TrimEnd('/');
                if (!baseUrl.Contains("/translate/v1"))
                {
                    var uri = new Uri(baseUrl);
                    baseUrl = $"{uri.Scheme}://{uri.Authority}/translate/v1/batch-translate";
                }
                else if (!baseUrl.Contains("/batch-translate"))
                {
                    baseUrl += "/batch-translate";
                }

                var textList = new List<string>();
                for (int i = 0; i < concurrency; i++)
                {
                    textList.Add(promptText);
                }

                var requestData = new
                {
                    source_lang = "en",
                    target_lang = "zh",
                    text_list = textList.ToArray()
                };

                var jsonContent = JsonSerializer.Serialize(requestData);

                using var httpClient = new HttpClient
                {
                    Timeout = TimeSpan.FromSeconds(120)
                };

                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                
                // 计时
                var sw = System.Diagnostics.Stopwatch.StartNew();
                var response = await httpClient.PostAsync(baseUrl, content);
                sw.Stop();

                var elapsedMs = sw.Elapsed.TotalMilliseconds;

                if (!response.IsSuccessStatusCode)
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning($"HTTP 请求失败，状态码: {response.StatusCode}, 响应: {errorContent}");
                    return (false, elapsedMs, 0);
                }

                // 估算 tokens：每个 prompt 约按空格分词估算
                int promptTokens = promptText.Split(new[] { ' ', '\n', '\r', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
                int totalOutputTokens = concurrency * promptTokens; // 简单估算
                double tokensPerSec = elapsedMs > 0 ? (totalOutputTokens / elapsedMs * 1000.0) : 0;

                return (true, elapsedMs, tokensPerSec);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"HTTP 并发测试失败，并发数: {concurrency}");
                return (false, 0, 0);
            }
        }

        /// <summary>
        /// 运行 HTTP 并发请求测试（向后兼容版本）
        /// 支持分批发送：当GPU/显存占用率>=95%时，每批最高50并发
        /// </summary>
        private async Task<bool> RunHttpConcurrencyTestAsync(string apiUrl, int gpuIndex, int concurrency, string promptText, int maxBatchSize = 0)
        {
            try
            {
                // 构建 API URL（使用 /translate/v1/batch-translate 端点）
                var baseUrl = apiUrl.TrimEnd('/');
                if (!baseUrl.Contains("/translate/v1"))
                {
                    // 从基础 URL 构建翻译 API URL
                    var uri = new Uri(baseUrl);
                    baseUrl = $"{uri.Scheme}://{uri.Authority}/translate/v1/batch-translate";
                }
                else if (!baseUrl.Contains("/batch-translate"))
                {
                    baseUrl += "/batch-translate";
                }

                // 确定批次大小
                // maxBatchSize > 0 时使用指定值，否则一次性发送所有请求
                var effectiveBatchSize = maxBatchSize > 0 ? Math.Min(maxBatchSize, concurrency) : concurrency;
                var remaining = concurrency;
                var batchIndex = 0;

                while (remaining > 0)
                {
                    batchIndex++;
                    var currentBatch = Math.Min(remaining, effectiveBatchSize);
                    
                    // 构建文本列表
                    var textList = new List<string>();
                    for (int i = 0; i < currentBatch; i++)
                    {
                        textList.Add(promptText);
                    }

                    // 构建请求体
                    var requestData = new
                    {
                        source_lang = "en",
                        target_lang = "zh",
                        text_list = textList.ToArray()
                    };

                    var jsonContent = JsonSerializer.Serialize(requestData);

                    // 发起 batch 请求
                    using var httpClient = new HttpClient
                    {
                        Timeout = TimeSpan.FromSeconds(120)
                    };

                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                    var response = await httpClient.PostAsync(baseUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        _logger.LogWarning($"HTTP 请求失败（批次{batchIndex}，大小{currentBatch}），状态码: {response.StatusCode}, 响应: {errorContent}");
                        return false;
                    }

                    remaining -= currentBatch;
                    
                    if (effectiveBatchSize < concurrency && remaining > 0)
                    {
                        _logger.LogDebug($"批次{batchIndex}完成（{currentBatch}个请求），剩余{remaining}个");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"HTTP 并发测试失败，并发数: {concurrency}");
                return false;
            }
        }

        /// <summary>
        /// 检查当前 GPU 和 CPU 使用率
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task<ResourceUsage> CheckResourceUsageAsync(int gpuIndex)
        {
            var result = new ResourceUsage();

            try
            {
                // 获取 GPU 使用率
                var gpuInfo = await _gpuService.RefreshGpuStatusAsync(gpuIndex);
                if (gpuInfo != null)
                {
                    result.GpuUtilization = gpuInfo.GpuUtilization;
                    // 计算 GPU 显存使用率（基于专属显存）
                    if (gpuInfo.DedicatedMemoryBytes > 0)
                    {
                        result.GpuMemoryUtilization = (double)gpuInfo.UsedMemoryBytes / gpuInfo.DedicatedMemoryBytes * 100.0;
                    }
                    result.UseGpu = true;
                }
                else
                {
                    result.UseGpu = false;
                }

                // 获取 CPU 使用率
                using var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                cpuCounter.NextValue();
                await Task.Delay(500); // 等待采样
                result.CpuUtilization = cpuCounter.NextValue();

                // 获取内存使用率
                using var memoryCounter = new PerformanceCounter("Memory", "% Committed Bytes In Use");
                memoryCounter.NextValue();
                await Task.Delay(100);
                result.CpuMemoryUtilization = memoryCounter.NextValue();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "获取资源使用率失败");
            }

            return result;
        }

        /// <summary>
        /// 运行单次 benchmark 测试（使用指定 prompt）
        /// 在运行期间持续监控 GPU 显存峰值
        /// </summary>
        /// <returns>返回 (是否成功, 峰值显存字节数)</returns>
        private async Task<(bool success, long peakMemoryBytes)> RunBenchmarkTestWithPromptAsync(string benchmarkExePath, string vocabPath, string modelPath, int gpuIndex, int batchSize, string promptText)
        {
            try
            {
                var arguments = $"--weights \"{modelPath}\" " +
                               $"--vocab \"{vocabPath}\" " +
                               $"--batch-prompts-text \"{promptText}\" " +
                               $"--batch-size {batchSize} " +
                               $"--batch-steps 5 " +
                               $"--batch-temp 1.0";

                var startInfo = new ProcessStartInfo
                {
                    FileName = benchmarkExePath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = OutputEncoding,
                    StandardErrorEncoding = OutputEncoding,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(benchmarkExePath)
                };

                startInfo.EnvironmentVariables["CUDA_VISIBLE_DEVICES"] = gpuIndex.ToString();

                using var process = new Process { StartInfo = startInfo };
                var output = new System.Text.StringBuilder();
                var errors = new System.Text.StringBuilder();

                process.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        output.AppendLine(e.Data);
                };

                process.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                        errors.AppendLine(e.Data);
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                // 在 benchmark 运行期间持续监控 GPU 显存峰值
                var monitorTask = MonitorPeakGpuMemoryAsync(gpuIndex, process);

                // 等待 60 秒
                var completed = await Task.Run(() => process.WaitForExit(60000));

                if (!completed)
                {
                    process.Kill();
                    return (false, 0);
                }

                // 获取峰值显存
                long peakMemoryBytes = await monitorTask;

                // 检查退出码和输出
                bool success = process.ExitCode == 0 && !errors.ToString().Contains("out of memory");
                return (success, peakMemoryBytes);
            }
            catch
            {
                return (false, 0);
            }
        }

        /// <summary>
        /// 在进程运行期间持续监控 GPU 显存峰值
        /// </summary>
        private async Task<long> MonitorPeakGpuMemoryAsync(int gpuIndex, Process process)
        {
            long peakMemoryBytes = 0;
            var dedicatedMemoryBytes = 0L;

            try
            {
                // 获取专属显存总量
                var gpuInfo = await _gpuService.RefreshGpuStatusAsync(gpuIndex);
                if (gpuInfo != null)
                {
                    dedicatedMemoryBytes = gpuInfo.DedicatedMemoryBytes;
                }

                // 持续监控直到进程结束
                while (!process.HasExited)
                {
                    var currentGpu = await _gpuService.RefreshGpuStatusAsync(gpuIndex);
                    if (currentGpu != null && currentGpu.UsedMemoryBytes > peakMemoryBytes)
                    {
                        peakMemoryBytes = currentGpu.UsedMemoryBytes;
                    }

                    // 每 100ms 采样一次
                    await Task.Delay(100);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "监控 GPU 显存峰值时发生异常");
            }

            return peakMemoryBytes;
        }

        /// <summary>
        /// 基于GPU显存计算并发数
        /// </summary>
        private int CalculateGpuBasedConcurrency(GpuInfo gpu, ModelInfo model)
        {
            // 获取可用显存（GB）
            double availableMemory = gpu.FreeMemoryGB;

            // 获取模型基础显存需求
            double modelBaseMemory = model.EstimatedMemoryGB ?? EstimateBaseMemory(model);

            // 获取每并发额外显存消耗
            double perConcurrencyMemory = GetPerConcurrencyMemory(model);

            // 计算可用于并发的显存
            double memoryForConcurrency = availableMemory - modelBaseMemory;

            if (memoryForConcurrency <= 0)
            {
                _logger.LogWarning($"显存不足以运行模型: 可用={availableMemory:F2}GB, 需要={modelBaseMemory:F1}GB");
                return 1; // 最小并发数
            }

            // 计算最大并发数
            int concurrency = (int)Math.Floor(memoryForConcurrency / perConcurrencyMemory);

            // 确保至少为1
            return Math.Max(1, concurrency);
        }

        /// <summary>
        /// 估算模型基础显存需求
        /// </summary>
        private double EstimateBaseMemory(ModelInfo model)
        {
            // 从模型名称估算参数量
            var name = model.ModelName.ToLower();
            
            if (name.Contains("1.5b") || name.Contains("1.6b"))
                return 3.0;
            if (name.Contains("2.9b") || name.Contains("3b"))
                return 6.0;
            if (name.Contains("7b") || name.Contains("7.2b"))
                return 14.0;
            if (name.Contains("13b") || name.Contains("14b"))
                return 26.0;

            // 基于文件大小估算（假设FP16）
            // 每1GB参数约需要2GB显存
            double fileSizeGB = model.FileSizeGB;
            return fileSizeGB * 2.0;
        }

        /// <summary>
        /// 获取每并发额外显存消耗
        /// </summary>
        private double GetPerConcurrencyMemory(ModelInfo model)
        {
            var name = model.ModelName.ToLower();
            
            if (name.Contains("1.5b") || name.Contains("1.6b"))
                return 0.3;
            if (name.Contains("2.9b") || name.Contains("3b"))
                return 0.5;
            if (name.Contains("7b") || name.Contains("7.2b"))
                return 1.0;
            if (name.Contains("13b") || name.Contains("14b"))
                return 2.0;

            return 0.5;
        }

        /// <summary>
        /// 应用安全预留（预留10%）
        /// </summary>
        private int ApplySafetyMargin(int concurrency)
        {
            return (int)Math.Floor(concurrency * (1 - SafetyMargin));
        }

        /// <summary>
        /// 快速估算并发数（不刷新GPU状态）
        /// </summary>
        public int QuickEstimate(GpuInfo gpu, ModelInfo model)
        {
            double availableMemory = gpu.TotalMemoryGB * 0.8; // 假设80%可用
            double modelBaseMemory = model.EstimatedMemoryGB ?? EstimateBaseMemory(model);
            double perConcurrencyMemory = GetPerConcurrencyMemory(model);

            int concurrency = (int)Math.Floor((availableMemory - modelBaseMemory) / perConcurrencyMemory);
            concurrency = ApplySafetyMargin(concurrency);
            concurrency = Math.Min(concurrency, Environment.ProcessorCount * 2);

            return Math.Max(1, concurrency);
        }

        /// <summary>
        /// 获取推荐的并发数范围
        /// </summary>
        public (int Min, int Recommended, int Max) GetConcurrencyRange(GpuInfo gpu, ModelInfo model)
        {
            int min = 1;
            int recommended = QuickEstimate(gpu, model);
            int max = recommended * 2;

            // 限制最大值
            max = Math.Min(max, Environment.ProcessorCount * 4);
            max = Math.Min(max, 100); // 绝对上限

            return (min, recommended, max);
        }
    }
}
