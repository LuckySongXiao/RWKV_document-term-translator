using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DocumentTranslator.Services.Translation.Translators;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 翻译服务管理类，负责管理各种翻译器和翻译配置
    /// </summary>
    public class TranslationService
    {
        private readonly ILogger<TranslationService> _logger;
        private readonly Dictionary<string, ITranslator> _translators;
        private readonly Dictionary<string, object> _config;
        private string _currentTranslatorType;
        private volatile bool _stopFlag;
        private readonly List<string> _currentOperations;
        private readonly Dictionary<string, int?> _translatorParallelismOverride; // 每个翻译器的并发覆盖
        private readonly Dictionary<string, int?> _translatorSuggestedParallelism;

        public TranslationService(ILogger<TranslationService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _translators = new Dictionary<string, ITranslator>();
            _config = new Dictionary<string, object>();
            _currentOperations = new List<string>();
            _translatorParallelismOverride = new Dictionary<string, int?>();
            _translatorSuggestedParallelism = new Dictionary<string, int?>();
            _currentTranslatorType = "rwkv"; // 默认翻译器改为RWKV，避免误用外网服务

            // 先加载配置，再根据配置覆盖当前翻译器类型
            LoadConfiguration();
            try
            {
                if (_config.TryGetValue("current_translator_type", out var typeObj))
                {
                    var typeStr = typeObj?.ToString();
                    if (!string.IsNullOrWhiteSpace(typeStr))
                    {
                        _currentTranslatorType = typeStr;
                        _logger.LogInformation($"从配置加载当前翻译器类型: {_currentTranslatorType}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "从配置读取当前翻译器类型失败，将使用默认值 'rwkv'");
            }

            InitializeTranslators();
        }

        /// <summary>
        /// 当前翻译器类型
        /// </summary>
        public string CurrentTranslatorType
        {
            get => _currentTranslatorType;
            set
            {
                if (_translators.ContainsKey(value))
                {
                    _currentTranslatorType = value;
                    _logger.LogInformation($"切换到翻译器: {value}");
                }
                else
                {
                    _logger.LogWarning($"翻译器类型 {value} 不存在");
                }
            }
        }

        /// <summary>
        /// 获取当前翻译器
        /// </summary>
        public ITranslator CurrentTranslator => _translators.TryGetValue(_currentTranslatorType, out var translator) ? translator : null;

        /// <summary>
        /// 获取所有可用的翻译器类型
        /// </summary>
        /// <summary>
        /// 计算建议的最大并发数（基于本地资源和当前翻译器类型）
        /// </summary>
        public int GetSuggestedMaxParallelism()
        {
            try
            {
                var type = _currentTranslatorType?.ToLowerInvariant() ?? "";

                if (_translatorSuggestedParallelism.TryGetValue(type, out var runtimeSuggested) && runtimeSuggested.HasValue)
                {
                    return Math.Max(1, runtimeSuggested.Value);
                }

                var cores = Math.Max(1, Environment.ProcessorCount);
                int baseSuggestion = type == "rwkv"
                    ? Math.Max(2, Math.Min(cores * 2, 30))
                    : Math.Max(2, Math.Min(cores * 2, 8));

                try
                {
                    var gcInfo = GC.GetGCMemoryInfo();
                    var total = gcInfo.TotalAvailableMemoryBytes;
                    var memoryCap = type == "rwkv" ? 30 : 16;
                    var byMem = (int)Math.Max(1, Math.Min(memoryCap, total / (512L * 1024 * 1024)));
                    baseSuggestion = Math.Max(1, Math.Min(baseSuggestion, byMem));
                }
                catch { }

                return baseSuggestion;
            }
            catch
            {
                return Math.Max(1, Environment.ProcessorCount);
            }
        }

        /// <summary>
        /// 获取当前有效的最大并发数（考虑手动覆盖）
        /// 只对RWKV引擎应用自定义并发，其他引擎使用建议值
        /// </summary>
        public int GetEffectiveMaxParallelism()
        {
            var type = _currentTranslatorType?.ToLowerInvariant() ?? "";
            if (_translatorParallelismOverride.TryGetValue(type, out var parallelismOverride) && parallelismOverride.HasValue)
                return Math.Max(1, parallelismOverride.Value);
            return GetSuggestedMaxParallelism();
        }

        /// <summary>
        /// 设置指定翻译器的并发覆盖（只对RWKV引擎有效）
        /// </summary>
        public void SetMaxParallelismOverride(int? value)
        {
            var type = _currentTranslatorType?.ToLowerInvariant() ?? "";
            _translatorParallelismOverride[type] = value.HasValue ? Math.Max(1, value.Value) : null;
            _logger.LogInformation($"{type}引擎并发覆盖设置为: {_translatorParallelismOverride[type]?.ToString() ?? "未设置（使用建议值）"}");
        }

        public void SetSuggestedMaxParallelism(int? value, string translatorType = null)
        {
            var type = (translatorType ?? _currentTranslatorType ?? "rwkv").ToLowerInvariant();
            _translatorSuggestedParallelism[type] = value.HasValue ? Math.Max(1, value.Value) : null;
            _logger.LogInformation($"{type}引擎建议并发设置为: {_translatorSuggestedParallelism[type]?.ToString() ?? "未设置（使用自动估算）"}");
        }

        /// <summary>
        /// 每批次并发请求上限（GPU资源超警戒时使用）
        /// </summary>
        public int BatchConcurrencyLimit { get; set; } = 50;

        public IEnumerable<string> AvailableTranslatorTypes => _translators.Keys;

        /// <summary>
        /// 重新初始化指定的翻译器
        /// </summary>
        public void ReinitializeTranslator(string translatorType)
        {
            try
            {
                _logger.LogInformation($"重新初始化{translatorType}翻译器");
                LoadConfiguration();

                if (_translators.ContainsKey(translatorType))
                {
                    _translators.Remove(translatorType);
                }

                ITranslator translator = translatorType switch
                {
                    "rwkv" => InitializeRWKVTranslator(),
                    _ => null
                };

                if (translator != null)
                {
                    _translators[translatorType] = translator;
                    _logger.LogInformation($"{translatorType}翻译器重新初始化成功");
                }
                else
                {
                    _logger.LogWarning($"{translatorType}翻译器重新初始化失败");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"重新初始化{translatorType}翻译器失败");
            }
        }

        /// <summary>
        /// 使用偏好模型重新初始化指定翻译器（不必保存到配置文件）
        /// </summary>
        /// <param name="translatorType">翻译器类型</param>
        /// <param name="preferredModel">优先使用的模型名</param>
        /// <param name="setCurrent">是否将其设置为当前翻译器</param>
        public void ReinitializeTranslator(string translatorType, string preferredModel, bool setCurrent = false)
        {
            try
            {
                _logger.LogInformation($"以偏好模型重新初始化 {translatorType} 翻译器，模型: {preferredModel}");
                LoadConfiguration();

                if (_translators.ContainsKey(translatorType))
                {
                    _translators.Remove(translatorType);
                }

                ITranslator translator = translatorType switch
                {
                    "rwkv" => InitializeRWKVTranslator(preferredModel),
                    _ => null
                };

                if (translator != null)
                {
                    _translators[translatorType] = translator;
                    if (setCurrent)
                    {
                        _currentTranslatorType = translatorType;
                    }
                    _logger.LogInformation($"{translatorType} 翻译器已用偏好模型重新初始化");
                }
                else
                {
                    _logger.LogWarning($"{translatorType} 翻译器用偏好模型重新初始化失败");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"以偏好模型重新初始化 {translatorType} 翻译器失败");
            }
        }

        /// <summary>
        /// 加载配置文件
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                if (File.Exists(configPath))
                {
                    var configJson = File.ReadAllText(configPath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, object>>(configJson);
                    if (config != null)
                    {
                        foreach (var kvp in config)
                        {
                            _config[kvp.Key] = kvp.Value;
                        }
                    }
                }
                _logger.LogInformation("配置文件加载完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载配置文件失败");
            }
        }

        /// <summary>
        /// 初始化所有翻译器
        /// </summary>
        private void InitializeTranslators()
        {
            try
            {
                // 初始化RWKV翻译器
                var rwkvTranslator = InitializeRWKVTranslator();
                if (rwkvTranslator != null)
                {
                    _translators["rwkv"] = rwkvTranslator;
                }

                _logger.LogInformation($"已初始化 {_translators.Count} 个翻译器");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化翻译器失败");
            }
        }

        /// <summary>
        /// 初始化RWKV翻译器
        /// </summary>
        private ITranslator InitializeRWKVTranslator(string preferredModel = null)
        {
            try
            {
                var rwkvConfig = GetConfig("rwkv_translator") as Dictionary<string, object> ?? new Dictionary<string, object>();
                var apiUrl = rwkvConfig.GetValueOrDefault("api_url", "").ToString();
                var apiKey = rwkvConfig.GetValueOrDefault("api_key", "").ToString();
                var modelFromConfig = rwkvConfig.GetValueOrDefault("model", "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118").ToString();
                var timeoutFromConfig = Convert.ToInt32(rwkvConfig.GetValueOrDefault("timeout", 60));
                var useBatchTranslate = Convert.ToBoolean(rwkvConfig.GetValueOrDefault("use_batch_translate", true));

                try
                {
                    var cfgPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config", "rwkv_api.json");
                    if (System.IO.File.Exists(cfgPath))
                    {
                        var json = System.IO.File.ReadAllText(cfgPath);
                        var cfg = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();

                        var apiUrlCandidate = cfg.GetValueOrDefault("api_url", "").ToString();
                        if (!string.IsNullOrWhiteSpace(apiUrlCandidate))
                        {
                            apiUrl = apiUrlCandidate;
                        }

                        var modelCandidate = cfg.GetValueOrDefault("default_model", string.Empty)?.ToString();
                        if (!string.IsNullOrWhiteSpace(modelCandidate))
                            modelFromConfig = modelCandidate;

                        var useBatchTranslateCandidate = cfg.GetValueOrDefault("use_batch_translate", null);
                        if (useBatchTranslateCandidate != null)
                            useBatchTranslate = Convert.ToBoolean(useBatchTranslateCandidate);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "读取 API_config/rwkv_api.json 失败");
                }

                // 若仍未获取到 apiUrl，使用内置默认地址（免密本地局域网 OpenAI 兼容服务）
                if (string.IsNullOrWhiteSpace(apiUrl))
                {
                    apiUrl = "http://localhost:8000/translate/v1/batch-translate";
                    _logger.LogInformation("使用内置默认RWKV_API地址: {ApiUrl}", apiUrl);
                }

                if (!string.IsNullOrWhiteSpace(apiUrl))
                {
                    var model = preferredModel ?? modelFromConfig;
                    var timeout = timeoutFromConfig;
                    var maxConcurrency = GetEffectiveMaxParallelism();

                    _logger.LogInformation($"初始化RWKV翻译器，使用模型: {model}, API: {apiUrl}, 批量翻译模式: {useBatchTranslate}, 并发数: {maxConcurrency}");
                    
                    // 判断是否为llama.cpp推理服务（URL包含/v1/chat/completions或端口为8080）
                    var isLlamaCpp = apiUrl.Contains("/v1/chat/completions", StringComparison.OrdinalIgnoreCase) ||
                                     apiUrl.Contains("/v1/completions", StringComparison.OrdinalIgnoreCase) ||
                                     (apiUrl.Contains(":8080") && !apiUrl.Contains("/translate/"));
                    
                    var translator = new RWKVTranslator(_logger, apiUrl, model, timeout, apiKey, useBatchTranslate, maxConcurrency, isLlamaCpp);
                    RWKVTranslator.GetMaxConcurrency = GetEffectiveMaxParallelism;
                    return translator;
                }
                else
                {
                    _logger.LogWarning("未配置RWKV_API地址(api_url)，无法初始化RWKV翻译器");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化RWKV翻译器失败");
            }
            return null;
        }

        /// <summary>
        /// 获取API密钥
        /// </summary>
        private string GetApiKey(string service)
        {
            try
            {
                var envKeyName = $"{service.ToUpper()}_API_KEY";
                var envApiKey = Environment.GetEnvironmentVariable(envKeyName, EnvironmentVariableTarget.User);
                if (!string.IsNullOrEmpty(envApiKey))
                {
                    _logger.LogDebug($"从环境变量获取{service} API密钥");
                    return envApiKey;
                }

                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config", $"{service}_api.json");
                if (File.Exists(configPath))
                {
                    var configJson = File.ReadAllText(configPath);
                    var config = JsonConvert.DeserializeObject<Dictionary<string, string>>(configJson);
                    var apiKey = config?.GetValueOrDefault("api_key", "");
                    if (!string.IsNullOrEmpty(apiKey))
                    {
                        _logger.LogDebug($"从配置文件获取{service} API密钥");
                        return apiKey;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"获取{service} API密钥失败");
            }
            return string.Empty;
        }

        /// <summary>
        /// 获取配置项
        /// </summary>
        private object GetConfig(string key)
        {
            return _config.TryGetValue(key, out var value) ? value : null;
        }

        /// <summary>
        /// 智能清理文本：提取指定语言内容
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <param name="keepLanguage">保留的语言代码 (如 "zh", "en")</param>
        /// <returns>清理后的文本</returns>
        public virtual async Task<string> SmartCleanTextAsync(string text, string keepLanguage)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;

            // 1. 过滤：纯数字/符号、公式
            // 纯数字/符号 (允许空格、数字、标点、加减乘除、百分号)
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[\d\s\p{P}\+\-\*/\.\,]+$"))
                return text;
            
            // 公式 (简单判断 LaTeX 或 常见数学符号，如 $$...$$, \begin...\end)
            if (text.Contains("$$") || text.Contains(@"\begin") || text.Contains(@"\end") || 
                (text.Trim().StartsWith("$") && text.Trim().EndsWith("$")) ||
                System.Text.RegularExpressions.Regex.IsMatch(text, @"^\s*[\$\\]"))
                return text;

            // 2. 获取语言名称
            string langName = keepLanguage;
            if (LanguageDetector.LanguageDefinitions.TryGetValue(keepLanguage, out var def))
            {
                langName = def.Name;
            }
            // 处理英文名为 English 以获得更好效果
            if (keepLanguage == "en" || langName == "英文") langName = "English";
            if (keepLanguage == "zh" || langName == "中文") langName = "Simplified Chinese";

            // 3. 构建提示词
            // 使用明确的指令，要求仅输出结果，不包含解释
            string prompt = $"You are a text cleaning assistant. Your task is to extract ONLY the {langName} content from the text provided below.\n" +
                            $"Rules:\n" +
                            $"1. Remove all content in languages other than {langName}.\n" +
                            $"2. Return ONLY the extracted {langName} text.\n" +
                            $"3. Do NOT add any explanations, notes, or headers.\n" +
                            $"4. If the text contains only {langName}, return it as is.\n" +
                            $"5. If the text does not contain {langName}, return an empty string.\n" +
                            $"Text to process:\n{text}";

            // 4. 调用模型
            try 
            {
                // 使用 ChatAsync 进行指令交互
                // 注意：ChatAsync 可能会带上下文，这里不需要上下文，传 null
                var result = await ChatAsync(prompt, null);
                
                // 清理可能的前后缀（有些模型喜欢加 "Here is..." 或 markdown 代码块）
                if (string.IsNullOrWhiteSpace(result)) return "";

                result = result.Trim();
                // 去除可能Markdown代码块标记
                if (result.StartsWith("```"))
                {
                    var lines = result.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    if (lines.Count > 1)
                    {
                        // 去掉第一行（```xxx）
                        lines.RemoveAt(0);
                        // 去掉最后一行（```）
                        if (lines.Count > 0 && lines.Last().Trim().StartsWith("```")) 
                            lines.RemoveAt(lines.Count - 1);
                        result = string.Join("\n", lines);
                    }
                    else
                    {
                         // 只有一行且是 ``` 开头？不太可能，或者是 ```text content```
                         result = result.Trim('`');
                         if (result.StartsWith("xml") || result.StartsWith("json") || result.StartsWith("text"))
                             result = result.Substring(result.IndexOf(' ') + 1);
                    }
                }
                
                return result.Trim();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "智能清理文本失败");
                return text; // 出错时保留原文
            }
        }

        /// <summary>
        /// 翻译文本
        /// </summary>
        public virtual async Task<string> TranslateTextAsync(string text, Dictionary<string, string> terminologyDict = null,
            string sourceLang = "zh", string targetLang = "en", string prompt = null, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // 检查是否需要停止
            if (_stopFlag)
            {
                _logger.LogInformation("翻译操作被停止");
                return string.Empty;
            }

            var translator = CurrentTranslator;
            if (translator == null)
            {
                _logger.LogError($"未找到{_currentTranslatorType}翻译器");
                throw new InvalidOperationException($"未找到可用的翻译器");
            }

            try
            {
                return await translator.TranslateAsync(text, terminologyDict, sourceLang, targetLang, prompt, originalText);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"{_currentTranslatorType}翻译失败");
                throw;
            }
        }

        /// <summary>
        /// 答疑助手对话
        /// </summary>
        public virtual async Task<string> ChatAsync(string question, string context = null)
        {
            if (string.IsNullOrWhiteSpace(question))
                return string.Empty;

            // 检查是否需要停止
            if (_stopFlag)
            {
                _logger.LogInformation("对话操作被停止");
                return string.Empty;
            }

            var translator = CurrentTranslator;
            if (translator == null)
            {
                _logger.LogError($"未找到{_currentTranslatorType}翻译器");
                throw new InvalidOperationException($"未找到可用的翻译器");
            }

            try
            {
                _logger.LogInformation($"使用 {_currentTranslatorType} 引擎进行对话");
                return await translator.ChatAsync(question, context);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"{_currentTranslatorType}对话失败");
                throw;
            }
        }

        /// <summary>
        /// 获取可用模型列表
        /// </summary>
        public async Task<List<string>> GetAvailableModelsAsync(string translatorType = null)
        {
            var type = translatorType ?? _currentTranslatorType;
            if (_translators.TryGetValue(type, out var translator))
            {
                try
                {
                    return await translator.GetAvailableModelsAsync();
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, $"获取{type}模型列表失败");
                }
            }
            return new List<string>();
        }

        /// <summary>
        /// 停止当前操作
        /// </summary>
        public void StopCurrentOperations()
        {
            try
            {
                _stopFlag = true;
                _logger.LogInformation("设置停止标志，正在停止当前翻译操作...");

                // 清理当前操作列表
                lock (_currentOperations)
                {
                    _currentOperations.Clear();
                }

                // 重置停止标志
                Task.Run(async () =>
                {
                    await Task.Delay(1000); // 等待1秒后重置
                    _stopFlag = false;
                    _logger.LogInformation("停止标志已重置");
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停止当前操作失败");
            }
        }
    }
}
