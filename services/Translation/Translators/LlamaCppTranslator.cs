using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DocumentTranslator.Services.Translation.Translators
{
    /// <summary>
    /// llama.cpp 翻译器，通过 OpenAI 兼容接口调用 llama-server
    /// 支持 chat 模式 (/v1/chat/completions) 和续写模式 (/v1/completions)
    /// 
    /// 翻译规则基于 rwkv-world 对话模板，与 RWKVTranslator 的翻译效果一致：
    /// rwkv-world 模板格式:
    ///   System: {system}\n\n
    ///   User: {user}\n\n
    ///   Assistant: {response}
    ///   
    /// rwkv_lightning 翻译端点使用原生格式: {src_lang}: {text}\n\n{tgt_lang}:
    /// 本翻译器通过 rwkv-world 模板包装等效的翻译 prompt，使最终输入模型的
    /// token 序列与 rwkv_lightning 翻译端点一致，确保翻译效果相同。
    /// </summary>
    public class LlamaCppTranslator : BaseTranslator
    {
        private readonly string _apiUrl;
        private readonly string _model;
        private readonly int _timeout;
        private readonly string _apiKey;
        private readonly LlamaCppMode _mode;
        private readonly HttpClient _httpClient;
        private static System.Threading.SemaphoreSlim _globalRequestSemaphore;
        private static readonly object _semaphoreLock = new object();
        private static int _maxConcurrency;
        private static Func<int> _getMaxConcurrency;

        // 可配置的翻译规则
        private readonly string _eosToken;
        private readonly Dictionary<string, string> _langSeparators;

        public static int MaxConcurrency => _maxConcurrency;
        public static Func<int> GetMaxConcurrency
        {
            get => _getMaxConcurrency;
            set => _getMaxConcurrency = value;
        }

        /// <summary>
        /// llama.cpp 推理模式
        /// </summary>
        public enum LlamaCppMode
        {
            /// <summary>
            /// 续写模式 (/v1/completions) - 直接构造与 rwkv_lightning 等效的裸 prompt
            /// </summary>
            Completions,
            /// <summary>
            /// Chat 模式 (/v1/chat/completions) - 通过 rwkv-world 模板包装翻译 prompt
            /// </summary>
            Chat
        }

        public LlamaCppTranslator(ILogger logger, string apiUrl, string model = "",
            int timeout = 60, string apiKey = null, int maxConcurrency = 4,
            LlamaCppMode mode = LlamaCppMode.Completions,
            string eosToken = null,
            Dictionary<string, string> langSeparators = null)
            : base(logger)
        {
            _model = model ?? string.Empty;
            _timeout = timeout;
            _apiKey = apiKey ?? string.Empty;
            _mode = mode;
            _eosToken = eosToken ?? string.Empty;
            _langSeparators = langSeparators ?? new Dictionary<string, string>();

            lock (_semaphoreLock)
            {
                if (_globalRequestSemaphore == null || _maxConcurrency != maxConcurrency)
                {
                    _maxConcurrency = maxConcurrency;
                    _globalRequestSemaphore = new System.Threading.SemaphoreSlim(maxConcurrency);
                    _logger.LogInformation($"LlamaCpp全局请求信号量已更新为: {maxConcurrency}");
                }
            }

            _apiUrl = NormalizeApiUrl(apiUrl);

            _httpClient = new HttpClient { Timeout = TimeSpan.FromSeconds(timeout) };
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "DocumentTranslator/1.0");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            _httpClient.DefaultRequestHeaders.Add("Connection", "keep-alive");

            if (!string.IsNullOrEmpty(_apiKey))
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            }

            _logger.LogInformation($"初始化LlamaCpp翻译器，模型: {model}, API URL: {_apiUrl}, 模式: {_mode}, 并发: {maxConcurrency}");
        }

        /// <summary>
        /// 规范化 API URL，提取基础地址并拼接正确的端点路径
        /// </summary>
        private string NormalizeApiUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            var u = url.TrimEnd('/');
            var lower = u.ToLowerInvariant();

            // 从URL中提取基础地址（去掉已有的API路径）
            if (lower.Contains("/translate/v1/batch-translate"))
            {
                u = u.Substring(0, lower.IndexOf("/translate/v1/batch-translate"));
            }
            else if (lower.Contains("/translate/v1"))
            {
                u = u.Substring(0, lower.IndexOf("/translate/v1"));
            }
            else if (lower.Contains("/chat/completions"))
            {
                var chatIndex = lower.IndexOf("/chat/completions");
                u = u.Substring(0, chatIndex);
                if (u.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ||
                    u.EndsWith("/v2", StringComparison.OrdinalIgnoreCase) ||
                    u.EndsWith("/v3", StringComparison.OrdinalIgnoreCase))
                {
                    u = u.Substring(0, u.LastIndexOf('/'));
                }
            }
            else if (lower.Contains("/completions"))
            {
                var compIndex = lower.IndexOf("/completions");
                u = u.Substring(0, compIndex);
                if (u.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ||
                    u.EndsWith("/v2", StringComparison.OrdinalIgnoreCase))
                {
                    u = u.Substring(0, u.LastIndexOf('/'));
                }
            }

            // 拼接端点路径
            return _mode == LlamaCppMode.Chat
                ? u + "/v1/chat/completions"
                : u + "/v1/completions";
        }

        private static string GetLanguageCode(string lang)
        {
            switch (lang.ToLower())
            {
                case "zh":
                case "zh-cn":
                    return "zh-CN";
                case "ja":
                case "jp":
                    return "ja";
                case "en":
                    return "en";
                case "ko":
                    return "ko";
                case "fr":
                    return "fr";
                case "de":
                    return "de";
                case "es":
                    return "es";
                case "ru":
                    return "ru";
                default:
                    return "en";
            }
        }

        /// <summary>
        /// 构造翻译 prompt
        /// 
        /// RWKV 官方续写翻译 prompt 格式（参考 https://www.rwkv.cn/docs/RWKV-Prompts/Completion-Prompts）：
        ///   Japanese:
        ///   春の初め、桜の花が満開になる頃...
        ///   
        ///   English:
        /// 
        /// 即：{源语言全名}:\n{源文本}\n\n{目标语言全名}:
        /// 
        /// 可通过翻译规则配置自定义语言名称和分隔符。
        /// 默认使用语言全名（English、Chinese、Japanese 等）。
        /// </summary>
        private string CreateTranslationPrompt(string sourceLang, string targetLang, string text)
        {
            // 获取源语言名称（如 "Chinese"、"English"、"Japanese"）
            var srcName = GetLangName(sourceLang);
            // 获取目标语言分隔符（如 "\n\nEnglish"）
            var tgtSeparator = GetLangSeparator(targetLang);

            // RWKV 官方格式: {LangName}:\n{text}\n\n{TargetLangName}:
            return $"{srcName}:\n{text}{tgtSeparator}:";
        }

        /// <summary>
        /// 获取语言全名（用于 prompt 中的语言标识）
        /// 优先使用配置的分隔符中提取的语言名称，否则使用内置映射
        /// </summary>
        private string GetLangName(string langCode)
        {
            // 优先使用配置的分隔符中的语言名称部分
            if (_langSeparators.TryGetValue(langCode, out var separator))
            {
                // 分隔符格式如 "\n\nEnglish"，提取语言名称部分 "English"
                var langName = separator.TrimStart('\n', '\r', ' ', ':');
                if (!string.IsNullOrEmpty(langName))
                    return langName;
            }
            // 回退到内置语言全名映射
            return GetLanguageFullName(langCode);
        }

        /// <summary>
        /// 获取语言全名（RWKV 官方续写翻译格式使用语言全名而非代码）
        /// </summary>
        private static string GetLanguageFullName(string langCode)
        {
            switch (langCode.ToLower())
            {
                case "zh":
                case "zh-cn":
                case "cht":
                    return "Chinese";
                case "ja":
                case "jp":
                    return "Japanese";
                case "en":
                    return "English";
                case "ko":
                    return "Korean";
                case "fr":
                    return "French";
                case "de":
                    return "German";
                case "es":
                    return "Spanish";
                case "ru":
                    return "Russian";
                case "vi":
                    return "Vietnamese";
                case "pt":
                    return "Portuguese";
                case "it":
                    return "Italian";
                case "ar":
                    return "Arabic";
                case "th":
                    return "Thai";
                default:
                    return "English";
            }
        }

        /// <summary>
        /// 获取目标语言分隔符（包含换行和语言名称）
        /// 例如: "\n\nEnglish" 或 "\n\nChinese" 或默认的 "\n\nEnglish"
        /// </summary>
        private string GetLangSeparator(string langCode)
        {
            if (_langSeparators.TryGetValue(langCode, out var separator))
            {
                return separator;
            }
            // 回退到默认格式: \n\n{LanguageFullName}
            return $"\n\n{GetLanguageFullName(langCode)}";
        }

        /// <summary>
        /// 剥离模型输出中可能回显的 prompt 内容
        /// 
        /// Completions 模式下，某些模型可能会在输出开头回显部分或全部 prompt，
        /// 例如输出 "Chinese:\n原文\n\nEnglish:\n译文" 而非仅 "译文"。
        /// 此方法检测并剥离这类回显内容。
        /// </summary>
        private string StripPromptEcho(string output, string sourceCode, string targetCode, string sourceText)
        {
            if (string.IsNullOrWhiteSpace(output))
                return output;

            var trimmedOutput = output.TrimStart();

            // 获取语言全名用于回显检测
            var srcName = GetLangName(sourceCode);
            var tgtName = GetLangName(targetCode);

            // 检测完整的 prompt 回显: "{SrcName}:\n{text}\n\n{TgtName}:"
            var fullPromptPrefix = $"{srcName}:\n{sourceText}";
            if (trimmedOutput.StartsWith(fullPromptPrefix, StringComparison.OrdinalIgnoreCase))
            {
                var afterPrompt = trimmedOutput.Substring(fullPromptPrefix.Length).TrimStart('\n', '\r', ' ');
                // 跳过目标语言前缀
                if (afterPrompt.StartsWith($"{tgtName}:", StringComparison.OrdinalIgnoreCase))
                {
                    afterPrompt = afterPrompt.Substring(tgtName.Length + 1).TrimStart('\n', ' ');
                }
                _logger.LogDebug($"检测到完整prompt回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterPrompt.Length}");
                return afterPrompt;
            }

            // 检测源语言前缀回显: "{SrcName}:\n..."
            if (trimmedOutput.StartsWith($"{srcName}:", StringComparison.OrdinalIgnoreCase))
            {
                // 查找目标语言前缀的位置
                var targetPrefix = $"\n\n{tgtName}:";
                var targetIdx = trimmedOutput.IndexOf(targetPrefix, StringComparison.OrdinalIgnoreCase);
                if (targetIdx >= 0)
                {
                    var afterTarget = trimmedOutput.Substring(targetIdx + targetPrefix.Length).TrimStart('\n', ' ');
                    _logger.LogDebug($"检测到源语言前缀回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterTarget.Length}");
                    return afterTarget;
                }

                // 也检查单行格式: "{SrcName}: xxx {TgtName}: yyy"
                var targetPrefixSingle = $"{tgtName}:";
                var targetIdx2 = trimmedOutput.IndexOf(targetPrefixSingle, srcName.Length + 2, StringComparison.OrdinalIgnoreCase);
                if (targetIdx2 >= 0)
                {
                    var afterTarget = trimmedOutput.Substring(targetIdx2 + targetPrefixSingle.Length).TrimStart('\n', ' ');
                    _logger.LogDebug($"检测到单行格式回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterTarget.Length}");
                    return afterTarget;
                }
            }

            // 检测目标语言前缀回显: "{TgtName}:\n译文"
            if (trimmedOutput.StartsWith($"{tgtName}:", StringComparison.OrdinalIgnoreCase))
            {
                var afterTarget = trimmedOutput.Substring(tgtName.Length + 1).TrimStart('\n', ' ');
                _logger.LogDebug($"检测到目标语言前缀回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterTarget.Length}");
                return afterTarget;
            }

            // 兼容旧格式：使用语言代码的回显
            if (trimmedOutput.StartsWith($"{sourceCode}:", StringComparison.OrdinalIgnoreCase))
            {
                var targetPrefix = $"\n\n{targetCode}:";
                var targetIdx = trimmedOutput.IndexOf(targetPrefix, StringComparison.OrdinalIgnoreCase);
                if (targetIdx >= 0)
                {
                    var afterTarget = trimmedOutput.Substring(targetIdx + targetPrefix.Length).TrimStart(' ');
                    _logger.LogDebug($"检测到旧格式源语言前缀回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterTarget.Length}");
                    return afterTarget;
                }
            }

            if (trimmedOutput.StartsWith($"{targetCode}:", StringComparison.OrdinalIgnoreCase))
            {
                var afterTarget = trimmedOutput.Substring(targetCode.Length + 1).TrimStart(' ');
                _logger.LogDebug($"检测到旧格式目标语言前缀回显，已剥离。原始输出长度: {output.Length}, 剥离后长度: {afterTarget.Length}");
                return afterTarget;
            }

            return output;
        }

        /// <summary>
        /// 构建 stop 序列列表
        /// 包含：可配置的 EOS token、源语言名称（防止模型重复生成）
        /// </summary>
        private List<string> BuildStopSequences(string sourceLang)
        {
            var stopSequences = new List<string>();

            // 添加可配置的 EOS token
            if (!string.IsNullOrWhiteSpace(_eosToken))
            {
                stopSequences.Add(_eosToken);
            }

            // 添加源语言名称作为 stop 序列，防止模型过度生成
            // RWKV 官方格式使用语言全名（如 Chinese、English、Japanese）
            var srcName = GetLangName(sourceLang);
            stopSequences.Add($"\n\n{srcName}:");
            stopSequences.Add($"\n{srcName}:");

            // 同时添加语言代码格式的 stop 序列（兼容旧格式）
            var srcCode = GetLanguageCode(sourceLang);
            if (srcCode != srcName)
            {
                stopSequences.Add($"\n\n{srcCode}:");
                stopSequences.Add($"\n{srcCode}:");
            }

            return stopSequences;
        }

        private void UpdateSemaphoreIfNeeded(int currentMaxConcurrency)
        {
            if (currentMaxConcurrency != _maxConcurrency)
            {
                lock (_semaphoreLock)
                {
                    if (currentMaxConcurrency != _maxConcurrency)
                    {
                        _maxConcurrency = currentMaxConcurrency;
                        _globalRequestSemaphore = new System.Threading.SemaphoreSlim(currentMaxConcurrency);
                        _logger.LogInformation($"LlamaCpp全局请求信号量已动态更新为: {currentMaxConcurrency}");
                    }
                }
            }

            if (_globalRequestSemaphore == null)
            {
                lock (_semaphoreLock)
                {
                    if (_globalRequestSemaphore == null)
                    {
                        _maxConcurrency = currentMaxConcurrency;
                        _globalRequestSemaphore = new System.Threading.SemaphoreSlim(currentMaxConcurrency);
                        _logger.LogInformation($"LlamaCpp全局请求信号量已初始化: {currentMaxConcurrency}");
                    }
                }
            }
        }

        #region TranslateAsync

        public override async Task<string> TranslateAsync(string text, Dictionary<string, string> terminologyDict = null,
            string sourceLang = "zh", string targetLang = "en", string prompt = null, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var textForLog = originalText ?? text;

            // 预处理术语
            string finalPromptText = text;
            if (terminologyDict != null && terminologyDict.Any())
            {
                finalPromptText = ReplaceTermsInText(text, terminologyDict, sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh"));
            }

            // 检查是否需要跳过翻译（与 RWKVTranslator 一致）
            if (ShouldSkipTranslation(text))
                return text;

            if (!ShouldTranslateBasedOnLanguageRatio(text, sourceLang))
                return text;

            int maxRetries = 3;
            int retryDelayMs = 1000;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var currentMaxConcurrency = _getMaxConcurrency?.Invoke() ?? _maxConcurrency;
                    UpdateSemaphoreIfNeeded(currentMaxConcurrency);

                    var currentSemaphore = _globalRequestSemaphore;
                    await currentSemaphore.WaitAsync();
                    try
                    {
                        string translation;

                        if (_mode == LlamaCppMode.Chat)
                        {
                            translation = await TranslateWithChatApiAsync(finalPromptText, sourceLang, targetLang);
                        }
                        else
                        {
                            translation = await TranslateWithCompletionApiAsync(finalPromptText, sourceLang, targetLang);
                        }

                        // 后处理：与 RWKVTranslator 使用相同的 FilterOutput
                        var processedTranslation = FilterOutput(translation, sourceLang, targetLang);
                        if (string.IsNullOrWhiteSpace(processedTranslation))
                        {
                            processedTranslation = translation;
                        }

                        // 质量检测（与 RWKVTranslator 一致）
                        processedTranslation = ProcessAbnormalTranslation(text, processedTranslation, sourceLang, targetLang);

                        LogTranslation(textForLog, processedTranslation, sourceLang, targetLang);
                        if (finalPromptText != text)
                        {
                            LogPreprocessedTranslation(finalPromptText, processedTranslation, sourceLang, targetLang);
                        }

                        return processedTranslation;
                    }
                    finally
                    {
                        currentSemaphore.Release();
                    }
                }
                catch (Exception ex) when (attempt < maxRetries - 1)
                {
                    _logger.LogWarning(ex, $"LlamaCpp翻译失败(尝试 {attempt + 1}/{maxRetries})");
                    await Task.Delay(retryDelayMs);
                    retryDelayMs *= 2;
                }
            }

            throw new Exception($"LlamaCpp翻译在 {maxRetries} 次尝试后仍然失败");
        }

        /// <summary>
        /// 使用 chat 接口 (/v1/chat/completions) 翻译
        /// 
        /// 通过 rwkv-world 对话模板包装翻译 prompt，使最终输入模型的 token 序列
        /// 与 rwkv_lightning 的 /translate/v1/batch-translate 端点等效。
        /// 
        /// rwkv-world 模板会将 messages 格式化为:
        ///   System: {system}\n\nUser: {user}\n\nAssistant: 
        /// 
        /// 等效策略: 将翻译 prompt "{src}: {text}\n\n{tgt}:" 放入 User 消息，
        /// 不设 System 消息。最终模型输入为:
        ///   User: {src}: {text}\n\n{tgt}:\n\nAssistant: 
        /// 
        /// 这与 rwkv_lightning 翻译端点的裸 prompt "{src}: {text}\n\n{tgt}:"
        /// 在模型看来是等效的（仅多了 "User: " 前缀和 "Assistant: " 引导），
        /// RWKV 翻译模型在此格式下能正确输出纯净译文。
        /// </summary>
        private async Task<string> TranslateWithChatApiAsync(string text, string sourceLang, string targetLang)
        {
            var sourceCode = GetLanguageCode(sourceLang);
            var targetCode = GetLanguageCode(targetLang);

            // 构造翻译 prompt（使用可配置的语言分隔符）
            var translationPrompt = CreateTranslationPrompt(sourceLang, targetLang, text);

            // 构建 stop 序列
            var stopSequences = BuildStopSequences(sourceLang);

            var requestData = new Dictionary<string, object>
            {
                ["messages"] = new[]
                {
                    new { role = "user", content = translationPrompt }
                },
                ["max_tokens"] = 2048,
                ["temperature"] = 1.0,
                ["top_k"] = 1,
                ["top_p"] = 0.0,
                ["presence_penalty"] = 0.0,
                ["frequency_penalty"] = 0.0,
                ["stream"] = false,
                ["stop"] = stopSequences
            };

            if (!string.IsNullOrEmpty(_model))
            {
                requestData["model"] = _model;
            }

            var jsonContent = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(_apiUrl, content);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new Exception($"LlamaCpp chat API请求失败 HTTP {(int)response.StatusCode}: {errorContent}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var responseData = JObject.Parse(responseJson);

            // 解析 chat 响应: {"choices":[{"message":{"role":"assistant","content":"翻译结果"},...}]}
            var choice = responseData["choices"]?[0];
            var message = choice?["message"];
            var rawText = message?["content"]?.ToString()?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(rawText))
            {
                // 回退：尝试 text 字段
                rawText = choice?["text"]?.ToString()?.Trim() ?? string.Empty;
            }

            // 剥离模型可能回显的 prompt 内容
            rawText = StripPromptEcho(rawText, sourceCode, targetCode, text);

            return rawText;
        }

        /// <summary>
        /// 使用续写接口 (/v1/completions) 翻译
        /// 
        /// 直接构造与 rwkv_lightning 翻译端点完全等效的裸 prompt，
        /// 不经过任何聊天模板处理，确保 token 序列完全一致。
        /// 
        /// rwkv_lightning 的 create_translation_prompt 格式: "{src_lang}: {text}\n\n{tgt_lang}:"
        /// 推理参数: top_k=1, top_p=0.0 (确定性解码)
        /// </summary>
        private async Task<string> TranslateWithCompletionApiAsync(string text, string sourceLang, string targetLang)
        {
            var sourceCode = GetLanguageCode(sourceLang);
            var targetCode = GetLanguageCode(targetLang);

            // 构造翻译 prompt（使用可配置的语言分隔符）
            var prompt = CreateTranslationPrompt(sourceLang, targetLang, text);

            // 构建 stop 序列
            var stopSequences = BuildStopSequences(sourceLang);

            var requestData = new Dictionary<string, object>
            {
                ["prompt"] = prompt,
                ["max_tokens"] = 2048,
                ["temperature"] = 1.0,
                ["top_k"] = 1,
                ["top_p"] = 0.0,
                ["presence_penalty"] = 0.0,
                ["frequency_penalty"] = 0.0,
                ["stream"] = false,
                ["stop"] = stopSequences
            };

            if (!string.IsNullOrEmpty(_model))
            {
                requestData["model"] = _model;
            }

            var jsonContent = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync(_apiUrl, content);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                throw new Exception($"LlamaCpp completions API请求失败 HTTP {(int)response.StatusCode}: {errorContent}");
            }

            var responseJson = await response.Content.ReadAsStringAsync();
            var responseData = JObject.Parse(responseJson);

            // 解析续写响应: {"choices":[{"text":"翻译结果",...}]}
            var choice = responseData["choices"]?[0];
            var rawText = choice?["text"]?.ToString()?.Trim() ?? string.Empty;

            // 剥离模型可能回显的 prompt 内容
            rawText = StripPromptEcho(rawText, sourceCode, targetCode, text);

            return rawText;
        }

        private string ReplaceTermsInText(string text, Dictionary<string, string> terms, bool isCnToForeign)
        {
            if (string.IsNullOrWhiteSpace(text) || terms == null || !terms.Any())
                return text;

            var result = new StringBuilder(text);
            var sortedTerms = terms.OrderByDescending(t => t.Key.Length).ToList();

            foreach (var term in sortedTerms)
            {
                result.Replace(term.Key, term.Value);
            }

            return result.ToString();
        }

        #endregion

        #region ITranslator

        public override async Task<string> DetectLanguageAsync(string text)
        {
            return await base.DetectLanguageAsync(text);
        }

        public override async Task<string> ChatAsync(string question, string context = null)
        {
            try
            {
                var messages = new List<object>();

                if (!string.IsNullOrEmpty(context))
                {
                    messages.Add(new { role = "system", content = context });
                }

                messages.Add(new { role = "user", content = question });

                var requestData = new Dictionary<string, object>
                {
                    ["messages"] = messages.ToArray(),
                    ["max_tokens"] = 2048,
                    ["temperature"] = 0.7,
                    ["stream"] = false
                };

                if (!string.IsNullOrEmpty(_model))
                {
                    requestData["model"] = _model;
                }

                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                var chatUrl = _mode == LlamaCppMode.Chat ? _apiUrl : _apiUrl.Replace("/v1/completions", "/v1/chat/completions");
                var response = await _httpClient.PostAsync(chatUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    var responseJson = await response.Content.ReadAsStringAsync();
                    var responseData = JObject.Parse(responseJson);

                    var choice = responseData["choices"]?[0];
                    var message = choice?["message"];
                    var answer = message?["content"]?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(answer))
                    {
                        answer = choice?["text"]?.ToString()?.Trim() ?? string.Empty;
                    }

                    return answer;
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    throw new Exception($"HTTP {response.StatusCode}: {errorContent}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "LlamaCpp对话请求失败");
                throw;
            }
        }

        public override async Task<List<string>> GetAvailableModelsAsync()
        {
            try
            {
                var baseUrl = _apiUrl.Substring(0, _apiUrl.IndexOf("/", 8));
                var response = await _httpClient.GetAsync($"{baseUrl}/v1/models");

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    var data = JObject.Parse(json);
                    var models = data["data"]?
                        .Select(m => m["id"]?.ToString())
                        .Where(id => id != null)
                        .ToList() ?? new List<string>();

                    return models;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "获取LlamaCpp模型列表失败");
            }

            return new List<string>();
        }

        public override async Task<bool> TestConnectionAsync()
        {
            _logger.LogInformation($"开始测试LlamaCpp API连接，URL: {_apiUrl}");

            try
            {
                // 先尝试轻量级 /v1/models 端点
                var baseUrl = _apiUrl.Substring(0, _apiUrl.IndexOf("/", 8));
                try
                {
                    using var modelsCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));
                    var modelsResponse = await _httpClient.GetAsync($"{baseUrl}/v1/models", modelsCts.Token);
                    if (modelsResponse.IsSuccessStatusCode)
                    {
                        _logger.LogInformation($"LlamaCpp API连接测试成功（/v1/models）");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogDebug($"轻量连接测试失败，尝试完整测试: {ex.Message}");
                }

                // 完整测试：发送翻译请求
                if (_mode == LlamaCppMode.Chat)
                {
                    // 使用与翻译相同的 prompt 格式进行测试
                    var testData = new
                    {
                        messages = new[]
                        {
                            new { role = "user", content = "en: Hello\n\nzh-CN:" }
                        },
                        max_tokens = 10,
                        temperature = 1.0,
                        top_k = 1,
                        stream = false
                    };
                    var jsonContent = JsonConvert.SerializeObject(testData);
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                    var response = await _httpClient.PostAsync(_apiUrl, content, cts.Token);
                    return response.IsSuccessStatusCode;
                }
                else
                {
                    var testData = new
                    {
                        prompt = "en: Hello\n\nzh-CN:",
                        max_tokens = 10,
                        temperature = 1.0,
                        top_k = 1,
                        stream = false
                    };
                    var jsonContent = JsonConvert.SerializeObject(testData);
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                    using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(15));
                    var response = await _httpClient.PostAsync(_apiUrl, content, cts.Token);
                    return response.IsSuccessStatusCode;
                }
            }
            catch (TaskCanceledException ex)
            {
                _logger.LogWarning(ex, $"LlamaCpp API连接测试超时");
                return false;
            }
            catch (HttpRequestException ex)
            {
                _logger.LogWarning(ex, $"LlamaCpp API连接测试错误: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"LlamaCpp API连接测试异常: {ex.Message}");
                return false;
            }
        }

        #endregion

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _httpClient?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
