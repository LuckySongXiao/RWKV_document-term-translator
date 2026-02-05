using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace DocumentTranslator.Services.Translation.Translators
{
    /// <summary>
    /// RWKV翻译器，用于RWKV环境的API调用
    /// 支持的API接口：
    /// 1. /translate/v1/batch-translate - 批量翻译（非流式）
    /// 2. /v1/chat/completions - V1：基础批处理（支持流式和非流式）
    /// 3. /v2/chat/completions - V2：连续批处理（支持流式和非流式）
    /// 4. /v3/chat/completions - V3：异步高并发（VLLM风格，支持流式和非流式）
    /// </summary>
    public class RWKVTranslator : BaseTranslator
    {
        private readonly string _apiUrl;
        private readonly string _model;
        private readonly int _timeout;
        private readonly string _apiKey;
        private readonly bool _useBatchTranslate;
        private readonly HttpClient _httpClient;
        private static System.Threading.SemaphoreSlim _globalRequestSemaphore;
        private static readonly object _semaphoreLock = new object();
        private static int _maxConcurrency;
        private static Func<int> _getMaxConcurrency;

        public static int MaxConcurrency => _maxConcurrency;
        public static Func<int> GetMaxConcurrency
        {
            get => _getMaxConcurrency;
            set => _getMaxConcurrency = value;
        }

        private string Normalize(string url, bool batchMode)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentNullException(nameof(url));
            var u = url.TrimEnd('/');
            var lower = u.ToLowerInvariant();

            // 处理本地URL，确保正确格式
            if (lower.StartsWith("http://localhost:") || lower.StartsWith("http://127.0.0.1:"))
            {
                // 尝试获取WSL IP地址
                string wslIp = GetWslIpAddress();
                if (!string.IsNullOrEmpty(wslIp))
                {
                    // 使用WSL IP地址替换localhost
                    var port = u.Substring(u.LastIndexOf(":") + 1);
                    u = $"http://{wslIp}:{port}";
                    _logger.LogInformation($"检测到WSL环境，使用WSL IP地址: {u}");
                }

                // 清理URL中已有的API路径
                if (lower.Contains("/translate/v1/batch-translate"))
                {
                    u = u.Substring(0, lower.IndexOf("/translate/v1/batch-translate"));
                }
                else if (lower.Contains("/translate/v1"))
                {
                    u = u.Substring(0, lower.IndexOf("/translate/v1"));
                }

                if (batchMode)
                {
                    if (lower.EndsWith("batch-translate")) return u + "/batch-translate";
                    return u + "/translate/v1/batch-translate";
                }
                else
                {
                    // 非批量模式使用chat/completions
                if (lower.Contains("/chat/completions"))
                {
                    if (lower.Contains("/v3/chat/completions")) return u + "/v3/chat/completions";
                    if (lower.Contains("/v2/chat/completions")) return u + "/v2/chat/completions";
                    if (lower.Contains("/v1/chat/completions")) return u + "/v1/chat/completions";
                    return u + "/v3/chat/completions";
                }
                
                // 默认使用v3 chat completions
                return u + "/v3/chat/completions";
            }
            }

            // 非本地URL处理
            // 清理URL中已有的API路径
            if (lower.Contains("/translate/v1/batch-translate"))
            {
                u = u.Substring(0, lower.IndexOf("/translate/v1/batch-translate"));
            }
            else if (lower.Contains("/translate/v1"))
            {
                u = u.Substring(0, lower.IndexOf("/translate/v1"));
            }

            if (batchMode)
            {
                if (lower.Contains("/chat/completions"))
                {
                    var chatIndex = lower.IndexOf("/chat/completions");
                    u = u.Substring(0, chatIndex);
                }
                if (lower.EndsWith("batch-translate")) return u + "/batch-translate";
                if (lower.Contains("/translate/v1")) return u + "/translate/v1/batch-translate";
                return u + "/translate/v1/batch-translate";
            }
            else
            {
                // 非批量模式使用chat/completions
                if (lower.Contains("/chat/completions"))
                {
                    // 优先保留原有的V版本
                    if (lower.Contains("/v3/chat/completions")) return u + "/v3/chat/completions";
                    if (lower.Contains("/v2/chat/completions")) return u + "/v2/chat/completions";
                    if (lower.Contains("/v1/chat/completions")) return u + "/v1/chat/completions";
                    
                    // 如果没有明确V版本，但包含chat/completions，尝试升级到V3（如果原URL不含v1/v2）
                    // 这里我们保守一点，只在纯chat/completions时默认使用V3
                    return u + "/v3/chat/completions";
                }
                
                // 如果URL不包含chat/completions，尝试智能推断
                if (lower.Contains("/v3")) return u + "/v3/chat/completions";
                
                // 默认使用v3 chat completions以获得最佳并发性能
                return u + "/v3/chat/completions";
            }
        }

        private string GetWslIpAddress()
        {
            try
            {
                // 执行wsl命令获取IP地址
                var process = new System.Diagnostics.Process
                {
                    StartInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "wsl",
                        Arguments = "hostname -I",
                        RedirectStandardOutput = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                {
                    // 提取第一个IP地址
                    var ipAddresses = output.Trim().Split(' ');
                    if (ipAddresses.Length > 0)
                    {
                        return ipAddresses[0].Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogDebug($"获取WSL IP地址失败: {ex.Message}");
            }
            return null;
        }

        public RWKVTranslator(ILogger logger, string apiUrl, string model = "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118", 
            int timeout = 60, string apiKey = null, bool useBatchTranslate = false, int maxConcurrency = 30)
            : base(logger)
        {
            _model = model ?? throw new ArgumentNullException(nameof(model));
            _timeout = timeout;
            _apiKey = apiKey ?? string.Empty;
            _useBatchTranslate = useBatchTranslate;

            lock (_semaphoreLock)
            {
                if (_globalRequestSemaphore == null || _maxConcurrency != maxConcurrency)
                {
                    _maxConcurrency = maxConcurrency;
                    _globalRequestSemaphore = new System.Threading.SemaphoreSlim(maxConcurrency);
                    _logger.LogInformation($"RWKV全局请求信号量已更新为: {maxConcurrency}");
                }
            }

            _apiUrl = Normalize(apiUrl, _useBatchTranslate);

            var handler = new SocketsHttpHandler
            {
                PooledConnectionLifetime = TimeSpan.FromMinutes(5),
                PooledConnectionIdleTimeout = TimeSpan.FromMinutes(1),
                MaxConnectionsPerServer = 1000,
                EnableMultipleHttp2Connections = false,
                SslOptions = new System.Net.Security.SslClientAuthenticationOptions
                {
                    RemoteCertificateValidationCallback = (sender, certificate, chain, errors) => true,
                    EnabledSslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                    ClientCertificates = null,
                    CertificateRevocationCheckMode = System.Security.Cryptography.X509Certificates.X509RevocationMode.NoCheck
                },
                AutomaticDecompression = System.Net.DecompressionMethods.All,
                UseCookies = false,
                AllowAutoRedirect = true,
                MaxAutomaticRedirections = 5,
                UseProxy = false,
                // 增加本地连接的超时设置
                ConnectTimeout = TimeSpan.FromSeconds(30)
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(timeout)
            };
            
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "DocumentTranslator/1.0");
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
            _httpClient.DefaultRequestHeaders.Add("Connection", "keep-alive");
            
            if (!string.IsNullOrEmpty(_apiKey))
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            }

            _logger.LogInformation($"初始化RWKV翻译器，模型: {model}, API URL: {_apiUrl}, 批量翻译模式: {_useBatchTranslate}, 已提供API Key: {!string.IsNullOrEmpty(_apiKey)}, 翻译日志路径: {_translationLogPath}, {_translationLogPath2}");
        }

        public override async Task<string> TranslateAsync(string text, Dictionary<string, string> terminologyDict = null,
            string sourceLang = "zh", string targetLang = "en", string prompt = null, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var textForLog = originalText ?? text;

            // 预处理术语，避免在重试时重复处理
            string finalPromptText = text;
            if (terminologyDict != null && terminologyDict.Any())
            {
                finalPromptText = ReplaceTermsInText(text, terminologyDict, sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh"));
            }

            // 检查是否需要跳过翻译
            if (ShouldSkipTranslation(text))
                return text;

            if (!ShouldTranslateBasedOnLanguageRatio(text, sourceLang))
                return text;

            int maxRetries = 2; // 正常错误的最大重试次数
            int qualityRetryCount = 0; // 质量不佳错误的重试次数
            int retryDelayMs = 500; // 减少重试延迟时间
            bool isGatewayTimeout = false;
            bool hasQualityRetried = false; // 标记是否已经因为质量问题重试过

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    var currentMaxConcurrency = _getMaxConcurrency?.Invoke() ?? _maxConcurrency;

                    // 动态更新并发控制
                    if (currentMaxConcurrency != _maxConcurrency)
                    {
                        lock (_semaphoreLock)
                        {
                            if (currentMaxConcurrency != _maxConcurrency)
                            {
                                _maxConcurrency = currentMaxConcurrency;
                                _globalRequestSemaphore = new System.Threading.SemaphoreSlim(currentMaxConcurrency);
                                _logger.LogInformation($"RWKV全局请求信号量已动态更新为: {currentMaxConcurrency}");
                            }
                        }
                    }

                    // 确保信号量已初始化
                    if (_globalRequestSemaphore == null)
                    {
                        lock (_semaphoreLock)
                        {
                            if (_globalRequestSemaphore == null)
                            {
                                _maxConcurrency = currentMaxConcurrency;
                                _globalRequestSemaphore = new System.Threading.SemaphoreSlim(currentMaxConcurrency);
                                _logger.LogInformation($"RWKV全局请求信号量已初始化: {currentMaxConcurrency}");
                            }
                        }
                    }

                    // 使用局部变量捕获当前信号量，防止在等待期间信号量被替换（由于并发设置变更）
                    // 导致 release 错误的信号量对象
                    var currentSemaphore = _globalRequestSemaphore;
                    await currentSemaphore.WaitAsync();
                    try
                    {
                        string translation;

                        if (_useBatchTranslate)
                        {
                            translation = await TranslateWithBatchApiAsync(finalPromptText, sourceLang, targetLang);
                        }
                        else
                        {
                            translation = await TranslateWithChatApiAsync(finalPromptText, sourceLang, targetLang);
                        }

                        // 如果已经因为质量问题重试过，不再进行质量检测，直接返回译文
                        if (hasQualityRetried)
                        {
                            _logger.LogTrace($"质量重试后直接返回译文，原文: {text}, 译文: {translation}");
                            LogTranslation(textForLog, translation, sourceLang, targetLang);
                            return translation;
                        }

                        // 处理译文（包含质量检测）
                        var processedTranslation = ProcessAbnormalTranslation(text, translation, sourceLang, targetLang);

                        // 翻译成功，返回译文
                        _logger.LogTrace($"翻译成功，原文: {text}, 译文: {processedTranslation}");
                        LogTranslation(textForLog, processedTranslation, sourceLang, targetLang);

                        return processedTranslation;
                    }
                    finally
                    {
                        currentSemaphore.Release();

                    }
                }
                catch (Exception ex) when (attempt < maxRetries - 1)
                {
                    // 检查是否是质量不佳错误
                    bool isQualityError = ex.Message.Contains("译文质量不佳");
                    
                    if (isQualityError)
                    {
                        // 质量不佳错误，只允许重试一次
                        if (qualityRetryCount >= 1)
                        {
                            _logger.LogWarning($"译文质量不佳，已达到最大重试次数，不再重试 - 原文: {text}");
                            break;
                        }
                        
                        qualityRetryCount++;
                        hasQualityRetried = true; // 标记已经因为质量问题重试过
                        _logger.LogWarning($"RWKV翻译失败 (尝试 {qualityRetryCount}/1)，{retryDelayMs}ms后重试 - 原文: {text}，错误: {ex.Message}");
                    }
                    else if (ex is TaskCanceledException)
                    {
                        // 超时错误，保持正常重试
                        _logger.LogDebug($"RWKVAPI请求超时 (尝试 {attempt + 1}/{maxRetries})，{retryDelayMs}ms后重试 - 原文: {text}");
                    }
                    else if (ex is HttpRequestException)
                    {
                        // 网络错误，保持正常重试
                        isGatewayTimeout = ex.Message.Contains("GatewayTimeout") || ex.Message.Contains("504");
                        _logger.LogDebug($"无法连接到RWKVAPI (尝试 {attempt + 1}/{maxRetries})，{retryDelayMs}ms后重试 - 原文: {text}，错误: {ex.Message}");
                    }
                    else
                    {
                        // 其他错误，保持正常重试
                        _logger.LogDebug($"RWKV翻译失败 (尝试 {attempt + 1}/{maxRetries})，{retryDelayMs}ms后重试 - 原文: {text}，错误: {ex.Message}");
                    }
                    
                    await Task.Delay(retryDelayMs);
                    retryDelayMs += 500; // 固定增加延迟，而非指数增长
                }
            }

            // 最后一次尝试，使用原文作为回退
            try
            {
                await _globalRequestSemaphore.WaitAsync();
                try
                {
                    string translation;

                    if (_useBatchTranslate)
                    {
                        translation = await TranslateWithBatchApiAsync(finalPromptText, sourceLang, targetLang);
                    }
                    else
                    {
                        translation = await TranslateWithChatApiAsync(finalPromptText, sourceLang, targetLang);
                    }

                    // 如果已经因为质量问题重试过，不再进行质量检测，直接返回译文
                    if (hasQualityRetried)
                    {
                        _logger.LogTrace($"质量重试后直接返回译文，原文: {text}, 译文: {translation}");
                        return translation;
                    }

                    try
                    {
                        var processedTranslation = ProcessAbnormalTranslation(text, translation, sourceLang, targetLang);
                        return processedTranslation;
                    }
                    catch (Exception ex)
                    {
                        // 最后一次尝试如果质量仍然不佳，直接返回译文
                        _logger.LogWarning(ex, $"最后一次尝试译文质量不佳，直接返回译文 - 原文: {text}");
                        return translation;
                    }
                }
                finally
                {
                    _globalRequestSemaphore.Release();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"RWKV翻译最终失败，使用原文作为回退，原文: {text}");
                return text;
            }
        }


        private string GetLanguageCode(string lang)
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

        private async Task<string> TranslateWithBatchApiAsync(string promptText, string sourceLang, string targetLang)
        {
            var requestData = new
            {
                source_lang = GetLanguageCode(sourceLang),
                target_lang = GetLanguageCode(targetLang),
                text_list = new[] { promptText }
            };

            var jsonContent = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            try
            {
                var response = await _httpClient.PostAsync(_apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    var responseJson = await response.Content.ReadAsStringAsync();
                    var responseData = JsonConvert.DeserializeObject<dynamic>(responseJson);

                    var translationItem = responseData?.translations?[0];
                    var rawTranslation = string.Empty;

                    if (translationItem != null)
                    {
                        if (translationItem.text != null)
                        {
                            rawTranslation = translationItem.text.ToString()?.Trim() ?? string.Empty;
                        }
                        else
                        {
                            rawTranslation = translationItem.ToString()?.Trim() ?? string.Empty;
                        }
                    }

                    var translation = FilterOutput(rawTranslation, sourceLang, targetLang);

                    if (string.IsNullOrWhiteSpace(translation))
                    {
                        _logger.LogWarning("批量翻译结果为空，返回原始响应");
                        translation = rawTranslation;
                    }

                    _logger.LogDebug($"RWKV批量翻译成功，结果长度: {translation.Length}");
                    return translation;
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    _logger.LogError($"RWKV批量翻译API请求失败，状态码: {response.StatusCode}，响应内容: {errorContent}");
                    throw new Exception($"HTTP {response.StatusCode}: {errorContent}");
                }
            }
            catch (HttpRequestException ex)
            {
                _logger.LogError(ex, $"RWKV批量翻译API连接失败: {ex.Message}");
                throw new Exception($"RWKV批量翻译API连接失败: {ex.Message}", ex);
            }
            catch (TaskCanceledException ex)
            {
                _logger.LogError(ex, $"RWKV批量翻译API请求超时: {ex.Message}");
                throw new Exception($"RWKV批量翻译API请求超时: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"RWKV批量翻译API请求异常: {ex.Message}");
                throw new Exception($"RWKV批量翻译API请求异常: {ex.Message}", ex);
            }
        }

        private async Task<string> TranslateWithChatApiAsync(string promptText, string sourceLang, string targetLang)
        {
            var sourceLangName = GetLanguageName(sourceLang);
            var targetLangName = GetLanguageName(targetLang);

            var requestData = BuildChatApiRequestData(promptText, sourceLangName, targetLangName);

            var jsonContent = JsonConvert.SerializeObject(requestData);
            var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            try
            {
                var response = await _httpClient.PostAsync(_apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    var responseJson = await response.Content.ReadAsStringAsync();
                    var responseData = JsonConvert.DeserializeObject<dynamic>(responseJson);

                    var rawTranslation = string.Empty;
                    var choiceItem = responseData?.choices?[0];

                    if (choiceItem != null)
                    {
                        if (choiceItem.message != null && choiceItem.message.content != null)
                        {
                            rawTranslation = choiceItem.message.content.ToString()?.Trim() ?? string.Empty;
                        }
                        else if (choiceItem.text != null)
                        {
                            rawTranslation = choiceItem.text.ToString()?.Trim() ?? string.Empty;
                        }
                        else
                        {
                            rawTranslation = choiceItem.ToString()?.Trim() ?? string.Empty;
                        }
                    }

                    var translation = FilterOutput(rawTranslation, sourceLang, targetLang);

                    if (string.IsNullOrWhiteSpace(translation))
                    {
                        _logger.LogWarning("翻译结果为空，返回原始响应");
                        translation = rawTranslation;
                    }

                    _logger.LogDebug($"RWKV翻译成功，结果长度: {translation.Length}");
                    return translation;
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    _logger.LogError($"RWKVAPI请求失败，状态码: {response.StatusCode}，响应内容: {errorContent}");
                    throw new Exception($"HTTP {response.StatusCode}: {errorContent}");
                }
            }
            catch (HttpRequestException ex)
            {
                _logger.LogError(ex, $"RWKVAPI连接失败: {ex.Message}");
                throw new Exception($"RWKVAPI连接失败: {ex.Message}", ex);
            }
            catch (TaskCanceledException ex)
            {
                _logger.LogError(ex, $"RWKVAPI请求超时: {ex.Message}");
                throw new Exception($"RWKVAPI请求超时: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"RWKVAPI请求异常: {ex.Message}");
                throw new Exception($"RWKVAPI请求异常: {ex.Message}", ex);
            }
        }

        private object BuildChatApiRequestData(string promptText, string sourceLangName, string targetLangName)
        {
            // 根据不同的API版本调整请求格式
            var isV1OrLater = (_apiUrl?.Contains("/v1/") ?? false) || (_apiUrl?.Contains("/v2/") ?? false) || (_apiUrl?.Contains("/v3/") ?? false);
            var isLocalhost = (_apiUrl?.ToLowerInvariant().StartsWith("http://localhost:") ?? false) || (_apiUrl?.ToLowerInvariant().StartsWith("http://127.0.0.1:") ?? false);
            
            if (isV1OrLater || isLocalhost)
            {
                // V1/V2/V3 API格式（RWKV专用格式）或本地模型
                var baseRequest = new Dictionary<string, object>
                {
                    ["contents"] = new[] { promptText },
                    ["max_tokens"] = 1024,
                    ["temperature"] = 1.0,
                    ["stream"] = false,
                    // 添加必要的字段以匹配RWKV API格式
                    ["stop_tokens"] = new[] { 0, 261, 24281 }
                };

                // 添加password字段（RWKV API需要）
                if (!string.IsNullOrEmpty(_apiKey))
                {
                    baseRequest["password"] = _apiKey;
                }

                // 为本地模型添加额外的必要字段
                if (isLocalhost)
                {
                    baseRequest["top_k"] = 1;
                    baseRequest["top_p"] = 0.3;
                }

                return baseRequest;
            }
            else
            {
                // 旧版API格式
                var baseRequest = new Dictionary<string, object>
                {
                    ["contents"] = new[] { promptText },
                    ["max_tokens"] = 1024,
                    ["temperature"] = 1.0,
                    ["stream"] = false,
                    ["stop_tokens"] = new[] { 0, 261, 24281 }
                };

                if (!string.IsNullOrEmpty(_apiKey))
                {
                    baseRequest["password"] = _apiKey;
                }

                return baseRequest;
            }
        }

        public override async Task<string> ChatAsync(string question, string context = null)
        {
            if (string.IsNullOrWhiteSpace(question))
                return string.Empty;

            try
            {
                var contents = new List<string>();

                if (!string.IsNullOrWhiteSpace(context))
                {
                    contents.Add(context);
                }

                contents.Add($"Question: {question}\nPlease provide a helpful and detailed answer.");

                var requestData = new Dictionary<string, object>
                {
                    ["contents"] = contents.ToArray(),
                    ["max_tokens"] = 2048,
                    ["temperature"] = 0.8,
                    ["stream"] = false,
                    ["stop_tokens"] = new[] { 0, 261, 24281 }
                };

                if (!string.IsNullOrEmpty(_apiKey))
                {
                    requestData["password"] = _apiKey;
                }

                // 为本地模型添加额外的必要字段
                var isLocalhost = (_apiUrl?.ToLowerInvariant().StartsWith("http://localhost:") ?? false) || (_apiUrl?.ToLowerInvariant().StartsWith("http://127.0.0.1:") ?? false);
                if (isLocalhost)
                {
                    requestData["top_k"] = 1;
                    requestData["top_p"] = 0.3;
                }

                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                try
                {
                    var chatUrl = Normalize(_apiUrl, false);
                    _logger.LogInformation($"RWKV对话请求，使用URL: {chatUrl}");

                    var response = await _httpClient.PostAsync(chatUrl, content);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseJson = await response.Content.ReadAsStringAsync();
                        var responseData = JsonConvert.DeserializeObject<dynamic>(responseJson);

                        var answer = string.Empty;
                        var choiceItem = responseData?.choices?[0];

                        if (choiceItem != null)
                        {
                            if (choiceItem.message != null && choiceItem.message.content != null)
                            {
                                answer = choiceItem.message.content.ToString()?.Trim() ?? string.Empty;
                            }
                            else if (choiceItem.text != null)
                            {
                                answer = choiceItem.text.ToString()?.Trim() ?? string.Empty;
                            }
                            else
                            {
                                answer = choiceItem.ToString()?.Trim() ?? string.Empty;
                            }
                        }

                        _logger.LogInformation($"RWKV对话成功，结果长度: {answer.Length}");
                        return answer;
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        _logger.LogError($"RWKV对话请求失败，状态码: {response.StatusCode}，响应内容: {errorContent}");
                        throw new Exception($"HTTP {response.StatusCode}: {errorContent}");
                    }
                }
                catch (HttpRequestException ex)
                {
                    _logger.LogError(ex, $"RWKV对话连接失败: {ex.Message}");
                    throw new Exception($"RWKV对话连接失败: {ex.Message}", ex);
                }
                catch (TaskCanceledException ex)
                {
                    _logger.LogError(ex, $"RWKV对话请求超时: {ex.Message}");
                    throw new Exception($"RWKV对话请求超时: {ex.Message}", ex);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"RWKV对话发生异常: {ex.Message}");
                throw new Exception($"RWKV对话发生异常: {ex.Message}", ex);
            }
        }

        public override async Task<List<string>> GetAvailableModelsAsync()
        {
            return await Task.FromResult(new List<string>
            {
                "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118",
                "rwkv7-g1d-2.9b-20260131-ctx8192",
                "rwkv7-g1c-13.3b-20251231-ctx8192"
            });
        }

        public override async Task<bool> TestConnectionAsync()
        {
            _logger.LogInformation($"开始测试RWKV API连接，URL: {_apiUrl}");

            try
            {
                object testData;
                var isLocalhost = (_apiUrl?.ToLowerInvariant().StartsWith("http://localhost:") ?? false) || (_apiUrl?.ToLowerInvariant().StartsWith("http://127.0.0.1:") ?? false);

                if (_useBatchTranslate)
                {
                    testData = new
                    {
                        source_lang = "en",
                        target_lang = "zh-CN",
                        text_list = new[] { "Hello" }
                    };
                }
                else
                {
                    var isV1OrLater = (_apiUrl?.Contains("/v1/") ?? false) || (_apiUrl?.Contains("/v2/") ?? false) || (_apiUrl?.Contains("/v3/") ?? false);

                    if (isV1OrLater || isLocalhost)
                    {
                        testData = new Dictionary<string, object>
                        {
                            ["contents"] = new[] { "你好" },
                            ["max_tokens"] = 10,
                            ["temperature"] = 1.0,
                            ["stream"] = false,
                            ["stop_tokens"] = new[] { 0, 261, 24281 }
                        };

                        // 为本地模型添加额外的必要字段
                        if (isLocalhost)
                        {
                            var dict = (Dictionary<string, object>)testData;
                            dict["top_k"] = 1;
                            dict["top_p"] = 0.3;
                        }
                    }
                    else
                    {
                        testData = new Dictionary<string, object>
                        {
                            ["contents"] = new[] { "你好" },
                            ["max_tokens"] = 10,
                            ["temperature"] = 1.0,
                            ["stream"] = false,
                            ["stop_tokens"] = new[] { 0, 261, 24281 }
                        };
                    }

                    if (!string.IsNullOrEmpty(_apiKey))
                    {
                        var dict = (Dictionary<string, object>)testData;
                        dict["password"] = _apiKey;
                    }
                }

                var jsonContent = JsonConvert.SerializeObject(testData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // 为本地连接增加超时时间
                var timeoutSeconds = isLocalhost ? 15 : 5;
                using var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
                var response = await _httpClient.PostAsync(_apiUrl, content, cts.Token);

                if (response.IsSuccessStatusCode)
                {
                    _logger.LogInformation($"RWKVAPI连接测试成功，使用URL: {_apiUrl}");
                    return true;
                }
                else
                {
                    var errorContent = await response.Content.ReadAsStringAsync();
                    _logger.LogWarning($"RWKVAPI连接测试失败，状态码: {response.StatusCode}，响应内容: {errorContent}");
                    return false;
                }
            }
            catch (TaskCanceledException ex)
            {
                _logger.LogWarning(ex, $"RWKVAPI连接测试超时: {ex.Message}");
                return false;
            }
            catch (HttpRequestException ex)
            {
                _logger.LogWarning(ex, $"RWKVAPI连接测试错误: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"RWKVAPI连接测试异常: {ex.Message}");
                return false;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _httpClient?.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool IsChinese(char c)
        {
            return c >= 0x4E00 && c <= 0x9FFF;
        }

        private string ReplaceTermsInText(string text, Dictionary<string, string> terms, bool isCnToForeign)
        {
            if (string.IsNullOrWhiteSpace(text) || terms == null || !terms.Any())
            {
                return text;
            }

            // 快速检查文本是否可能包含任何术语
            bool ContainsAnyTerm(string textToCheck, Dictionary<string, string> termDict)
            {
                foreach (var term in termDict.Keys)
                {
                    if (textToCheck.Contains(term, isCnToForeign ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }

            // 如果文本不包含任何术语，直接返回原文
            if (!ContainsAnyTerm(text, terms))
            {
                return text;
            }

            // 按术语长度降序排序，优先替换长术语
            var sorted = terms.Where(kvp => kvp.Key.Length <= text.Length).OrderByDescending(kvp => kvp.Key.Length).ToList();
            if (sorted.Count == 0)
            {
                return text;
            }

            var result = text;
            var replacedTerms = new List<string>();
            int replacedCount = 0;

            // 只在调试级别记录开始信息
            _logger?.LogDebug($"开始术语预处理，共 {sorted.Count} 个候选术语，翻译方向: {(isCnToForeign ? "中文→外文" : "外文→中文")}");

            foreach (var kv in sorted)
            {
                var src = kv.Key;
                var dst = kv.Value;
                if (string.IsNullOrWhiteSpace(src))
                    continue;

                // 跳过长度超过当前结果文本的术语
                if (src.Length > result.Length)
                    continue;

                try
                {
                    // 先进行简单的字符串包含检查，避免不必要的正则匹配
                    if (!result.Contains(src, isCnToForeign ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string pattern;
                    RegexOptions options = RegexOptions.None;

                    if (isCnToForeign)
                    {
                        // 使用边界限制，避免部分匹配 (与TermExtractor保持一致)
                        string srcEscaped = Regex.Escape(src);
                        
                        // 检查起始字符是否为ASCII字母/数字
                        bool startIsAlnum = src.Length > 0 && ((src[0] >= 'a' && src[0] <= 'z') || (src[0] >= 'A' && src[0] <= 'Z') || (src[0] >= '0' && src[0] <= '9'));
                        // 检查结束字符是否为ASCII字母/数字
                        bool endIsAlnum = src.Length > 0 && ((src[^1] >= 'a' && src[^1] <= 'z') || (src[^1] >= 'A' && src[^1] <= 'Z') || (src[^1] >= '0' && src[^1] <= '9'));

                        string prefix = startIsAlnum ? "(?<![a-zA-Z0-9])" : "";
                        string suffix = endIsAlnum ? "(?![a-zA-Z0-9])" : "";
                        
                        pattern = $"{prefix}{srcEscaped}{suffix}";
                    }
                    else
                    {
                        pattern = $"\\b{Regex.Escape(src)}\\b";
                        options = RegexOptions.IgnoreCase;
                    }

                    // 执行正则替换
                    var originalLength = result.Length;
                    result = Regex.Replace(result, pattern, dst, options);
                    
                    // 检查是否有实际替换
                    if (result.Length != originalLength)
                    {
                        replacedCount++;
                        
                        // 只记录前5个替换的术语，避免日志过多
                        if (replacedTerms.Count < 5)
                        {
                            replacedTerms.Add($"{src}→{dst}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"术语预替换异常: {src} -> {dst}");
                }
            }

            // 只在有替换时记录信息
            if (replacedCount > 0)
            {
                var logMessage = $"术语替换完成，共替换 {replacedCount} 个术语";
                if (replacedTerms.Any())
                {
                    if (replacedCount > replacedTerms.Count)
                    {
                        logMessage += $"，前 {replacedTerms.Count} 个: {string.Join(", ", replacedTerms)}";
                    } else {
                        logMessage += $": {string.Join(", ", replacedTerms)}";
                    }
                }
                _logger?.LogInformation(logMessage);
            }

            // 只在调试级别记录完成信息
            _logger?.LogDebug($"术语预处理完成，原文长度: {text.Length}，处理后长度: {result.Length}");

            return result;
        }

        private bool ShouldSkipTranslation(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return true;

            var trimmed = text.Trim();

            if (trimmed.Length == 0)
                return true;

            var digitCount = 0;
            var letterCount = 0;
            var chineseCount = 0;
            var otherCount = 0;

            foreach (char c in trimmed)
            {
                if (char.IsDigit(c))
                {
                    digitCount++;
                }
                else if (char.IsLetter(c))
                {
                    if (IsChinese(c))
                    {
                        chineseCount++;
                    }
                    else
                    {
                        letterCount++;
                    }
                }
                else
                {
                    otherCount++;
                }
            }

            var totalChars = trimmed.Length;

            if (digitCount == totalChars)
            {
                _logger.LogInformation($"检测到纯数字文本: '{trimmed}'");
                return true;
            }

            if (digitCount + otherCount == totalChars)
            {
                _logger.LogInformation($"检测到纯数字/编码文本: '{trimmed}'");
                return true;
            }

            if (letterCount == totalChars && IsPureCode(trimmed))
            {
                _logger.LogInformation($"检测到纯编码文本: '{trimmed}'");
                return true;
            }

            return false;
        }

        private bool IsPureCode(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.Trim();

            if (trimmed.Length < 2)
                return false;

            var hasMultipleConsecutiveLetters = false;
            var consecutiveLetterCount = 0;
            var maxConsecutiveLetters = 0;

            foreach (char c in trimmed)
            {
                if (char.IsLetter(c) && !IsChinese(c))
                {
                    consecutiveLetterCount++;
                    maxConsecutiveLetters = Math.Max(maxConsecutiveLetters, consecutiveLetterCount);
                }
                else
                {
                    consecutiveLetterCount = 0;
                }
            }

            hasMultipleConsecutiveLetters = maxConsecutiveLetters >= 2;

            var hasSpecialChars = false;
            foreach (char c in trimmed)
            {
                if (c == '_' || c == '-' || c == '.' || c == ':' || c == '/' || c == '\\')
                {
                    hasSpecialChars = true;
                    break;
                }
            }

            var hasDigits = trimmed.Any(char.IsDigit);

            return hasMultipleConsecutiveLetters && (hasSpecialChars || hasDigits);
        }

        private bool ShouldTranslateBasedOnLanguageRatio(string text, string sourceLang)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.Trim();

            if (trimmed.Length == 0)
                return false;

            var isZhSource = sourceLang == "zh" || sourceLang == "zh-cn" || sourceLang == "auto";

            var chineseCount = 0;
            var nonChineseCount = 0;

            foreach (char c in trimmed)
            {
                if (IsChinese(c))
                {
                    chineseCount++;
                }
                else if (char.IsLetter(c) || char.IsDigit(c))
                {
                    nonChineseCount++;
                }
            }

            var totalChars = chineseCount + nonChineseCount;

            if (totalChars == 0)
                return false;

            if (isZhSource)
            {
                var chineseRatio = (double)chineseCount / totalChars;
                return chineseRatio >= 0.2;
            }
            else
            {
                var nonChineseRatio = (double)nonChineseCount / totalChars;
                return nonChineseRatio >= 0.2;
            }
        }
    }
}
