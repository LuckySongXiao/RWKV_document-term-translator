using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 翻译器基础抽象类，提供通用的翻译功能和输出过滤
    /// </summary>
    public abstract class BaseTranslator : ITranslator, IDisposable
    {
        protected readonly ILogger _logger;
        protected string _translationLogPath;
        protected string _translationLogPath2;
        protected string _preprocessedLogPath;
        protected string _preprocessedLogPath2;
        protected string _duplicateLogPath;
        protected string _abnormalLogPath;
        protected string _failureLogPath;
        protected string _timeoutLogPath;
        private static readonly object _globalLogLock = new object();
        protected string _batchTimestamp;
        
        protected BaseTranslator(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _batchTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            UpdateLogPaths();
            EnsureLogDirectoryExists();
        }

        private void UpdateLogPaths()
        {
            var baseDir = PathHelper.GetSafeBaseDirectory();
            var logsDir = Path.Combine(baseDir, "logs");
            
            _translationLogPath = Path.Combine(logsDir, "Success", $"translation_log_{_batchTimestamp}.jsonl");
            _translationLogPath2 = Path.Combine(logsDir, "Success_Bilingual", $"translation_log_{_batchTimestamp}.jsonl");
            
            _preprocessedLogPath = Path.Combine(logsDir, "Preprocessed", $"preprocessed_log_{_batchTimestamp}.jsonl");
            _preprocessedLogPath2 = Path.Combine(logsDir, "Preprocessed_Bilingual", $"preprocessed_log_{_batchTimestamp}.jsonl");
            
            _duplicateLogPath = Path.Combine(logsDir, "Duplicate", $"duplicate_log_{_batchTimestamp}.txt");
            _abnormalLogPath = Path.Combine(logsDir, "Abnormal", $"abnormal_log_{_batchTimestamp}.txt");
            _failureLogPath = Path.Combine(logsDir, "Failure", $"failure_log_{_batchTimestamp}.txt");
            _timeoutLogPath = Path.Combine(logsDir, "TimeOut", $"timeout_log_{_batchTimestamp}.txt");
        }

        public void StartNewBatch()
        {
            _batchTimestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            UpdateLogPaths();
            EnsureLogDirectoryExists();
            _logger.LogInformation($"开始新的翻译批次，时间戳: {_batchTimestamp}");
        }

        public void SetBatchTimestamp(string timestamp)
        {
            _batchTimestamp = timestamp;
            UpdateLogPaths();
            EnsureLogDirectoryExists();
            _logger.LogInformation($"设置批次时间戳: {_batchTimestamp}");
        }

        public string GetBatchTimestamp()
        {
            return _batchTimestamp;
        }

        /// <summary>
        /// 翻译文本的抽象方法，由具体翻译器实现
        /// </summary>
        public abstract Task<string> TranslateAsync(string text, Dictionary<string, string> terminologyDict = null,
            string sourceLang = "zh", string targetLang = "en", string prompt = null, string originalText = null);

        /// <summary>
        /// 检测文本语言的默认实现（基于Unicode范围）
        /// </summary>
        public virtual Task<string> DetectLanguageAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return Task.FromResult("other");

            var languageDefinitions = new Dictionary<string, (string Name, string UnicodeRange)>
            {
                ["zh"] = ("中文", @"\u4e00-\u9fff"),
                ["en"] = ("英文", @"[a-zA-Z]"),
                ["ja"] = ("日语", @"\u3040-\u309f|\u30a0-\u30ff|\u4e00-\u9fff"),
                ["ko"] = ("韩语", @"\uac00-\ud7af|\u1100-\u11ff"),
                ["fr"] = ("法语", @"[a-zA-Zàâäéèêëïîôùûüÿç]"),
                ["de"] = ("德语", @"[a-zA-ZäöüßÄÖÜ]"),
                ["es"] = ("西班牙语", @"[a-zA-Záéíóúüñ¿¡]"),
                ["it"] = ("意大利语", @"[a-zA-Zàèéìòù]"),
                ["ru"] = ("俄语", @"\u0400-\u04ff"),
                ["pt"] = ("葡萄牙语", @"[a-zA-Zãõáàâéêíóôúç]"),
                ["vi"] = ("越南语", @"\u1ea0-\u1ef9|a-zA-Z]")
            };

            int totalChars = 0;
            var languageCounts = new Dictionary<string, int>();

            foreach (var lang in languageDefinitions.Keys)
            {
                languageCounts[lang] = 0;
            }

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c) || char.IsPunctuation(c))
                    continue;

                totalChars++;

                foreach (var lang in languageDefinitions.Keys)
                {
                    var (_, unicodeRange) = languageDefinitions[lang];
                    if (Regex.IsMatch(c.ToString(), $"[{unicodeRange}]"))
                    {
                        languageCounts[lang]++;
                        break;
                    }
                }
            }

            if (totalChars == 0)
                return Task.FromResult("other");

            var result = languageCounts
                .Where(kv => kv.Value > 0)
                .OrderByDescending(kv => kv.Value)
                .FirstOrDefault();

            return Task.FromResult(result.Key ?? "other");
        }

        /// <summary>
        /// 对话接口的默认实现，使用翻译接口实现对话功能
        /// </summary>
        public virtual async Task<string> ChatAsync(string question, string context = null)
        {
            if (string.IsNullOrWhiteSpace(question))
                return string.Empty;

            var promptBuilder = new System.Text.StringBuilder();

            if (!string.IsNullOrWhiteSpace(context))
            {
                promptBuilder.AppendLine("以下是历史对话记录：");
                promptBuilder.AppendLine(context);
                promptBuilder.AppendLine();
            }

            promptBuilder.AppendLine("请回答以下问题：");
            promptBuilder.AppendLine(question);

            var prompt = promptBuilder.ToString();

            return await TranslateAsync(prompt, null, "zh", "zh", prompt);
        }

        /// <summary>
        /// 获取可用模型列表的抽象方法，由具体翻译器实现
        /// </summary>
        public abstract Task<List<string>> GetAvailableModelsAsync();

        /// <summary>
        /// 测试连接的抽象方法，由具体翻译器实现
        /// </summary>
        public abstract Task<bool> TestConnectionAsync();

        /// <summary>
        /// 过滤模型输出，去除思维链、不必要的标记和提示性文本
        /// </summary>
        /// <param name="text">模型输出的文本</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        /// <returns>过滤后的文本</returns>
        protected virtual string FilterOutput(string text, string sourceLang = "zh", string targetLang = "en")
        {
            // 如果输入为空，直接返回
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // 如果存在<think>标签，提取非思维链部分
            if (text.Contains("<think>"))
            {
                // 分割所有的<think>块
                var parts = text.Split(new[] { "<think>" }, StringSplitOptions.None);
                // 获取最后一个非思维链的内容
                var finalText = parts.Last().Trim();
                // 如果最后一部分还包含</think>，取其后面的内容
                if (finalText.Contains("</think>"))
                {
                    finalText = finalText.Split(new[] { "</think>" }, StringSplitOptions.None).Last().Trim();
                }
                text = finalText;
            }

            // 去除常见的标记前缀
            text = text.Trim();

            // 定义需要过滤的提示性文本列表
            var promptTexts = new[]
            {
                "请提供具体的英文文本，以便进行翻译",
                "请提供具体术语内容，以便进行翻译",
                "请提供具体需要翻译的英文文本，以便我进行准确的翻译",
                "请提供英文文本",
                "请提供需要翻译的文本",
                "请提供具体的文本内容",
                "请提供原文",
                "请提供需要翻译的内容",
                "以下是翻译结果",
                "以下是我的翻译",
                "以下是译文",
                "这是翻译结果",
                "这是我的翻译",
                "翻译如下",
                "翻译结果如下",
                "译文如下"
            };

            // 检查文本是否只包含提示性文本
            foreach (var promptText in promptTexts)
            {
                if (text.Trim() == promptText || 
                    text.Trim() == promptText + "。" || 
                    text.Trim() == promptText + ":")
                {
                    return string.Empty; // 如果文本只包含提示性文本，则返回空字符串
                }
            }

            // 去除"原文："和"译文："等标记
            var lines = text.Split('\n');
            var filteredLines = new List<string>();

            bool hasOriginalTranslationPair = lines.Any(line => line.Contains("原文：") && line.Contains("译文："));

            if (hasOriginalTranslationPair && targetLang == "zh")
            {
                var resultText = string.Empty;
                foreach (var line in lines)
                {
                    if (line.Contains("译文："))
                    {
                        var translationPart = line.Split(new[] { "译文：" }, 2, StringSplitOptions.None)[1].Trim();
                        if (!string.IsNullOrEmpty(translationPart))
                        {
                            resultText += translationPart + "\n";
                        }
                    }
                }
                if (!string.IsNullOrWhiteSpace(resultText))
                {
                    return resultText.Trim();
                }
            }

            foreach (var line in lines)
            {
                var lineStripped = line.Trim();

                if (lineStripped == "原文：" || lineStripped == "译文：")
                    continue;

                var processedLine = line;

                if (lineStripped.StartsWith("原文："))
                    continue;
                if (lineStripped.StartsWith("译文："))
                {
                    var index = line.IndexOf("译文：");
                    if (index >= 0)
                    {
                        processedLine = line.Substring(0, index) + line.Substring(index + "译文：".Length);
                        processedLine = processedLine.Trim();
                    }
                }

                if (lineStripped.EndsWith("原文："))
                    processedLine = line.Replace("原文：", "").Trim();
                if (lineStripped.EndsWith("译文："))
                    processedLine = line.Replace("译文：", "").Trim();

                bool isPromptText = promptTexts.Any(promptText => 
                    lineStripped.Contains(promptText) || lineStripped.StartsWith(promptText));

                if (isPromptText)
                    continue;

                filteredLines.Add(processedLine);
            }

            // 重新组合文本
            var filteredText = string.Join("\n", filteredLines);

            // 剥离输出开头可能残留的语言代码前缀（如 "en:", "zh-CN:", "ja:" 等）
            var langCodePrefixes = new[] { "zh-CN:", "zh-TW:", "zh:", "en:", "ja:", "ko:", "fr:", "de:", "es:", "ru:", "vi:", "pt:", "it:", "nl:", "ar:", "th:", "id:", "ms:", "tr:" };
            var trimmedFiltered = filteredText.TrimStart();
            foreach (var prefix in langCodePrefixes)
            {
                if (trimmedFiltered.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    var afterPrefix = trimmedFiltered.Substring(prefix.Length).TrimStart(' ');
                    if (!string.IsNullOrWhiteSpace(afterPrefix))
                    {
                        filteredText = afterPrefix;
                    }
                    break;
                }
            }

            // 外语→中文翻译模式下的额外过滤
            if (targetLang == "zh")
            {
                // 去除常见的前缀短语
                var prefixesToRemove = new[]
                {
                    "以下是翻译：",
                    "翻译结果：",
                    "翻译如下：",
                    "译文如下：",
                    "翻译：",
                    "译文：",
                    "这是翻译：",
                    "这是译文：",
                    "以下是中文翻译：",
                    "中文翻译：",
                    "中文译文："
                };

                foreach (var prefix in prefixesToRemove)
                {
                    if (filteredText.StartsWith(prefix))
                    {
                        filteredText = filteredText.Substring(prefix.Length).Trim();
                    }
                }

                // 去除常见的后缀短语
                var suffixesToRemove = new[]
                {
                    "（以上是翻译结果）",
                    "（这是翻译结果）",
                    "（翻译完成）",
                    "（完成翻译）",
                    "（以上为翻译）",
                    "（以上为译文）"
                };

                foreach (var suffix in suffixesToRemove)
                {
                    if (filteredText.EndsWith(suffix))
                    {
                        filteredText = filteredText.Substring(0, filteredText.Length - suffix.Length).Trim();
                    }
                }
            }

            // 如果过滤后的文本为空，但原文不为空，则可能过滤过度
            // 在这种情况下，返回原始文本
            if (string.IsNullOrWhiteSpace(filteredText) && !string.IsNullOrWhiteSpace(text))
            {
                return text.Trim();
            }

            return filteredText;
        }

        /// <summary>
        /// 根据语言代码获取语言名称
        /// </summary>
        /// <param name="langCode">语言代码</param>
        /// <returns>语言名称</returns>
        protected virtual string GetLanguageName(string langCode)
        {
            if (string.IsNullOrWhiteSpace(langCode))
                return "Unknown";

            var languageMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"zh", "Chinese"},
                {"cht", "Traditional Chinese"},
                {"en", "English"},
                {"ja", "Japanese"},
                {"ko", "Korean"},
                {"fr", "French"},
                {"de", "German"},
                {"es", "Spanish"},
                {"it", "Italian"},
                {"ru", "Russian"},
                {"pt", "Portuguese"},
                {"nl", "Dutch"},
                {"ar", "Arabic"},
                {"th", "Thai"},
                {"vi", "Vietnamese"},
                {"auto", "Auto"},
                {"id", "Indonesian"},
                {"ms", "Malay"},
                {"tr", "Turkish"}
            };

            return languageMap.TryGetValue(langCode, out var name) ? name : "Unknown";
        }

        /// <summary>
        /// 确保日志目录存在
        /// </summary>
        protected void EnsureLogDirectoryExists()
        {
            try
            {
                var logDirectory = Path.GetDirectoryName(_translationLogPath);
                if (!string.IsNullOrEmpty(logDirectory) && !Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                    _logger.LogInformation($"创建日志目录: {logDirectory}");
                }

                var logDirectory2 = Path.GetDirectoryName(_translationLogPath2);
                if (!string.IsNullOrEmpty(logDirectory2) && !Directory.Exists(logDirectory2))
                {
                    Directory.CreateDirectory(logDirectory2);
                    _logger.LogInformation($"创建日志目录: {logDirectory2}");
                }

                var preprocessedDirectory = Path.GetDirectoryName(_preprocessedLogPath);
                if (!string.IsNullOrEmpty(preprocessedDirectory) && !Directory.Exists(preprocessedDirectory))
                {
                    Directory.CreateDirectory(preprocessedDirectory);
                    _logger.LogInformation($"创建日志目录: {preprocessedDirectory}");
                }

                var preprocessedDirectory2 = Path.GetDirectoryName(_preprocessedLogPath2);
                if (!string.IsNullOrEmpty(preprocessedDirectory2) && !Directory.Exists(preprocessedDirectory2))
                {
                    Directory.CreateDirectory(preprocessedDirectory2);
                    _logger.LogInformation($"创建日志目录: {preprocessedDirectory2}");
                }

                var duplicateDirectory = Path.GetDirectoryName(_duplicateLogPath);
                if (!string.IsNullOrEmpty(duplicateDirectory) && !Directory.Exists(duplicateDirectory))
                {
                    Directory.CreateDirectory(duplicateDirectory);
                    _logger.LogInformation($"创建日志目录: {duplicateDirectory}");
                }

                var abnormalDirectory = Path.GetDirectoryName(_abnormalLogPath);
                if (!string.IsNullOrEmpty(abnormalDirectory) && !Directory.Exists(abnormalDirectory))
                {
                    Directory.CreateDirectory(abnormalDirectory);
                    _logger.LogInformation($"创建日志目录: {abnormalDirectory}");
                }

                var failureDirectory = Path.GetDirectoryName(_failureLogPath);
                if (!string.IsNullOrEmpty(failureDirectory) && !Directory.Exists(failureDirectory))
                {
                    Directory.CreateDirectory(failureDirectory);
                    _logger.LogInformation($"创建日志目录: {failureDirectory}");
                }

                var timeoutDirectory = Path.GetDirectoryName(_timeoutLogPath);
                if (!string.IsNullOrEmpty(timeoutDirectory) && !Directory.Exists(timeoutDirectory))
                {
                    Directory.CreateDirectory(timeoutDirectory);
                    _logger.LogInformation($"创建日志目录: {timeoutDirectory}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"创建日志目录失败");
            }
        }

        /// <summary>
        /// 记录翻译日志到JSON文件
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        protected void LogTranslation(string originalText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return;

            try
            {
                var sourceLangName = GetLanguageName(sourceLang);
                var targetLangName = GetLanguageName(targetLang);

                var logEntry1 = new
                {
                    text = $"{sourceLangName}: {originalText}\n\n{targetLangName}: {translatedText}"
                };

                var logEntry2 = new
                {
                    text = $"{targetLangName}: {translatedText}\n\n{sourceLangName}: {originalText}"
                };

                var jsonLine1 = JsonConvert.SerializeObject(logEntry1, Formatting.None);
                var jsonLine2 = JsonConvert.SerializeObject(logEntry2, Formatting.None);

                lock (_globalLogLock)
                {
                    File.AppendAllText(_translationLogPath, jsonLine1 + Environment.NewLine);
                    File.AppendAllText(_translationLogPath2, jsonLine2 + Environment.NewLine);
                }

                _logger.LogDebug($"翻译日志已记录: {originalText.Substring(0, Math.Min(50, originalText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录翻译日志失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 记录预处理后的文本与译文的对照日志到JSON文件
        /// </summary>
        /// <param name="preprocessedText">预处理后的文本（术语替换后）</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        public void LogPreprocessedTranslation(string preprocessedText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(preprocessedText) || string.IsNullOrWhiteSpace(translatedText))
                return;

            if (preprocessedText == translatedText)
                return;

            try
            {
                var sourceLangName = GetLanguageName(sourceLang);
                var targetLangName = GetLanguageName(targetLang);

                var logEntry1 = new
                {
                    text = $"{sourceLangName}: {preprocessedText}\n\n{targetLangName}: {translatedText}"
                };

                var logEntry2 = new
                {
                    text = $"{targetLangName}: {translatedText}\n\n{sourceLangName}: {preprocessedText}"
                };

                var jsonLine1 = JsonConvert.SerializeObject(logEntry1, Formatting.None);
                var jsonLine2 = JsonConvert.SerializeObject(logEntry2, Formatting.None);

                lock (_globalLogLock)
                {
                    File.AppendAllText(_preprocessedLogPath, jsonLine1 + Environment.NewLine);
                    File.AppendAllText(_preprocessedLogPath2, jsonLine2 + Environment.NewLine);
                }

                _logger.LogDebug($"预处理翻译日志已记录: {preprocessedText.Substring(0, Math.Min(50, preprocessedText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录预处理翻译日志失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检测并处理异常译文
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        /// <returns>处理后的译文</returns>
        protected string ProcessAbnormalTranslation(string originalText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(translatedText))
                return translatedText;

            // 检测译文中是否包含原文内容（记录日志但不中断翻译）
            if (ContainsOriginalText(originalText, translatedText))
            {
                LogAbnormalTranslation(originalText, translatedText, sourceLang, targetLang);
                _logger.LogWarning($"译文包含原文内容，源语言: {sourceLang}, 目标语言: {targetLang}, 原文: {originalText.Substring(0, Math.Min(30, originalText.Length))}...");
            }

            // 检测译文是否包含重复内容，如有则去重
            if (IsDuplicateTranslation(translatedText))
            {
                var deduplicated = RemoveDuplicateContent(translatedText);
                if (!string.IsNullOrWhiteSpace(deduplicated) && deduplicated.Length < translatedText.Length * 0.8)
                {
                    _logger.LogWarning($"检测到重复译文，已去重。原文长度: {translatedText.Length}, 去重后长度: {deduplicated.Length}");
                    translatedText = deduplicated;
                }
            }

            return translatedText;
        }

        /// <summary>
        /// 检测译文是否包含重复内容
        /// </summary>
        /// <param name="text">译文文本</param>
        /// <returns>如果包含重复内容返回true</returns>
        protected bool IsDuplicateTranslation(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var words = text.Split(new[] { ' ', '\t', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length < 10)
                return false;

            var wordGroups = words.GroupBy(w => w.ToLower().Trim());
            var maxCount = wordGroups.Max(g => g.Count());

            if (maxCount > words.Length * 0.3)
                return true;

            var sentences = text.Split(new[] { '.', '!', '?', '。', '！', '？' }, StringSplitOptions.RemoveEmptyEntries);
            if (sentences.Length > 5)
            {
                var sentenceGroups = sentences.GroupBy(s => s.ToLower().Trim());
                var maxSentenceCount = sentenceGroups.Max(g => g.Count());

                if (maxSentenceCount > sentences.Length * 0.3)
                    return true;
            }

            var phrases = Regex.Split(text, @"[\s,;:]+");
            if (phrases.Length > 10)
            {
                var phraseGroups = phrases.GroupBy(p => p.ToLower().Trim());
                var maxPhraseCount = phraseGroups.Max(g => g.Count());

                if (maxPhraseCount > phrases.Length * 0.4)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// 去除译文中的重复内容
        /// </summary>
        /// <param name="text">译文文本</param>
        /// <returns>去除重复后的文本</returns>
        protected string RemoveDuplicateContent(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            var uniqueLines = new List<string>();
            var seenLines = new HashSet<string>(StringComparer.Ordinal);

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (!string.IsNullOrWhiteSpace(trimmedLine) && !seenLines.Contains(trimmedLine))
                {
                    seenLines.Add(trimmedLine);
                    uniqueLines.Add(trimmedLine);
                }
            }

            var result = string.Join("\n", uniqueLines);

            return result;
        }

        /// <summary>
        /// 检测译文是否异常（长度异常）
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <returns>如果译文异常返回true</returns>
        protected bool IsAbnormalTranslation(string originalText, string translatedText)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return false;

            var originalLength = originalText.Trim().Length;
            var translatedLength = translatedText.Trim().Length;

            if (originalLength < 10 && translatedLength > 2000)
                return true;

            if (originalLength > 10 && translatedLength > originalLength * 5)
                return true;

            if (originalLength > 100 && translatedLength < originalLength * 0.2)
                return true;

            return false;
        }

        /// <summary>
        /// 检测译文中是否包含原文内容
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <returns>如果译文中包含原文内容返回true</returns>
        protected bool ContainsOriginalText(string originalText, string translatedText)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return false;

            // 检查是否包含数字、符号或单位的文本
            bool ContainsNumbersOrSymbols(string text)
            {
                return Regex.IsMatch(text, @"[0-9]+|[≥≤<>%‰℃°kgkwwattslpmminmm]", RegexOptions.Compiled);
            }

            // 对于包含数字、符号或单位的文本，更加宽松
            if (ContainsNumbersOrSymbols(originalText))
            {
                return false;
            }

            // 对于短文本，更加宽松
            if (originalText.Length < 10)
            {
                return false;
            }

            // 去除空格和标点符号进行比较
            string CleanText(string text)
            {
                return Regex.Replace(text, @"\s+|[\p{P}]", "", RegexOptions.Compiled).ToLower();
            }

            var cleanedOriginal = CleanText(originalText);
            var cleanedTranslated = CleanText(translatedText);

            // 如果原文长度小于10个字符，不进行严格检查
            if (cleanedOriginal.Length < 10)
            {
                return false;
            }
            else
            {
                // 计算原文在译文中的匹配比例
                int matchCount = 0;
                int totalChecks = 0;
                
                // 只检查长度大于3的连续字符
                for (int i = 0; i < cleanedOriginal.Length - 3; i++)
                {
                    string substring = cleanedOriginal.Substring(i, 4); // 检查4个连续字符
                    if (cleanedTranslated.Contains(substring))
                    {
                        matchCount++;
                    }
                    totalChecks++;
                }

                // 如果匹配比例超过50%，才认为译文中包含原文内容（进一步提高阈值以减少误判）
                return totalChecks > 0 && (double)matchCount / totalChecks > 0.5;
            }
        }

        /// <summary>
        /// 检测译文质量
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        /// <returns>译文质量评估结果</returns>
        protected bool IsQualityTranslation(string originalText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(translatedText))
                return false;

            // 检查是否包含数字、符号或单位的文本
            bool ContainsNumbersOrSymbols(string text)
            {
                return Regex.IsMatch(text, @"[0-9]+|[≥≤<>%‰℃°kgkwwattslpmminmm]", RegexOptions.Compiled);
            }

            // 对于包含数字、符号或单位的文本，更加宽松
            if (ContainsNumbersOrSymbols(originalText))
            {
                // 只要译文不为空且长度合理，就认为质量可接受
                if (translatedText.Length > originalText.Length * 0.3)
                {
                    // 检查译文是否包含关键数字和符号
                    bool ContainsKeyElements(string text, string translated)
                    {
                        // 提取原文中的数字和符号
                        var originalElements = Regex.Matches(text, @"[0-9]+|[≥≤<>%‰℃°kgkwwattslpmminmm]", RegexOptions.Compiled);
                        if (originalElements.Count == 0) return true;

                        // 检查译文中是否包含至少一个关键元素
                        foreach (Match match in originalElements)
                        {
                            if (translated.Contains(match.Value))
                            {
                                return true;
                            }
                        }
                        return false;
                    }

                    return ContainsKeyElements(originalText, translatedText);
                }
                return false;
            }

            // 检测译文是否异常
            if (IsAbnormalTranslation(originalText, translatedText))
                return false;

            // 检测译文中是否包含原文内容
            if (ContainsOriginalText(originalText, translatedText))
                return false;

            // 检测译文是否与原文过于相似（排除短文本）
            if (originalText.Length > 15) // 提高阈值，减少误判
            {
                string CleanText(string text)
                {
                    return Regex.Replace(text, @"\s+|[\p{P}]", "", RegexOptions.Compiled).ToLower();
                }

                var cleanedOriginal = CleanText(originalText);
                var cleanedTranslated = CleanText(translatedText);

                // 计算相似度（简单的包含关系检查）
                if (cleanedTranslated.Contains(cleanedOriginal) || cleanedOriginal.Contains(cleanedTranslated))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 记录重复译文到Duplicate文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        public void LogDuplicateTranslation(string originalText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return;

            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logLine = $"[{timestamp}] 源语言: {sourceLang}, 目标语言: {targetLang}\n原文: {originalText}\n译文: {translatedText}\n---\n";

                lock (_globalLogLock)
                {
                    File.AppendAllText(_duplicateLogPath, logLine);
                }

                _logger.LogInformation($"重复译文已记录到Duplicate文件夹: {originalText.Substring(0, Math.Min(30, originalText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录重复译文失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 记录异常译文到Abnormal文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="translatedText">译文</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        public void LogAbnormalTranslation(string originalText, string translatedText, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return;

            // 对于包含数字、符号或单位的文本，跳过异常记录
            bool ContainsNumbersOrSymbols(string text)
            {
                return Regex.IsMatch(text, @"[0-9]+|[≥≤<>%‰℃°kggt]", RegexOptions.Compiled);
            }

            if (ContainsNumbersOrSymbols(originalText))
            {
                _logger.LogDebug($"跳过异常译文记录（包含数字/符号）: {originalText.Substring(0, Math.Min(20, originalText.Length))}...");
                return;
            }

            // 限制异常记录的频率，每10次只记录一次
            // if (Interlocked.Increment(ref _abnormalLogCounter) % 10 != 0)
            // {
            //     _logger.LogDebug($"跳过异常译文记录（频率限制）: {originalText.Substring(0, Math.Min(20, originalText.Length))}...");
            //     return;
            // }

            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logLine = $"[{timestamp}] 源语言: {sourceLang}, 目标语言: {targetLang}\n原文: {originalText}\n译文: {translatedText}\n原文长度: {originalText.Length}, 译文长度: {translatedText.Length}\n---\n";

                lock (_globalLogLock)
                {
                    File.AppendAllText(_abnormalLogPath, logLine);
                }

                _logger.LogDebug($"异常译文已记录到Abnormal文件夹: {originalText.Substring(0, Math.Min(30, originalText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录异常译文失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 记录翻译失败到Failure文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="errorMessage">错误信息</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        public void LogFailureTranslation(string originalText, string errorMessage, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(originalText))
                return;

            // 限制失败记录的频率，每5次只记录一次
            // if (Interlocked.Increment(ref _failureLogCounter) % 5 != 0)
            // {
            //     _logger.LogDebug($"跳过失败记录（频率限制）: {originalText.Substring(0, Math.Min(20, originalText.Length))}...");
            //     return;
            // }

            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logLine = $"[{timestamp}] 源语言: {sourceLang}, 目标语言: {targetLang}\n原文: {originalText}\n错误信息: {errorMessage}\n---\n";

                lock (_globalLogLock)
                {
                    File.AppendAllText(_failureLogPath, logLine);
                }

                _logger.LogDebug($"翻译失败已记录到Failure文件夹: {_failureLogPath}, 原文: {originalText.Substring(0, Math.Min(30, originalText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录翻译失败失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 记录翻译超时到TimeOut文件夹
        /// </summary>
        /// <param name="originalText">原文</param>
        /// <param name="timeoutSeconds">超时秒数</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        public void LogTimeoutTranslation(string originalText, int timeoutSeconds, string sourceLang, string targetLang)
        {
            if (string.IsNullOrWhiteSpace(originalText))
                return;

            try
            {
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                var logLine = $"[{timestamp}] 源语言: {sourceLang}, 目标语言: {targetLang}\n原文: {originalText}\n超时时间: {timeoutSeconds}秒\n---\n";

                lock (_globalLogLock)
                {
                    File.AppendAllText(_timeoutLogPath, logLine);
                }

                _logger.LogInformation($"翻译超时已记录到TimeOut文件夹: {_timeoutLogPath}, 原文: {originalText.Substring(0, Math.Min(30, originalText.Length))}...");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"记录翻译超时失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 释放资源
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// 释放资源的虚方法，供子类重写
        /// </summary>
        /// <param name="disposing">是否正在释放托管资源</param>
        protected virtual void Dispose(bool disposing)
        {
            // 基类默认不需要释放资源
        }

        #region 翻译前检查

        /// <summary>
        /// 判断是否为中文字符
        /// </summary>
        protected static bool IsChinese(char c)
        {
            return c >= 0x4E00 && c <= 0x9FFF;
        }

        /// <summary>
        /// 判断文本是否应跳过翻译（纯数字、纯编码等）
        /// </summary>
        protected bool ShouldSkipTranslation(string text)
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

        /// <summary>
        /// 判断文本是否为纯代码标识符
        /// </summary>
        private static bool IsPureCode(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.Trim();

            if (trimmed.Length < 2)
                return false;

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

            var hasMultipleConsecutiveLetters = maxConsecutiveLetters >= 2;

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

        /// <summary>
        /// 根据源语言字符比例判断是否需要翻译
        /// </summary>
        protected bool ShouldTranslateBasedOnLanguageRatio(string text, string sourceLang)
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

        #endregion
    }
}
