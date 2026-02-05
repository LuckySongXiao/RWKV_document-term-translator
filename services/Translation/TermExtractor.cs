using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 术语提取器，负责从术语库中提取相关术语
    /// </summary>
    public class TermExtractor
    {
        private readonly ILogger<TermExtractor> _logger;
        private Dictionary<string, Dictionary<string, string>> _terminologyData;

        public TermExtractor(ILogger<TermExtractor> logger = null)
        {
            _logger = logger;
            LoadTerminologyData();
        }

        /// <summary>
        /// 加载术语数据
        /// </summary>
        public void LoadTerminologyData()
        {
            try
            {
                var terminologyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "terminology.json");
                if (File.Exists(terminologyPath))
                {
                    var jsonContent = File.ReadAllText(terminologyPath);

                    // 兼容解析：既支持字符串值，也支持对象值（取 term 字段）
                    var result = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
                    using (var doc = JsonDocument.Parse(jsonContent))
                    {
                        if (doc.RootElement.ValueKind != JsonValueKind.Object)
                        {
                            _logger?.LogWarning("术语文件根节点不是对象，已忽略");
                            _terminologyData = new Dictionary<string, Dictionary<string, string>>();
                            return;
                        }

                        foreach (var langProp in doc.RootElement.EnumerateObject())
                        {
                            var language = langProp.Name;
                            var langValue = langProp.Value;
                            var terms = new Dictionary<string, string>(StringComparer.Ordinal);

                            if (langValue.ValueKind == JsonValueKind.Object)
                            {
                                foreach (var termProp in langValue.EnumerateObject())
                                {
                                    var termKey = termProp.Name;
                                    var termVal = termProp.Value;

                                    try
                                    {
                                        if (termVal.ValueKind == JsonValueKind.String)
                                        {
                                            var foreign = termVal.GetString();
                                            if (!string.IsNullOrWhiteSpace(termKey) && !string.IsNullOrWhiteSpace(foreign))
                                            {
                                                terms[termKey] = foreign;
                                            }
                                        }
                                        else if (termVal.ValueKind == JsonValueKind.Object)
                                        {
                                            if (termVal.TryGetProperty("term", out var termField) && termField.ValueKind == JsonValueKind.String)
                                            {
                                                var foreign = termField.GetString();
                                                if (!string.IsNullOrWhiteSpace(termKey) && !string.IsNullOrWhiteSpace(foreign))
                                                {
                                                    terms[termKey] = foreign;
                                                }
                                            }
                                            else
                                            {
                                                _logger?.LogWarning($"术语 '{termKey}' 在语言 '{language}' 中为对象，但缺少可用的 'term' 字段，已跳过");
                                            }
                                        }
                                        else
                                        {
                                            _logger?.LogWarning($"术语 '{termKey}' 在语言 '{language}' 中类型为 {termVal.ValueKind}，未被支持，已跳过");
                                        }
                                    }
                                    catch (Exception exItem)
                                    {
                                        _logger?.LogWarning(exItem, $"解析术语 '{termKey}' 失败（语言: {language}），已跳过");
                                    }
                                }
                            }
                            else
                            {
                                _logger?.LogWarning($"语言 '{language}' 的值不是对象（类型: {langValue.ValueKind}），已跳过");
                            }

                            result[language] = terms;
                        }
                    }

                    _terminologyData = result;
                    var langCount = _terminologyData?.Count ?? 0;
                    _logger?.LogInformation($"成功加载术语数据，包含 {langCount} 种语言");
                    if (langCount > 0)
                    {
                        _logger?.LogInformation($"支持语言键: {string.Join(", ", (_terminologyData?.Keys)?.AsEnumerable() ?? Enumerable.Empty<string>())}");
                        // 统计每种语言词条数
                        foreach (var kv in _terminologyData)
                        {
                            _logger?.LogInformation($"语言 '{kv.Key}' 词条数: {kv.Value?.Count ?? 0}");
                        }
                    }
                }
                else
                {
                    _logger?.LogWarning($"术语文件不存在: {terminologyPath}");
                    _terminologyData = new Dictionary<string, Dictionary<string, string>>();
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "加载术语数据失败");
                _terminologyData = new Dictionary<string, Dictionary<string, string>>();
            }
        }

        /// <summary>
        /// 获取术语库中支持的所有语种
        /// </summary>
        /// <returns>支持的语种列表</returns>
        public List<string> GetSupportedLanguages()
        {
            if (_terminologyData == null)
            {
                LoadTerminologyData();
            }

            return _terminologyData?.Keys.ToList() ?? new List<string>();
        }

        /// <summary>
        /// 将多种写法的语言名称/代码归一化到术语库顶层键（例如："英文"/"English"/"en" → "英语"）
        /// </summary>
        private static string NormalizeLanguageKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return key;
            var k = key.Trim();
            var lower = k.ToLowerInvariant();
            return lower switch
            {
                // 英语
                "英语" => "英语",
                "英文" => "英语",
                "english" => "英语",
                "en" => "英语",
                "en-us" => "英语",
                "en_gb" => "英语",
                "en-uk" => "英语",

                // 日语
                "日语" => "日语",
                "日文" => "日语",
                "japanese" => "日语",
                "ja" => "日语",

                // 韩语
                "韩语" => "韩语",
                "韩文" => "韩语",
                "korean" => "韩语",
                "ko" => "韩语",

                // 法语
                "法语" => "法语",
                "法文" => "法语",
                "french" => "法语",
                "fr" => "法语",

                // 德语
                "德语" => "德语",
                "德文" => "德语",
                "german" => "德语",
                "de" => "德语",

                // 西班牙语
                "西班牙语" => "西班牙语",
                "西语" => "西班牙语",
                "spanish" => "西班牙语",
                "es" => "西班牙语",

                // 意大利语
                "意大利语" => "意大利语",
                "italian" => "意大利语",
                "it" => "意大利语",

                // 俄语
                "俄语" => "俄语",
                "russian" => "俄语",
                "ru" => "俄语",

                _ => k
            };
        }

        /// <summary>
        /// 从文本中提取相关术语
        /// </summary>
        /// <param name="text">要分析的文本</param>
        /// <param name="targetLanguage">目标语言</param>
        /// <param name="sourceLang">源语言代码</param>
        /// <param name="targetLang">目标语言代码</param>
        /// <returns>相关术语字典</returns>
        public Dictionary<string, string> ExtractRelevantTerms(string text, string targetLanguage,
            string sourceLang = "zh", string targetLang = "en")
        {
            if (string.IsNullOrWhiteSpace(text) || _terminologyData == null)
            {
                return new Dictionary<string, string>();
            }

            try
            {
                // 获取目标语言的术语库（先归一化顶层键）
                var normalized = NormalizeLanguageKey(targetLanguage);
                if (!_terminologyData.TryGetValue(normalized, out var languageTerms))
                {
                    _logger?.LogWarning($"未找到目标语言 '{targetLanguage}' 的术语库（归一化为 '{normalized}'）");
                    return new Dictionary<string, string>();
                }

                var relevantTerms = new Dictionary<string, string>();

                // 根据翻译方向确定术语匹配逻辑
                bool isCnToForeign = sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh");

                // 按术语长度降序排序，优先匹配长术语
                var sortedTerms = isCnToForeign 
                    ? languageTerms.OrderByDescending(kv => kv.Key.Length).ToList()
                    : languageTerms.OrderByDescending(kv => kv.Value.Length).ToList();

                if (isCnToForeign)
                {
                    // 中文→外语：在文本中查找中文术语，返回对应的外语翻译
                    foreach (var (chineseTerm, foreignTerm) in sortedTerms)
                    {
                        if (!string.IsNullOrWhiteSpace(chineseTerm) && text.Contains(chineseTerm))
                        {
                            // 避免重复添加，长术语优先
                            if (!relevantTerms.ContainsKey(chineseTerm))
                            {
                                relevantTerms[chineseTerm] = foreignTerm;
                            }
                        }
                    }
                }
                else
                {
                    // 外语→中文：在文本中查找外语术语，返回对应的中文翻译
                    foreach (var (chineseTerm, foreignTerm) in sortedTerms)
                    {
                        if (!string.IsNullOrWhiteSpace(foreignTerm) && text.Contains(foreignTerm))
                        {
                            // 避免重复添加，长术语优先
                            if (!relevantTerms.ContainsKey(foreignTerm))
                            {
                                relevantTerms[foreignTerm] = chineseTerm;
                            }
                        }
                    }
                }

                _logger?.LogInformation($"从文本中提取到 {relevantTerms.Count} 个相关术语");
                return relevantTerms;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "提取相关术语失败");
                return new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// 预处理术语，按长度排序以优先匹配长术语
        /// 自动生成全角/半角括号的变体，增强匹配能力
        /// </summary>
        /// <param name="terms">术语字典</param>
        /// <returns>按长度排序的术语列表</returns>
        public List<KeyValuePair<string, string>> PreprocessTerms(Dictionary<string, string> terms)
        {
            if (terms == null || !terms.Any())
            {
                return new List<KeyValuePair<string, string>>();
            }

            var augmentedTerms = new Dictionary<string, string>(terms);

            // 自动生成括号变体
            foreach (var kv in terms)
            {
                var key = kv.Key;
                var value = kv.Value;
                
                // 如果包含半角括号，尝试生成全角括号变体
                if (key.Contains("(") || key.Contains(")"))
                {
                    var variant = key.Replace("(", "（").Replace(")", "）");
                    if (!augmentedTerms.ContainsKey(variant))
                    {
                        augmentedTerms[variant] = value;
                    }
                }
                
                // 如果包含全角括号，尝试生成半角括号变体
                if (key.Contains("（") || key.Contains("）"))
                {
                    var variant = key.Replace("（", "(").Replace("）", ")");
                    if (!augmentedTerms.ContainsKey(variant))
                    {
                        augmentedTerms[variant] = value;
                    }
                }
            }

            // 按源术语长度降序排序，确保优先匹配较长的术语
            var sortedTerms = augmentedTerms.OrderByDescending(kvp => kvp.Key.Length).ToList();

            _logger?.LogInformation($"术语预处理完成，共 {sortedTerms.Count} 个术语 (含自动生成的变体)");
            return sortedTerms;
        }

            /// <summary>
            /// 在文本中直接进行术语预替换：将匹配到的源术语替换为目标术语（非占位符）。
            /// - 中文→外语：按中文术语优先、长度降序，使用边界限制，避免误匹配
            /// - 外语→中文：按外语术语优先、长度降序，使用单词边界
            /// </summary>
            /// <param name="text">原始文本</param>
            /// <param name="terms">术语字典（键为源术语，值为目标术语）</param>
            /// <param name="isCnToForeign">是否中文→外语</param>
            /// <returns>替换后的文本</returns>
            public string ReplaceTermsInText(string text, Dictionary<string, string> terms, bool isCnToForeign)
            {
                if (string.IsNullOrWhiteSpace(text) || terms == null || !terms.Any())
                {
                    return text;
                }

                var sorted = terms.OrderByDescending(kvp => kvp.Key.Length).ToList();
                var result = text;

                foreach (var kv in sorted)
                {
                    var src = kv.Key;
                    var dst = kv.Value;
                    if (string.IsNullOrWhiteSpace(src)) continue;

                    try
                    {
                        string pattern;
                        if (isCnToForeign)
                        {
                            // 中文术语：仅在起始/结束为ASCII字母数字时添加边界检查，避免部分匹配
                            string srcEscaped = Regex.Escape(src);
                            
                            // 检查起始字符是否为ASCII字母/数字
                            bool startIsAlnum = src.Length > 0 && ((src[0] >= 'a' && src[0] <= 'z') || (src[0] >= 'A' && src[0] <= 'Z') || (src[0] >= '0' && src[0] <= '9'));
                            // 检查结束字符是否为ASCII字母/数字
                            bool endIsAlnum = src.Length > 0 && ((src[^1] >= 'a' && src[^1] <= 'z') || (src[^1] >= 'A' && src[^1] <= 'Z') || (src[^1] >= '0' && src[^1] <= '9'));

                            // 宽松的边界检查：仅排除字母前缀，允许数字前缀（适配 "10KW" 场景）
                            string prefix = startIsAlnum ? "(?<![a-zA-Z])" : "";
                            string suffix = endIsAlnum ? "(?![a-zA-Z0-9])" : "";
                            
                            pattern = $"{prefix}{srcEscaped}{suffix}";
                        }
                        else
                        {
                            // 外语术语：使用单词边界，避免部分词命中
                            pattern = $"\\b{Regex.Escape(src)}\\b";
                        }
                        // 用户需求：术语预处理过之后需要在术语后的术语文本前加一个英文字符的空格
                        result = Regex.Replace(result, pattern, " " + dst, RegexOptions.Compiled);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogWarning(ex, $"术语预替换失败: {src} -> {dst}");
                    }
                }

                return result;
            }



        /// <summary>
        /// 获取指定语言的所有术语
        /// </summary>
        /// <param name="targetLanguage">目标语言</param>
        /// <returns>术语字典</returns>
        public Dictionary<string, string> GetTermsForLanguage(string targetLanguage)
        {
            var normalized = NormalizeLanguageKey(targetLanguage);
            if (_terminologyData == null || !_terminologyData.TryGetValue(normalized, out var terms))
            {
                _logger?.LogWarning($"未找到语言 '{targetLanguage}' 的术语库（归一化为 '{normalized}'）");
                return new Dictionary<string, string>();
            }

            return new Dictionary<string, string>(terms);
        }

        /// <summary>
        /// 获取外译中的术语库（直接从data/reverse目录加载，不进行镜像）
        /// </summary>
        /// <param name="sourceLanguage">源语言（如"英语"）</param>
        /// <returns>外译中术语库（外语术语为键，中文术语为值）</returns>
        public Dictionary<string, string> GetReverseTermsForLanguage(string sourceLanguage)
        {
            try
            {
                var normalized = NormalizeLanguageKey(sourceLanguage);
                var reverseDataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "reverse");
                var reverseFilePath = Path.Combine(reverseDataDir, $"terminology_{normalized}.json");

                if (!File.Exists(reverseFilePath))
                {
                    _logger?.LogWarning($"未找到外译中术语库文件: {reverseFilePath}");
                    return new Dictionary<string, string>();
                }

                var jsonContent = File.ReadAllText(reverseFilePath);
                var reverseTerms = new Dictionary<string, string>(StringComparer.Ordinal);

                using (var doc = JsonDocument.Parse(jsonContent))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Object)
                    {
                        _logger?.LogWarning($"外译中术语库文件根节点不是对象: {reverseFilePath}");
                        return new Dictionary<string, string>();
                    }

                    foreach (var termProp in doc.RootElement.EnumerateObject())
                    {
                        var foreignTerm = termProp.Name;
                        var termVal = termProp.Value;

                        try
                        {
                            string chineseTerm = null;

                            if (termVal.ValueKind == JsonValueKind.String)
                            {
                                chineseTerm = termVal.GetString();
                            }
                            else if (termVal.ValueKind == JsonValueKind.Object)
                            {
                                if (termVal.TryGetProperty("term", out var termField) && termField.ValueKind == JsonValueKind.String)
                                {
                                    chineseTerm = termField.GetString();
                                }
                            }

                            if (!string.IsNullOrWhiteSpace(foreignTerm) && !string.IsNullOrWhiteSpace(chineseTerm))
                            {
                                reverseTerms[foreignTerm] = chineseTerm;
                            }
                        }
                        catch (Exception exItem)
                        {
                            _logger?.LogWarning(exItem, $"解析外译中术语 '{foreignTerm}' 失败（语言: {sourceLanguage}），已跳过");
                        }
                    }
                }

                _logger?.LogInformation($"成功加载外译中术语库: {sourceLanguage}，包含 {reverseTerms.Count} 个术语");

                var sampleTerms = reverseTerms.Take(3).Select(kv => $"'{kv.Key}' → '{kv.Value}'");
                _logger?.LogInformation($"外译中术语库样本: {string.Join(", ", sampleTerms)}");

                return reverseTerms;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"加载外译中术语库失败: {sourceLanguage}");
                return new Dictionary<string, string>();
            }
        }



        /// <summary>
        /// 镜像术语库：将中译外的术语库镜像为外译中的术语库
        /// </summary>
        /// <param name="sourceLanguage">源语言（如"英语"）</param>
        /// <returns>镜像后的术语库（外语术语为键，中文术语为值）</returns>
        public Dictionary<string, string> MirrorTerminology(string sourceLanguage)
        {
            try
            {
                var normalized = NormalizeLanguageKey(sourceLanguage);
                if (!_terminologyData.TryGetValue(normalized, out var originalTerms))
                {
                    _logger?.LogWarning($"未找到源语言 '{sourceLanguage}' 的术语库（归一化为 '{normalized}'）");
                    return new Dictionary<string, string>();
                }

                // 镜像术语库：原来是中文->外语，现在变成外语->中文
                var mirroredTerms = new Dictionary<string, string>();
                foreach (var (chineseTerm, foreignTerm) in originalTerms)
                {
                    if (!string.IsNullOrWhiteSpace(chineseTerm) && !string.IsNullOrWhiteSpace(foreignTerm))
                    {
                        // 镜像：外语术语作为键，中文术语作为值
                        mirroredTerms[foreignTerm] = chineseTerm;
                    }
                }

                _logger?.LogInformation($"术语库镜像完成：{sourceLanguage} -> 镜像库，包含 {mirroredTerms.Count} 个术语");

                // 显示镜像样本
                var sampleTerms = mirroredTerms.Take(3).Select(kv => $"'{kv.Key}' → '{kv.Value}'");
                _logger?.LogInformation($"镜像术语库样本: {string.Join(", ", sampleTerms)}");

                return mirroredTerms;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"术语库镜像失败: {ex.Message}");
                return new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// 从长到短遍历术语库进行预替换，并记录到CSV
        /// </summary>
        /// <param name="texts">待翻译的文本列表</param>
        /// <param name="mirroredTerms">镜像后的术语库</param>
        /// <param name="csvPath">CSV记录文件路径</param>
        /// <param name="pptFileName">PPT文件名（可选，用于在CSV中标识数据来源）</param>
        /// <returns>预替换后的文本列表</returns>
        public List<string> PreprocessTextsWithMirroredTerms(List<string> texts, Dictionary<string, string> mirroredTerms, string csvPath, string pptFileName = "")
        {
            if (texts == null || !texts.Any() || mirroredTerms == null || !mirroredTerms.Any())
            {
                _logger?.LogWarning("文本列表或镜像术语库为空，跳过预处理");
                return texts?.ToList() ?? new List<string>();
            }

            try
            {
                var processedTexts = new List<string>();
                var csvRecords = new List<string>();

                // CSV头部
                var header = "原始文本,预替换后文本,替换的术语数量,替换详情";
                if (!string.IsNullOrEmpty(pptFileName))
                {
                    header = $"PPT文件名,{header}";
                }
                csvRecords.Add(header);

                // 按术语长度从长到短排序
                var sortedTerms = mirroredTerms.OrderByDescending(kv => kv.Key.Length).ToList();
                _logger?.LogInformation($"开始从长到短遍历术语库进行预替换，术语数量: {sortedTerms.Count}");

                foreach (var originalText in texts)
                {
                    if (string.IsNullOrWhiteSpace(originalText))
                    {
                        processedTexts.Add(originalText);
                        continue;
                    }

                    var processedText = originalText;
                    var replacements = new List<string>();
                    var replacementCount = 0;

                    // 从长到短遍历术语库
                    foreach (var (foreignTerm, chineseTerm) in sortedTerms)
                    {
                        if (string.IsNullOrWhiteSpace(foreignTerm) || string.IsNullOrWhiteSpace(chineseTerm))
                            continue;

                        // 检查文本中是否包含该外语术语
                        if (processedText.Contains(foreignTerm))
                        {
                            // 判断术语是否包含特殊字符（非字母数字）
                            var hasSpecialChars = !Regex.IsMatch(foreignTerm, @"^[a-zA-Z0-9]+$");
                            
                            // 如果包含特殊字符，使用普通替换；否则使用单词边界精确匹配
                            if (hasSpecialChars)
                            {
                                // 对于包含特殊字符的术语，使用普通替换
                                // 用户需求：术语预处理过之后需要在术语后的术语文本前加一个英文字符的空格
                                processedText = processedText.Replace(foreignTerm, " " + chineseTerm);
                                replacements.Add($"{foreignTerm}→{chineseTerm}");
                                replacementCount++;
                            }
                            else
                            {
                                // 对于纯字母数字术语，使用正则表达式进行精确替换（单词边界）
                                var pattern = $@"\b{Regex.Escape(foreignTerm)}\b";
                                var matches = Regex.Matches(processedText, pattern, RegexOptions.IgnoreCase);

                                if (matches.Count > 0)
                                {
                                    // 用户需求：术语预处理过之后需要在术语后的术语文本前加一个英文字符的空格
                                    processedText = Regex.Replace(processedText, pattern, " " + chineseTerm, RegexOptions.IgnoreCase);
                                    replacements.Add($"{foreignTerm}→{chineseTerm}({matches.Count}次)");
                                    replacementCount += matches.Count;
                                }
                            }

                            _logger?.LogDebug($"术语预替换: '{foreignTerm}' → '{chineseTerm}'");
                        }
                    }

                    processedTexts.Add(processedText);

                    // 记录到CSV
                    var originalTextEscaped = EscapeCsvField(originalText);
                    var processedTextEscaped = EscapeCsvField(processedText);
                    var replacementDetails = EscapeCsvField(string.Join("; ", replacements));

                    var csvLine = $"{originalTextEscaped},{processedTextEscaped},{replacementCount},{replacementDetails}";
                    if (!string.IsNullOrEmpty(pptFileName))
                    {
                        csvLine = $"{EscapeCsvField(pptFileName)},{csvLine}";
                    }
                    csvRecords.Add(csvLine);

                    if (replacementCount > 0)
                    {
                        _logger?.LogInformation($"文本预替换完成: {replacementCount} 个术语被替换");
                        _logger?.LogDebug($"原文: {originalText.Substring(0, Math.Min(50, originalText.Length))}...");
                        _logger?.LogDebug($"替换后: {processedText.Substring(0, Math.Min(50, processedText.Length))}...");
                    }
                }

                // 写入CSV文件
                try
                {
                    System.IO.File.WriteAllLines(csvPath, csvRecords, System.Text.Encoding.UTF8);
                    _logger?.LogInformation($"术语预替换记录已保存到: {csvPath}");
                }
                catch (Exception csvEx)
                {
                    _logger?.LogError($"保存CSV记录失败: {csvEx.Message}");
                }

                return processedTexts;
            }
            catch (Exception ex)
            {
                _logger?.LogError($"术语预处理失败: {ex.Message}");
                return texts?.ToList() ?? new List<string>();
            }
        }

        /// <summary>
        /// 转义CSV字段
        /// </summary>
        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field))
                return "\"\"";

            // 如果包含逗号、引号或换行符，需要用引号包围并转义内部引号
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            }

            return field;
        }

        /// <summary>
        /// 添加或更新术语
        /// </summary>
        /// <param name="targetLanguage">目标语言</param>
        /// <param name="chineseTerm">中文术语</param>
        /// <param name="foreignTerm">外语术语</param>
        public void AddOrUpdateTerm(string targetLanguage, string chineseTerm, string foreignTerm)
        {
            try
            {
                if (_terminologyData == null)
                {
                    _terminologyData = new Dictionary<string, Dictionary<string, string>>();
                }

                if (!_terminologyData.ContainsKey(targetLanguage))
                {
                    _terminologyData[targetLanguage] = new Dictionary<string, string>();
                }

                _terminologyData[targetLanguage][chineseTerm] = foreignTerm;
                _logger?.LogInformation($"添加/更新术语: {chineseTerm} -> {foreignTerm} ({targetLanguage})");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "添加/更新术语失败");
            }
        }

        /// <summary>
        /// 删除术语
        /// </summary>
        /// <param name="targetLanguage">目标语言</param>
        /// <param name="chineseTerm">中文术语</param>
        public void RemoveTerm(string targetLanguage, string chineseTerm)
        {
            try
            {
                if (_terminologyData?.TryGetValue(targetLanguage, out var terms) == true)
                {
                    if (terms.Remove(chineseTerm))
                    {
                        _logger?.LogInformation($"删除术语: {chineseTerm} ({targetLanguage})");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "删除术语失败");
            }
        }

        /// <summary>
        /// 保存术语数据到文件
        /// </summary>
        public void SaveTerminologyData()
        {
            try
            {
                var terminologyPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "terminology.json");
                var directory = Path.GetDirectoryName(terminologyPath);

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var jsonContent = JsonSerializer.Serialize(_terminologyData, options);
                File.WriteAllText(terminologyPath, jsonContent);

                _logger?.LogInformation("术语数据保存成功");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "保存术语数据失败");
            }
        }

        /// <summary>
        /// 重新加载术语数据
        /// </summary>
        public void ReloadTerminologyData()
        {
            LoadTerminologyData();
        }

        /// <summary>
        /// 获取术语统计信息
        /// </summary>
        /// <returns>术语统计信息</returns>
        public Dictionary<string, int> GetTerminologyStatistics()
        {
            var statistics = new Dictionary<string, int>();

            if (_terminologyData != null)
            {
                foreach (var (language, terms) in _terminologyData)
                {
                    statistics[language] = terms.Count;
                }
            }

            return statistics;
        }

        /// <summary>
        /// 验证术语数据的完整性
        /// </summary>
        /// <returns>验证结果</returns>
        public (bool isValid, List<string> errors) ValidateTerminologyData()
        {
            var errors = new List<string>();

            if (_terminologyData == null)
            {
                errors.Add("术语数据为空");
                return (false, errors);
            }

            foreach (var (language, terms) in _terminologyData)
            {
                if (string.IsNullOrWhiteSpace(language))
                {
                    errors.Add("发现空的语言名称");
                }

                if (terms == null)
                {
                    errors.Add($"语言 '{language}' 的术语数据为空");
                    continue;
                }

                foreach (var (chineseTerm, foreignTerm) in terms)
                {
                    if (string.IsNullOrWhiteSpace(chineseTerm))
                    {
                        errors.Add($"语言 '{language}' 中发现空的中文术语");
                    }

                    if (string.IsNullOrWhiteSpace(foreignTerm))
                    {
                        errors.Add($"语言 '{language}' 中术语 '{chineseTerm}' 的外语翻译为空");
                    }
                }
            }

            return (errors.Count == 0, errors);
        }
    }
}
