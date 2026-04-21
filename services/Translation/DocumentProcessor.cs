using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentTranslator.Services.Translation.Translators;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 文档处理器，负责处理Word文档的翻译
    /// </summary>
    public class DocumentProcessor
    {
        private readonly TranslationService _translationService;
        private readonly ILogger<DocumentProcessor> _logger;
        private readonly TermExtractor _termExtractor;
        private bool _useTerminology = true;
        private bool _preprocessTerms = true;
        private bool _exportPdf = false;
        private string _sourceLang = "zh";
        private string _targetLang = "en";
        private bool _isCnToForeign = true;
        private string _outputFormat = "bilingual"; // 输出格式：bilingual（双语对照）或 translation_only（仅翻译结果）
        private bool _useModelForLanguageDetection = true; // 使用模型进行语言检测
        private Action<double, string> _progressCallback;
        private int _retryCount = 3;
        private int _retryDelay = 2000;
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private const double _minProgressUpdateInterval = 100; // 最小进度更新间隔100ms

        // 数学公式正则表达式模式
        private readonly List<Regex> _latexPatterns = new List<Regex>
        {
            new Regex(@"\$\$(.*?)\$\$", RegexOptions.Singleline), // 行间公式 $$...$$
            new Regex(@"\$(.*?)\$", RegexOptions.Singleline),      // 行内公式 $...$
            new Regex(@"\\begin\{equation\}(.*?)\\end\{equation\}", RegexOptions.Singleline), // equation环境
            new Regex(@"\\begin\{align\}(.*?)\\end\{align\}", RegexOptions.Singleline),       // align环境
            new Regex(@"\\begin\{eqnarray\}(.*?)\\end\{eqnarray\}", RegexOptions.Singleline)  // eqnarray环境
        };

        public DocumentProcessor(TranslationService translationService, ILogger<DocumentProcessor> logger)
        {
            _translationService = translationService ?? throw new ArgumentNullException(nameof(translationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _termExtractor = new TermExtractor();
        }

        /// <summary>
        /// 设置进度回调函数
        /// </summary>
        public void SetProgressCallback(Action<double, string> callback)
        {
            _progressCallback = callback;
        }

        /// <summary>
        /// 设置是否使用模型进行语言检测
        /// </summary>
        public void SetUseModelForLanguageDetection(bool useModel)
        {
            _useModelForLanguageDetection = useModel;
            _logger.LogInformation($"语言检测模式: {(useModel ? "AI模型" : "Unicode范围")}");
        }

        /// <summary>
        /// 设置翻译选项
        /// </summary>
        public void SetTranslationOptions(bool useTerminology = true, bool preprocessTerms = true,
            bool exportPdf = false, string sourceLang = "zh", string targetLang = "en", string outputFormat = "bilingual")
        {
            _useTerminology = useTerminology;
            _preprocessTerms = preprocessTerms;
            _exportPdf = exportPdf;
            _sourceLang = sourceLang;
            _targetLang = targetLang;
            _outputFormat = outputFormat;
            _isCnToForeign = sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh");

            _logger.LogInformation($"翻译方向: {(_isCnToForeign ? "中文→外语" : "外语→中文")}");
            _logger.LogInformation($"输出格式: {(_outputFormat == "bilingual" ? "双语对照" : "仅翻译结果")}");
        }

        /// <summary>
        /// 处理文档翻译
        /// </summary>
        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguage,
            Dictionary<string, string> terminology)
        {
            // 更新进度：开始处理
            UpdateProgress(0.01, "开始处理文档...");

            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                _logger.LogError($"文件不存在: {filePath}");
                throw new FileNotFoundException("文件不存在");
            }

            // 在程序根目录下创建输出目录
            var outputDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "输出");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // 获取输出路径，添加时间戳避免文件名冲突
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var originalExtension = Path.GetExtension(filePath).ToLower();
            var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            // 确保输出格式为Word文档格式
            var outputExtension = ".docx";
            var outputFileName = $"{fileName}_带翻译_{timeStamp}{outputExtension}";
            var outputPath = Path.Combine(outputDir, outputFileName);

            _logger.LogInformation($"Word处理器输出配置:");
            _logger.LogInformation($"  原始文件: {filePath}");
            _logger.LogInformation($"  原始扩展名: {originalExtension}");
            _logger.LogInformation($"  输出扩展名: {outputExtension}");
            _logger.LogInformation($"  输出路径: {outputPath}");

            // 更新进度：检查文件
            UpdateProgress(0.05, "检查文件权限...");

            try
            {
                // 复制文件到输出目录
                File.Copy(filePath, outputPath, true);

                // 更新进度：开始翻译
            UpdateProgress(0.1, "开始翻译文档内容...");

            // 开始新的翻译批次，为当前文档创建独立的日志文件
            _translationService.CurrentTranslator?.StartNewBatch();

            // 处理Word文档
            await ProcessWordDocumentAsync(outputPath, terminology);

                // 更新进度：完成
                UpdateProgress(1.0, "翻译完成");

                _logger.LogInformation($"文档翻译完成，输出路径: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "文档处理失败");
                throw;
            }
        }

        /// <summary>
        /// 处理Word文档
        /// </summary>
        private async Task ProcessWordDocumentAsync(string filePath, Dictionary<string, string> terminology)
        {
            using var document = WordprocessingDocument.Open(filePath, true);
            var body = document.MainDocumentPart.Document.Body;

            // 准备 CSV 路径
            var csvDir = Path.Combine(Path.GetDirectoryName(filePath)!, "日志");
            if (!Directory.Exists(csvDir)) Directory.CreateDirectory(csvDir);
            var csvPath = Path.Combine(csvDir, $"Word翻译日志_{DateTime.Now:yyyyMMddHHmmss}.csv");
            _logger.LogInformation($"将记录预处理与翻译到CSV: {csvPath}");

            // 写 CSV 头
            using var csvStream = new FileStream(csvPath, FileMode.Create, FileAccess.Write, FileShare.Read);
            using var csvWriter = new StreamWriter(csvStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            await csvWriter.WriteLineAsync("位置,原文,预处理后,译文");
            await csvWriter.FlushAsync();
            var csvLock = new object();

            // 获取所有段落（包括正文、页眉、页脚）
            var paragraphs = body.Descendants<Paragraph>().ToList();

            // 添加页眉段落
            foreach (var headerPart in document.MainDocumentPart.HeaderParts)
            {
                paragraphs.AddRange(headerPart.Header.Descendants<Paragraph>());
            }

            // 添加页脚段落
            foreach (var footerPart in document.MainDocumentPart.FooterParts)
            {
                paragraphs.AddRange(footerPart.Footer.Descendants<Paragraph>());
            }

            var totalParagraphs = paragraphs.Count;
            _logger.LogInformation($"开始处理 {totalParagraphs} 个段落");

            // 术语预处理：批量处理所有段落文本
            Dictionary<int, string> preprocessedDict = null;
            if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                _logger.LogInformation($"{(_isCnToForeign ? "中译外" : "外译中")}模式：开始批量术语预处理");

                // 收集所有段落文本
                var allParagraphTexts = paragraphs.Select(p => GetParagraphText(p)).ToList();

                // 准备CSV路径
                var preReplaceCsvPath = Path.Combine(csvDir, $"Word术语预替换_{DateTime.Now:yyyyMMddHHmmss}.csv");

                // 批量预处理
                var preprocessedTexts = _termExtractor.PreprocessTextsWithMirroredTerms(allParagraphTexts, terminology, preReplaceCsvPath);
                _logger.LogInformation($"术语预处理完成，已记录到: {preReplaceCsvPath}");

                // 存储预处理的文本供后续使用
                preprocessedDict = new Dictionary<int, string>();
                for (int i = 0; i < Math.Min(paragraphs.Count, preprocessedTexts.Count); i++)
                {
                    preprocessedDict[i] = preprocessedTexts[i];
                }
            }

            // 并发翻译阶段：仅做网络翻译调用，不写回文档
            var paragraphOriginals = paragraphs.Select(p => GetParagraphText(p)).ToList();
            var paragraphTranslations = new string[totalParagraphs];

            var tasks = new List<Task>();
            int finished = 0;

            var currentTranslatorType = _translationService.CurrentTranslatorType?.ToLowerInvariant() ?? "rwkv";
            int maxConcurrency = Math.Max(1, _translationService.GetEffectiveMaxParallelism());

            _logger.LogInformation($"当前翻译器: {currentTranslatorType}, 并发限制: {maxConcurrency}");
            var rwkvTranslator = _translationService.CurrentTranslator as RWKVTranslator;

            if (rwkvTranslator != null)
            {
                await ExecuteRwkvParagraphTranslationBatchesAsync(paragraphs, paragraphTranslations, terminology, preprocessedDict, csvWriter, csvLock, totalParagraphs, maxConcurrency, rwkvTranslator);
            }
            else
            {
                var semaphore = new SemaphoreSlim(maxConcurrency, maxConcurrency);

                for (int i = 0; i < totalParagraphs; i++)
                {
                    var index = i;
                    var para = paragraphs[index];
                    tasks.Add(Task.Run(async () =>
                    {
                        await semaphore.WaitAsync();
                        try
                        {
                            var textToTranslate = paragraphOriginals[index];
                            if (preprocessedDict != null && preprocessedDict.ContainsKey(index))
                            {
                                textToTranslate = preprocessedDict[index];
                            }

                            var (translated, original, preprocessed) = await ComputeParagraphTranslationAsync(para, terminology, textToTranslate);
                            paragraphTranslations[index] = translated;

                            if (!string.IsNullOrWhiteSpace(original))
                            {
                                lock (csvLock)
                                {
                                    string loc = $"段落_{index + 1}";
                                    csvWriter.WriteLineAsync($"{loc},{EscapeCsv(original)},{EscapeCsv(preprocessed)},{EscapeCsv(translated)}").GetAwaiter().GetResult();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, $"并发翻译段落失败，索引 {index}，将跳过写回");
                        }
                        finally
                        {
                            semaphore.Release();
                            var done = Interlocked.Increment(ref finished);
                            var progress = 0.1 + (0.7 * done / Math.Max(1, totalParagraphs));
                            UpdateProgress(progress, $"正在并发翻译段落 {done}/{totalParagraphs}");
                        }
                    }));
                }
                await Task.WhenAll(tasks);
                semaphore.Dispose();
            }

            UpdateProgress(0.8, $"段落翻译完成，准备写回 {totalParagraphs} 个段落");

            // 顺序写回阶段：避免OpenXML并发写入
            for (int i = 0; i < totalParagraphs; i++)
            {
                try
                {
                    var translated = paragraphTranslations[i];
                    if (!string.IsNullOrWhiteSpace(translated))
                    {
                        UpdateParagraphText(paragraphs[i], paragraphOriginals[i], translated);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"写回段落失败，索引 {i}，已跳过");
                }

                var writeProgress = 0.8 + (0.1 * (i + 1) / Math.Max(1, totalParagraphs));
                UpdateProgress(writeProgress, $"正在写回段落 {i + 1}/{totalParagraphs}");
            }

            // 处理表格：全面并发优化
            var tables = body.Descendants<Table>().ToList();

            // 添加页眉表格
            foreach (var headerPart in document.MainDocumentPart.HeaderParts)
            {
                tables.AddRange(headerPart.Header.Descendants<Table>());
            }

            // 添加页脚表格
            foreach (var footerPart in document.MainDocumentPart.FooterParts)
            {
                tables.AddRange(footerPart.Footer.Descendants<Table>());
            }

            if (tables.Any())
            {
                _logger.LogInformation($"开始处理 {tables.Count} 个表格");

                // 预处理术语字典
                Dictionary<string, string> termsToUse = terminology;
                if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
                {
                    if (!_isCnToForeign)
                    {
                        var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                        foreach (var kv in terminology)
                        {
                            var src = kv.Value; var dst = kv.Key;
                            if (string.IsNullOrWhiteSpace(src)) continue;
                            if (!reversed.ContainsKey(src)) reversed[src] = dst;
                        }
                        termsToUse = reversed;
                    }
                }

                var allTableJobs = new List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)>();
                
                for (int i = 0; i < tables.Count; i++)
                {
                    try 
                    {
                        var jobs = CollectTableTranslationJobs(tables[i], termsToUse, terminology);
                        allTableJobs.AddRange(jobs);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"收集表格 {i} 任务失败");
                    }
                }

                if (allTableJobs.Count > 0)
                {
                    _logger.LogInformation($"共收集到 {allTableJobs.Count} 个表格段落任务，开始并发翻译...");
                    UpdateProgress(0.9, $"正在并发翻译表格内容 ({allTableJobs.Count} 个段落)...");
                    
                    await ExecuteTableTranslationJobsAsync(allTableJobs, csvWriter, csvLock);
                }
            }

            // 保存文档
            document.MainDocumentPart.Document.Save();
        }

        private async Task ExecuteRwkvParagraphTranslationBatchesAsync(
            List<Paragraph> paragraphs,
            string[] paragraphTranslations,
            Dictionary<string, string> terminology,
            Dictionary<int, string> preprocessedDict,
            StreamWriter csvWriter,
            object csvLock,
            int totalParagraphs,
            int maxConcurrency,
            RWKVTranslator rwkvTranslator)
        {
            var jobs = new List<(int index, string original, string preprocessed, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)>();

            for (int i = 0; i < paragraphs.Count; i++)
            {
                var job = PrepareParagraphTranslationJob(paragraphs[i], terminology, preprocessedDict != null && preprocessedDict.TryGetValue(i, out var preprocessedText) ? preprocessedText : null);
                if (!string.IsNullOrWhiteSpace(job.original))
                {
                    jobs.Add((i, job.original, job.preprocessed, job.toTranslate, job.terms, job.placeholders));
                }
            }

            if (jobs.Count == 0)
            {
                return;
            }

            int batchSize = GetRwkvBatchSize(maxConcurrency);
            int finished = 0;
            int totalBatches = (jobs.Count + batchSize - 1) / batchSize;
            _logger.LogInformation($"RWKV段落翻译切换为真实批量模式，批大小: {batchSize}，任务数: {jobs.Count}，批次数: {totalBatches}，并发限制: {maxConcurrency}");

            // 构建所有批次
            var batches = new List<List<(int index, string original, string preprocessed, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)>>();
            for (int offset = 0; offset < jobs.Count; offset += batchSize)
            {
                batches.Add(jobs.Skip(offset).Take(batchSize).ToList());
            }

            // 并行执行所有批次，信号量由 TranslateBatchAsync 内部控制并发
            var batchTasks = batches.Select(async (batch, batchIndex) =>
            {
                IReadOnlyList<string> batchResults = null;

                try
                {
                    batchResults = await rwkvTranslator.TranslateBatchAsync(batch.Select(x => x.toTranslate).ToList(), _sourceLang, _targetLang);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"RWKV真实批量翻译失败，将回退到逐段翻译，批次: {batchIndex + 1}/{totalBatches}，批大小: {batch.Count}");
                }

                for (int i = 0; i < batch.Count; i++)
                {
                    var job = batch[i];
                    string translatedText = batchResults != null && i < batchResults.Count
                        ? batchResults[i]
                        : await TranslateTextWithRetryAsync(job.toTranslate, job.terms, job.original);

                    if (!string.IsNullOrWhiteSpace(translatedText))
                    {
                        paragraphTranslations[job.index] = RestoreMathFormulas(translatedText, job.placeholders);
                    }

                    lock (csvLock)
                    {
                        string loc = $"段落_{job.index + 1}";
                        csvWriter.WriteLineAsync($"{loc},{EscapeCsv(job.original)},{EscapeCsv(job.preprocessed)},{EscapeCsv(paragraphTranslations[job.index])}").GetAwaiter().GetResult();
                    }

                    int currentFinished = Interlocked.Increment(ref finished);
                    var progress = 0.1 + (0.7 * currentFinished / Math.Max(1, totalParagraphs));
                    UpdateProgress(progress, $"正在真实批量翻译段落 {currentFinished}/{totalParagraphs}");
                }
            }).ToArray();

            await Task.WhenAll(batchTasks);
        }

        private (string original, string preprocessed, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders) PrepareParagraphTranslationJob(
            Paragraph paragraph,
            Dictionary<string, string> terminology,
            string preprocessedText = null)
        {
            if (paragraph.Ancestors<Table>().Any()) return (string.Empty, string.Empty, string.Empty, null, null);

            var runs = paragraph.Descendants<Run>().ToList();
            if (!runs.Any()) return (string.Empty, string.Empty, string.Empty, null, null);

            var paragraphText = GetParagraphText(paragraph);
            if (string.IsNullOrWhiteSpace(paragraphText)) return (string.Empty, string.Empty, string.Empty, null, null);

            var (protectedText, mathPlaceholders) = ProtectMathFormulas(paragraphText);
            string toTranslate = protectedText;
            string preprocessed = protectedText;
            Dictionary<string, string> termsForTranslator = terminology;

            if (!string.IsNullOrWhiteSpace(preprocessedText))
            {
                toTranslate = preprocessedText;
                preprocessed = preprocessedText;
                termsForTranslator = null;
            }
            else if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                Dictionary<string, string> termsToUse = terminology;
                if (!_isCnToForeign)
                {
                    var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var kv in terminology)
                    {
                        var src = kv.Value;
                        var dst = kv.Key;
                        if (string.IsNullOrWhiteSpace(src)) continue;
                        if (!reversed.ContainsKey(src)) reversed[src] = dst;
                    }
                    termsToUse = reversed;
                }

                toTranslate = _termExtractor.ReplaceTermsInText(protectedText, termsToUse, _isCnToForeign);
                preprocessed = toTranslate;
                termsForTranslator = null;
            }

            return (paragraphText, preprocessed, toTranslate, termsForTranslator, mathPlaceholders);
        }

        private int GetRwkvBatchSize(int maxConcurrency)
        {
            // 单批次大小等于并发数，与并发测试时一致
            // batch-translate API 一次请求发送 maxConcurrency 个段落
            // GPU 显存占用与并发测试相同，充分利用 GPU 并行能力
            return Math.Max(1, maxConcurrency);
        }

        private string EscapeCsv(string s)
        {
            if (s == null) return "";
            s = s.Replace("\r\n", "\n").Replace("\r", "\n");
            if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }

        // 仅执行预处理与网络翻译，返回译文；不写回段落
        private async Task<(string translated, string original, string preprocessed)> ComputeParagraphTranslationAsync(Paragraph paragraph, Dictionary<string, string> terminology, string preprocessedText = null)
        {
            var job = PrepareParagraphTranslationJob(paragraph, terminology, preprocessedText);
            if (string.IsNullOrWhiteSpace(job.original)) return (string.Empty, string.Empty, string.Empty);

            var translatedText = await TranslateTextWithRetryAsync(job.toTranslate, job.terms, job.original);
            if (string.IsNullOrEmpty(translatedText)) return (string.Empty, job.original, job.preprocessed);

            // 恢复数学公式
            translatedText = RestoreMathFormulas(translatedText, job.placeholders);
            return (translatedText, job.original, job.preprocessed);
        }

        /// <summary>
        /// 处理段落
        /// </summary>
        private async Task ProcessParagraphAsync(Paragraph paragraph, Dictionary<string, string> terminology)
        {
            // 若位于表格内，交由表格逻辑处理，避免重复翻译
            if (paragraph.Ancestors<Table>().Any()) return;

            var runs = paragraph.Descendants<Run>().ToList();
            if (!runs.Any()) return;

            // 提取段落文本
            var paragraphText = GetParagraphText(paragraph);
            if (string.IsNullOrWhiteSpace(paragraphText)) return;

            // 保护数学公式
            var (protectedText, mathPlaceholders) = ProtectMathFormulas(paragraphText);

            // 术语预替换：与Excel一致，翻译前在文本中直接替换术语
            string toTranslate = protectedText;
            Dictionary<string, string> termsForTranslator = terminology;
            if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                // 规范方向：键始终是“源语言术语”
                var termsForReplace = terminology;
                if (!_isCnToForeign)
                {
                    var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var kv in terminology)
                    {
                        var src = kv.Value; // 外语
                        var dst = kv.Key;   // 中文
                        if (string.IsNullOrWhiteSpace(src)) continue;
                        if (!reversed.ContainsKey(src)) reversed[src] = dst;
                    }
                    termsForReplace = reversed;
                }

                var before = toTranslate;
                toTranslate = _termExtractor.ReplaceTermsInText(toTranslate, termsForReplace, _isCnToForeign);
                var beforeSnippet = before.Substring(0, Math.Min(80, before.Length));
                var afterSnippet = toTranslate.Substring(0, Math.Min(80, toTranslate.Length));
                _logger.LogInformation($"[Word] 术语预替换已执行（{(_isCnToForeign ? "中文→外语" : "外语→中文")}）：'{beforeSnippet}' -> '{afterSnippet}'");

                // 预处理后，发给翻译器的术语表置空，避免二次干预
                termsForTranslator = null;
            }

            // 翻译文本
            var translatedText = await TranslateTextWithRetryAsync(toTranslate, termsForTranslator, paragraphText);
            if (string.IsNullOrEmpty(translatedText)) return;

            // 恢复数学公式
            translatedText = RestoreMathFormulas(translatedText, mathPlaceholders);

            // 更新段落内容
            UpdateParagraphText(paragraph, paragraphText, translatedText);
        }

        private async Task ProcessParagraphInTableAsync(Paragraph paragraph, Dictionary<string, string> terminology)
        {
            var runs = paragraph.Descendants<Run>().ToList();
            if (!runs.Any()) return;

            // 提取段落文本
            var paragraphText = GetParagraphText(paragraph);
            if (string.IsNullOrWhiteSpace(paragraphText)) return;

            // 保护数学公式
            var (protectedText, mathPlaceholders) = ProtectMathFormulas(paragraphText);

            // 术语预替换
            string toTranslate = protectedText;
            Dictionary<string, string> termsForTranslator = terminology;
            if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                var termsForReplace = terminology;
                if (!_isCnToForeign)
                {
                    var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var kv in terminology)
                    {
                        var src = kv.Value; var dst = kv.Key;
                        if (string.IsNullOrWhiteSpace(src)) continue;
                        if (!reversed.ContainsKey(src)) reversed[src] = dst;
                    }
                    termsForReplace = reversed;
                }
                toTranslate = _termExtractor.ReplaceTermsInText(toTranslate, termsForReplace, _isCnToForeign);
                termsForTranslator = null;
            }

            // 翻译文本
            var translatedText = await TranslateTextWithRetryAsync(toTranslate, termsForTranslator, paragraphText);
            if (string.IsNullOrEmpty(translatedText)) return;

            // 恢复数学公式
            translatedText = RestoreMathFormulas(translatedText, mathPlaceholders);

            // 更新段落内容
            UpdateParagraphText(paragraph, paragraphText, translatedText);
        }

        /// <summary>
        /// 处理表格（单表处理，保留兼容性）
        /// </summary>
        private async Task ProcessTableAsync(Table table, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock)
        {
            // 0. 预处理术语字典
            Dictionary<string, string> termsToUse = terminology;
            if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                if (!_isCnToForeign)
                {
                    var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var kv in terminology)
                    {
                        var src = kv.Value; var dst = kv.Key;
                        if (string.IsNullOrWhiteSpace(src)) continue;
                        if (!reversed.ContainsKey(src)) reversed[src] = dst;
                    }
                    termsToUse = reversed;
                }
            }

            var jobs = CollectTableTranslationJobs(table, termsToUse, terminology);
            if (jobs.Count > 0)
            {
                await ExecuteTableTranslationJobsAsync(jobs, csvWriter, csvLock);
            }
        }

        private List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)> CollectTableTranslationJobs(Table table, Dictionary<string, string> termsToUse, Dictionary<string, string> rawTerminology)
        {
            var jobs = new List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)>();
            var rows = table.Descendants<TableRow>().ToList();

            foreach (var row in rows)
            {
                var cells = row.Descendants<TableCell>().ToList();
                foreach (var cell in cells)
                {
                    var cellText = cell.InnerText?.Trim();
                    if (ShouldSkipCellText(cellText)) continue;

                    var paragraphs = cell.Descendants<Paragraph>().ToList();
                    foreach (var p in paragraphs)
                    {
                        var paragraphText = GetParagraphText(p);
                        if (string.IsNullOrWhiteSpace(paragraphText)) continue;
                        if (!ContainsSourceLanguage(paragraphText)) continue;

                        var (protectedText, mathPlaceholders) = ProtectMathFormulas(paragraphText);

                        string toTranslate = protectedText;
                        Dictionary<string, string> termsForTranslator = rawTerminology;

                        if (_useTerminology && _preprocessTerms && termsToUse != null && termsToUse.Count > 0)
                        {
                            toTranslate = _termExtractor.ReplaceTermsInText(protectedText, termsToUse, _isCnToForeign);
                            termsForTranslator = null;
                        }

                        jobs.Add((p, paragraphText, toTranslate, termsForTranslator, mathPlaceholders));
                    }
                }
            }
            return jobs;
        }

        private async Task ExecuteTableTranslationJobsAsync(List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)> jobs, StreamWriter csvWriter, object csvLock)
        {
            if (jobs.Count == 0) return;

            var rwkvTranslator = _translationService.CurrentTranslator as RWKVTranslator;
            if (rwkvTranslator != null)
            {
                await ExecuteRwkvTableTranslationJobsAsync(jobs, csvWriter, csvLock, rwkvTranslator);
                return;
            }

            // 并发执行翻译
            int maxConcurrency = Math.Max(1, _translationService.GetEffectiveMaxParallelism());
            using var semaphore = new SemaphoreSlim(maxConcurrency);
            var tasks = new List<Task<(Paragraph p, string original, string translated, Dictionary<string, string> placeholders)>>();
            int finished = 0;
            int total = jobs.Count;

            foreach (var job in jobs)
            {
                tasks.Add(Task.Run(async () =>
                {
                    await semaphore.WaitAsync();
                    try
                    {
                        var translatedText = await TranslateTextWithRetryAsync(job.toTranslate, job.terms, job.original);
                        return (job.p, job.original, translatedText, job.placeholders);
                    }
                    finally
                    {
                        semaphore.Release();
                        var done = Interlocked.Increment(ref finished);
                        // 进度范围 0.9 ~ 1.0
                        var progress = 0.9 + (0.1 * done / Math.Max(1, total));
                        UpdateProgress(progress, $"正在并发翻译表格内容 {done}/{total}");
                    }
                }));
            }

            var results = await Task.WhenAll(tasks);

            // 顺序更新文档（OpenXML非线程安全）
            foreach (var result in results)
            {
                var translatedText = result.translated;
                if (!string.IsNullOrWhiteSpace(translatedText))
                {
                    // 恢复数学公式
                    translatedText = RestoreMathFormulas(translatedText, result.placeholders);

                    // 记录到CSV
                    if (csvWriter != null)
                    {
                        lock (csvLock)
                        {
                            string esc(string s) => s == null ? "" : "\"" + s.Replace("\r\n", "\n").Replace("\"", "\"\"") + "\"";
                            csvWriter.WriteLineAsync($"表格段落,{esc(result.original)},{esc(result.original)},{esc(translatedText)}").GetAwaiter().GetResult();
                        }
                    }

                    // 更新段落文本
                    try
                    {
                        UpdateParagraphText(result.p, result.original, translatedText);
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, $"表格段落回写失败: {result.original.Substring(0, Math.Min(20, result.original.Length))}...");
                    }
                }
            }
        }

        private async Task ExecuteRwkvTableTranslationJobsAsync(
            List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)> jobs,
            StreamWriter csvWriter,
            object csvLock,
            RWKVTranslator rwkvTranslator)
        {
            int batchSize = GetRwkvBatchSize(Math.Max(1, _translationService.GetEffectiveMaxParallelism()));
            int finished = 0;
            int total = jobs.Count;
            int totalBatches = (jobs.Count + batchSize - 1) / batchSize;

            _logger.LogInformation($"RWKV表格翻译切换为真实批量模式，批大小: {batchSize}，任务数: {total}，批次数: {totalBatches}");

            // 构建所有批次
            var batches = new List<List<(Paragraph p, string original, string toTranslate, Dictionary<string, string> terms, Dictionary<string, string> placeholders)>>();
            for (int offset = 0; offset < jobs.Count; offset += batchSize)
            {
                batches.Add(jobs.Skip(offset).Take(batchSize).ToList());
            }

            // 并行执行所有批次，信号量由 TranslateBatchAsync 内部控制并发
            var batchTasks = batches.Select(async (batch, batchIndex) =>
            {
                IReadOnlyList<string> batchResults = null;

                try
                {
                    batchResults = await rwkvTranslator.TranslateBatchAsync(batch.Select(x => x.toTranslate).ToList(), _sourceLang, _targetLang);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"RWKV表格真实批量翻译失败，将回退到逐段翻译，批次: {batchIndex + 1}/{totalBatches}，批大小: {batch.Count}");
                }

                for (int i = 0; i < batch.Count; i++)
                {
                    var job = batch[i];
                    var translatedText = batchResults != null && i < batchResults.Count
                        ? batchResults[i]
                        : await TranslateTextWithRetryAsync(job.toTranslate, job.terms, job.original);

                    if (!string.IsNullOrWhiteSpace(translatedText))
                    {
                        translatedText = RestoreMathFormulas(translatedText, job.placeholders);
                        UpdateParagraphText(job.p, job.original, translatedText);
                    }

                    if (csvWriter != null)
                    {
                        lock (csvLock)
                        {
                            csvWriter.WriteLineAsync($"表格,{EscapeCsv(job.original)},{EscapeCsv(job.toTranslate)},{EscapeCsv(translatedText)}").GetAwaiter().GetResult();
                        }
                    }

                    int currentFinished = Interlocked.Increment(ref finished);
                    var progress = 0.9 + (0.1 * currentFinished / Math.Max(1, total));
                    UpdateProgress(progress, $"正在真实批量翻译表格内容 {currentFinished}/{total}");
                }
            }).ToArray();

            await Task.WhenAll(batchTasks);
        }

        /// <summary>
        /// 保护数学公式
        /// </summary>
        private (string, Dictionary<string, string>) ProtectMathFormulas(string text)
        {
            var placeholders = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(text)) return (text, placeholders);

            string protectedText = text;
            int counter = 0;

            foreach (var pattern in _latexPatterns)
            {
                protectedText = pattern.Replace(protectedText, match =>
                {
                    string placeholder = $"MATH_FORMULA_{counter++}_END";
                    placeholders[placeholder] = match.Value;
                    return placeholder;
                });
            }

            return (protectedText, placeholders);
        }

        /// <summary>
        /// 恢复数学公式
        /// </summary>
        private string RestoreMathFormulas(string text, Dictionary<string, string> placeholders)
        {
            if (string.IsNullOrWhiteSpace(text) || placeholders.Count == 0) return text;

            string restoredText = text;
            foreach (var kv in placeholders)
            {
                restoredText = restoredText.Replace(kv.Key, kv.Value);
            }

            return restoredText;
        }

        /// <summary>
        /// 带重试的翻译
        /// </summary>
        private async Task<string> TranslateTextWithRetryAsync(string text, Dictionary<string, string> terminology, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;

            // 检查是否应该跳过
            if (ShouldSkipText(text)) return string.Empty;

            // 检查是否已经翻译（简单的启发式规则）
            // 优先检查原始文本（避免因术语替换导致误判），若无原始文本则检查当前文本
            var textToCheck = originalText ?? text;
            if (IsAlreadyTranslated(textToCheck))
            {
                _logger.LogInformation($"文本疑似已翻译，跳过: {textToCheck.Substring(0, Math.Min(20, textToCheck.Length))}...");
                return string.Empty;
            }

            for (int attempt = 0; attempt < _retryCount; attempt++)
            {
                try
                {
                    return await _translationService.TranslateTextAsync(text, terminology, _sourceLang, _targetLang);
                }
                catch (Exception ex)
                {
                    bool isQualityError = ex.Message.Contains("译文质量不佳");
                    
                    if (isQualityError)
                    {
                        // 质量不佳错误，不再重试，直接返回空字符串
                        _logger.LogDebug($"译文质量不佳，不再重试");
                        return string.Empty;
                    }
                    
                    _logger.LogWarning(ex, $"翻译失败，尝试 {attempt + 1}/{_retryCount}");

                    if (attempt == _retryCount - 1)
                    {
                        _logger.LogError("翻译重试次数已用完，跳过该文本");
                        return string.Empty;
                    }

                    await Task.Delay(_retryDelay);
                }
            }

            return string.Empty;
        }

        private string GetParagraphText(Paragraph p)
        {
            var sb = new StringBuilder();
            foreach (var run in p.Descendants<Run>())
            {
                foreach (var child in run.ChildElements)
                {
                    if (child is Text t) sb.Append(t.Text);
                    else if (child is Break) sb.Append("\n");
                    else if (child is CarriageReturn) sb.Append("\n");
                    else if (child is TabChar) sb.Append("\t");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// 更新段落文本
        /// </summary>
        private void UpdateParagraphText(Paragraph paragraph, string originalText, string translatedText)
        {
            var clean = CleanTranslationTextWithOriginal(originalText, translatedText);
            
            if (string.IsNullOrWhiteSpace(clean) && !string.IsNullOrWhiteSpace(translatedText))
            {
                _logger?.LogWarning($"译文被清洗为空，可能是因为包含了过滤词或格式问题。原文: {originalText.Substring(0, Math.Min(20, originalText.Length))}... 原始译文: {translatedText.Substring(0, Math.Min(50, translatedText.Length))}...");
            }

            var lines = clean.Replace("\r\n", "\n").Split('\n');

            if (_outputFormat == "bilingual")
            {
                var runs = paragraph.Descendants<Run>().ToList();
                var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

                paragraph.AppendChild(new Run(new Break()));

                for (int i = 0; i < lines.Length; i++)
                {
                    if (i > 0) paragraph.AppendChild(new Run(new Break()));

                    var run = new Run(new Text(lines[i]));
                    if (baseRunProps != null)
                        run.RunProperties = (RunProperties)baseRunProps.CloneNode(true);
                    else
                        run.RunProperties = new RunProperties();

                    if (run.RunProperties == null) run.RunProperties = new RunProperties();
                    run.RunProperties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "0066CC" };
                    paragraph.AppendChild(run);
                }
            }
            else
            {
                var runs = paragraph.Descendants<Run>().ToList();
                var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

                foreach (var run in runs) run.Remove();

                for (int i = 0; i < lines.Length; i++)
                {
                    if (i > 0) paragraph.AppendChild(new Run(new Break()));

                    var run = new Run(new Text(lines[i]));
                    if (baseRunProps != null)
                        run.RunProperties = (RunProperties)baseRunProps.CloneNode(true);
                    paragraph.AppendChild(run);
                }
            }
        }

        /// <summary>
        /// 更新进度
        /// </summary>
        private void UpdateProgress(double progress, string message)
        {
            try
            {
                var now = DateTime.Now;
                var elapsedSinceLastUpdate = (now - _lastProgressUpdate).TotalMilliseconds;

                // 节流：只在进度变化较大或超过最小间隔时更新UI
                if (elapsedSinceLastUpdate >= _minProgressUpdateInterval || progress >= 1.0)
                {
                    _lastProgressUpdate = now;
                    _progressCallback?.Invoke(progress, message);
                    _logger.LogInformation($"进度: {progress:P1} - {message}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "更新进度失败");
            }
        }

        private string CleanTranslationTextWithOriginal(string original, string translated)
        {
            var clean = CleanTranslationText(translated);
            if (string.IsNullOrWhiteSpace(clean)) return clean;

            // 移除过于激进的短文本截断逻辑，以修复"文本分段"问题（防止误删多行译文）
            // 原逻辑会强制保留第一行，导致多义词或长译文丢失
            
            return clean;
        }

        /// <summary>
        /// 清洗译文，移除前缀/代码块/说明等，返回纯净译文
        /// </summary>
        private string CleanTranslationText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var s = text.Trim();

            // 内置基础规则
            s = Regex.Replace(s, @"^(译文|翻译|Translation|Translated|Result|Output|Note|提示|系统|AI|Assistant|Assistant's answer)\s*[:：\-\]]*\s*", string.Empty, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^```[a-zA-Z0-9_-]*\s*", string.Empty);
            s = Regex.Replace(s, @"```\s*$", string.Empty);
            s = s.Trim().Trim('"', '\'', '“', '”', '‘', '’');
            s = Regex.Replace(s, @"(\n|\r)*【?注:.*$", string.Empty, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"(\n|\r)*\(?Translated by.*$", string.Empty, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^\(This is.*?\)$", string.Empty, RegexOptions.IgnoreCase);
            // 过滤“请提供纯正的XX语翻译”等指令性文本（中英）
            s = Regex.Replace(s, @"^(请提供|请给出|请生成|请输出|请用|请以).*?(中文|汉语|英文|英语|日文|日语|韩文|韩语|法文|法语|德文|德语|俄文|俄语|西班牙文|西班牙语|越南文|越南语|葡萄牙文|葡萄牙语|意大利文|意大利语).*$", string.Empty, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^(Please|Kindly).*$", string.Empty, RegexOptions.IgnoreCase);
            s = Regex.Replace(s, @"^(注意|说明)\b.*$", string.Empty, RegexOptions.IgnoreCase);
            // 越南语等目标语：常见“请提供纯正译文”类提示
            s = Regex.Replace(s, @"^Xin\s+cung\s+cấp\s+bản\s+dịch\s+tiếng\s+Việt\s+thuần\s+túy:?\s*$", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            s = Regex.Replace(s, @"^Vui\s+lòng\s+cung\s+cấp.*(bản\s+dịch|dịch).*", string.Empty, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            // 增强分行逻辑：根据主标点符号强制换行，确保译文与原文结构对齐
            // 中文标点：分号、句号、问号、感叹号（如果后面没有换行符）
            s = Regex.Replace(s, @"([；。！？])(?!\n)", "$1\n");
            // 英文标点：分号、问号、感叹号（后接空白符且非换行符）
            s = Regex.Replace(s, @"([;!?])([^\S\n]+)", "$1\n");

            // 应用用户自定义规则（逐条正则清洗）
            try
            {
                var rules = DocumentTranslator.Windows.TranslationRulesWindow.LoadGlobalRules();
                foreach (var pattern in rules)
                {
                    try { s = Regex.Replace(s, pattern, string.Empty, RegexOptions.Multiline | RegexOptions.IgnoreCase); }
                    catch { /* 忽略非法正则 */ }
                }
            }
            catch { }

            var lines = s.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var compact = new List<string>();
            foreach (var line in lines)
            {
                var t = line.Trim();
                if (t.Length == 0) continue;
                compact.Add(t);
            }
            return string.Join("\n", compact);
        }

        /// <summary>
        /// 检查文本是否已经包含翻译内容（自检过滤）
        /// </summary>
        private bool IsAlreadyTranslated(string text)
        {
            // 用户指示：无论文本是否被翻译，都要再走一遍翻译流程
            // 因此禁用所有"疑似已翻译"的跳过逻辑
            return false;
        }

        private bool ContainsChinese(string text)
        {
            foreach (char c in text)
            {
                // Basic CJK Unified Ideographs
                if (c >= 0x4e00 && c <= 0x9fff) return true;
                // CJK Symbols and Punctuation (e.g., 。 、 【 】)
                if (c >= 0x3000 && c <= 0x303f) return true;
                // Full-width characters (e.g., ， ： ； ！ ？)
                if (c >= 0xff00 && c <= 0xffef) return true;
                // CJK Unified Ideographs Extension A
                if (c >= 0x3400 && c <= 0x4dbf) return true;
            }
            return false;
        }

        private bool ContainsEnglish(string text)
        {
            foreach (char c in text)
            {
                if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) return true;
            }
            return false;
        }

        /// <summary>
        /// 检查文本是否包含混合语言（带阈值判断）
        /// 只有当英文内容占比超过阈值时，才认为可能已翻译
        /// </summary>
        private bool ContainsMixedLanguagesWithThreshold(string text, double englishRatioThreshold)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            int chineseCount = 0;
            int englishCount = 0;
            int totalChars = 0;

            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c) && !char.IsPunctuation(c))
                {
                    totalChars++;
                    if (c >= 0x4e00 && c <= 0x9fff)
                    {
                        chineseCount++;
                    }
                    else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                    {
                        englishCount++;
                    }
                }
            }

            if (totalChars == 0) return false;

            // 计算英文占比
            double englishRatio = (double)englishCount / totalChars;

            // 只有当英文占比超过阈值时，才认为可能已翻译
            // 这样可以避免将包含少量英文单位、缩写的中文原文误判为已翻译
            return englishRatio > englishRatioThreshold;
        }

        /// <summary>
        /// 检查文本是否包含混合语言
        /// </summary>
        private bool ContainsMixedLanguages(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            bool hasChinese = false;
            bool hasEnglish = false;

            foreach (char c in text)
            {
                if (c >= 0x4e00 && c <= 0x9fff) // 中文字符范围
                {
                    hasChinese = true;
                }
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    hasEnglish = true;
                }

                if (hasChinese && hasEnglish)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 检查文本是否主要是中文
        /// </summary>
        private bool IsChineseText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            int chineseCount = 0;
            int totalChars = 0;

            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c) && !char.IsPunctuation(c))
                {
                    totalChars++;
                    if (c >= 0x4e00 && c <= 0x9fff)
                    {
                        chineseCount++;
                    }
                }
            }

            return totalChars > 0 && (double)chineseCount / totalChars > 0.5;
        }

        /// <summary>
        /// 检查文本是否主要是英文
        /// </summary>
        private bool IsEnglishText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            int englishCount = 0;
            int totalChars = 0;

            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c) && !char.IsPunctuation(c))
                {
                    totalChars++;
                    if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                    {
                        englishCount++;
                    }
                }
            }

            return totalChars > 0 && (double)englishCount / totalChars > 0.5;
        }
        /// <summary>
        /// 检查文本是否包含原语种字符
        /// </summary>
        private bool ContainsSourceLanguage(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            
            // 如果是中译外，检查是否包含中文
            if (_isCnToForeign)
            {
                return ContainsChinese(text);
            }
            
            // 如果是外译中，检查是否包含英文（或其他拉丁字符）
            return ContainsEnglish(text);
        }

        /// <summary>
        /// 判断表格单元格文本是否应跳过翻译（空白/纯数字/纯编码样式/纯目标语种）
        /// </summary>
        private bool ShouldSkipCellText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return true;
            var t = text.Trim();

            // 1. 纯数字（包括小数、负数、百分号等），但不包含单位
            // 检查是否包含单位符号（kg, g, m, cm, mm, L, mL, %等）
            bool hasUnit = Regex.IsMatch(t, @"(kg|g|mg|µg|m|cm|mm|km|L|mL|µL|mol|mmol|µmol|°C|°F|K|Pa|kPa|MPa|bar|atm|%|ppm|ppb|pH|Hz|kHz|MHz|GHz|W|kW|MW|V|kV|A|mA|µA|Ω|kΩ|MΩ|s|min|h|d|wk|mo|yr|B|KB|MB|GB|TB|bps|kbps|Mbps|Gbps)", RegexOptions.IgnoreCase);
            
            if (Regex.IsMatch(t, @"^[\d\.,\-\+\s]+$") && !hasUnit)
                return true;

            // 2. 版本号（可带前缀v/V）或形如 1.2.3
            if (Regex.IsMatch(t, @"^[Vv]?\d+(\.\d+){1,3}$"))
                return true;

            // 3. 核心逻辑：如果文本不包含原语种字符，则跳过
            // 这意味着纯英文（在中译外时）或纯中文（在外译中时）会被跳过
            // 这符合用户"若为纯英文文本则跳过翻译，若不是则需要翻译"的要求
            if (!ContainsSourceLanguage(t))
                return true;

            return false;
        }



        /// <summary>
        /// 处理文本框（SdtBlock）
        /// </summary>
        private async Task ProcessSdtBlocksAsync(WordprocessingDocument document, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock)
        {
            try
            {
                var mainPart = document.MainDocumentPart;
                if (mainPart == null) return;

                var sdtBlocks = mainPart.Document.Body.Descendants<SdtBlock>().ToList();
                if (!sdtBlocks.Any())
                {
                    _logger.LogInformation("文档中未找到文本框（SdtBlock）");
                    return;
                }

                _logger.LogInformation($"开始处理 {sdtBlocks.Count} 个文本框");

                for (int i = 0; i < sdtBlocks.Count; i++)
                {
                    try
                    {
                        var sdtBlock = sdtBlocks[i];
                        var paragraphs = sdtBlock.Descendants<Paragraph>().ToList();

                        foreach (var paragraph in paragraphs)
                        {
                            await ProcessParagraphAsync(paragraph, terminology);
                        }
                        
                        // 处理文本框中的表格
                        var tables = sdtBlock.Descendants<Table>().ToList();
                        foreach (var table in tables)
                        {
                            await ProcessTableAsync(table, terminology, csvWriter, csvLock);
                        }

                        _logger.LogInformation($"已处理文本框 {i + 1}/{sdtBlocks.Count}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"处理文本框失败，跳过该文本框");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "处理文本框时发生错误");
            }
        }

        /// <summary>
        /// 处理页眉页脚
        /// </summary>
        private async Task ProcessHeadersAndFootersAsync(WordprocessingDocument document, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock)
        {
            try
            {
                var mainPart = document.MainDocumentPart;
                if (mainPart == null) return;

                int headerCount = 0;
                int footerCount = 0;
                int headerTableCount = 0;
                int footerTableCount = 0;

                // 处理页眉
                foreach (var headerPart in mainPart.HeaderParts)
                {
                    try
                    {
                        var paragraphs = headerPart.Header.Descendants<Paragraph>().ToList();
                        foreach (var paragraph in paragraphs)
                        {
                            await ProcessParagraphAsync(paragraph, terminology);
                        }
                        
                        // 处理页眉中的表格
                        var headerTables = headerPart.Header.Descendants<Table>().ToList();
                        foreach (var table in headerTables)
                        {
                            await ProcessTableAsync(table, terminology, csvWriter, csvLock);
                            headerTableCount++;
                        }
                        
                        headerCount++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "处理页眉失败");
                    }
                }

                // 处理页脚
                foreach (var footerPart in mainPart.FooterParts)
                {
                    try
                    {
                        var paragraphs = footerPart.Footer.Descendants<Paragraph>().ToList();
                        foreach (var paragraph in paragraphs)
                        {
                            await ProcessParagraphAsync(paragraph, terminology);
                        }
                        
                        // 处理页脚中的表格
                        var footerTables = footerPart.Footer.Descendants<Table>().ToList();
                        foreach (var table in footerTables)
                        {
                            await ProcessTableAsync(table, terminology, csvWriter, csvLock);
                            footerTableCount++;
                        }
                        
                        footerCount++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "处理页脚失败");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "处理页眉页脚失败");
            }
        }

        /// <summary>
        /// 检查文本是否应该跳过
        /// </summary>
        private bool ShouldSkipText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return true;

            var t = text.Trim();

            // 跳过纯数字
            if (Regex.IsMatch(t, @"^[\d\.,\-\+\s]+$"))
                return true;

            // 跳过版本号
            if (Regex.IsMatch(t, @"^[Vv]?\d+(\.\d+){1,3}$"))
                return true;

            // 跳过纯编码（短文本、无中文、只包含字母数字和特殊字符）
            if (t.Length <= 20 && !Regex.IsMatch(t, @"[\u4e00-\u9fff]") && Regex.IsMatch(t, @"^[A-Za-z0-9_\-\.]+$"))
                return true;

            // 跳过包含公式的文本
            if (t.Contains("$") || t.Contains("\\begin") || t.Contains("\\(") || t.Contains("\\["))
                return true;

            return false;
        }

        /// <summary>
        /// 获取语言名称
        /// </summary>
        private string GetLanguageName(string langCode)
        {
            var languageNames = new Dictionary<string, string>
            {
                ["zh"] = "中文",
                ["en"] = "英文",
                ["ja"] = "日语",
                ["ko"] = "韩语",
                ["fr"] = "法语",
                ["de"] = "德语",
                ["es"] = "西班牙语",
                ["it"] = "意大利语",
                ["ru"] = "俄语",
                ["pt"] = "葡萄牙语",
                ["vi"] = "越南语",
                ["other"] = "其他"
            };

            return languageNames.TryGetValue(langCode, out var name) ? name : langCode;
        }

        /// <summary>
        /// 更新表格单元格文本，保留段落结构
        /// </summary>
        private void UpdateTableCellTextWithStructure(List<Paragraph> paragraphs, string originalText, string translatedText)
        {
            if (!paragraphs.Any()) return;

            // 先清洗并应用分行逻辑，确保译文结构正确
            var cleanText = CleanTranslationText(translatedText);
            var translatedLines = cleanText.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int transIndex = 0;

            if (_outputFormat == "bilingual")
            {
                foreach (var paragraph in paragraphs)
                {
                    var pText = GetParagraphText(paragraph);
                    if (string.IsNullOrWhiteSpace(pText)) continue;

                    var pLines = pText.Replace("\r\n", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    var paraTransLines = new List<string>();

                    for (int k = 0; k < pLines.Length; k++)
                    {
                        if (transIndex < translatedLines.Length)
                        {
                            paraTransLines.Add(translatedLines[transIndex]);
                            transIndex++;
                        }
                    }

                    // 如果是最后一个段落，且仍有剩余译文，全部追加到此段落
                    if (paragraph == paragraphs.Last() && transIndex < translatedLines.Length)
                    {
                        while (transIndex < translatedLines.Length)
                        {
                            paraTransLines.Add(translatedLines[transIndex++]);
                        }
                    }

                    if (paraTransLines.Count > 0)
                    {
                        var combinedTrans = string.Join("\n", paraTransLines);
                        // CleanTranslationTextWithOriginal 会再次清洗，但由于CleanTranslationText已幂等，无副作用
                        var clean = CleanTranslationTextWithOriginal(pText, combinedTrans);
                        var cleanLines = clean.Replace("\r\n", "\n").Split('\n');

                        var runs = paragraph.Descendants<Run>().ToList();
                        var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

                        paragraph.AppendChild(new Run(new Break()));

                        for (int i = 0; i < cleanLines.Length; i++)
                        {
                            if (i > 0) paragraph.AppendChild(new Run(new Break()));
                            var run = new Run(new Text(cleanLines[i]));
                            if (baseRunProps != null) run.RunProperties = (RunProperties)baseRunProps.CloneNode(true);
                            else run.RunProperties = new RunProperties();

                            if (run.RunProperties == null) run.RunProperties = new RunProperties();
                            run.RunProperties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "0066CC" };
                            paragraph.AppendChild(run);
                        }
                    }
                }
            }
            else
            {
                var paragraphLineCounts = new List<int>();
                foreach (var p in paragraphs)
                {
                    var txt = GetParagraphText(p);
                    var cnt = string.IsNullOrWhiteSpace(txt) ? 0 : txt.Replace("\r\n", "\n").Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries).Length;
                    paragraphLineCounts.Add(cnt);
                }

                int currentPIndex = 0;
                foreach (var p in paragraphs)
                {
                    var runs = p.Descendants<Run>().ToList();
                    var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

                    int count = paragraphLineCounts[currentPIndex];
                    if (count == 0 && transIndex < translatedLines.Length) count = 1;

                    var paraTransLines = new List<string>();
                    for (int k = 0; k < count; k++)
                    {
                        if (transIndex < translatedLines.Length)
                        {
                            paraTransLines.Add(translatedLines[transIndex]);
                            transIndex++;
                        }
                    }

                    // 如果是最后一个段落，且仍有剩余译文，全部追加到此段落
                    if (p == paragraphs.Last() && transIndex < translatedLines.Length)
                    {
                        while (transIndex < translatedLines.Length)
                        {
                            paraTransLines.Add(translatedLines[transIndex++]);
                        }
                    }

                    // 如果该段落没有分配到译文，且原文不为空，保留原文（避免清空单元格）
                    // 仅当原文有内容（count > 0）时才执行保留逻辑
                    bool hasTranslation = paraTransLines.Count > 0;
                    
                    if (hasTranslation)
                    {
                        foreach (var run in runs) run.Remove();

                        var combined = string.Join("\n", paraTransLines);
                        var clean = CleanTranslationTextWithOriginal("", combined);
                        
                        // 如果清洗后为空，但原文不为空，回退到保留原文
                        if (string.IsNullOrWhiteSpace(clean) && count > 0)
                        {
                            // 恢复删除的Runs? 比较困难，不如直接不删除。
                            // 由于上面已经删除了，这里需要重新添加原文?
                            // 不，应该在删除前判断。
                            // 但逻辑顺序是：先收集译文，再决定。
                            
                            // 这里我们简单处理：如果清洗后为空，则视为"无有效译文"，
                            // 如果此时原文有内容，我们应该显示原文或者一个占位符，而不是空白。
                            // 但"CleanTranslationText"通常只清洗垃圾。如果变成了空，说明全是垃圾。
                            // 这种情况下，也许保持空白是正确的？
                            // 不，用户反馈"回写失败"，意味着他们期望看到东西。
                            // 如果是"熔料" -> ""，这绝对是错的。
                            // 所以如果 clean 为空，且 paraTransLines 不为空，说明被过度清洗了。
                            // 此时我们使用原始的 combined（未清洗的译文）或者原文。
                            
                            if (!string.IsNullOrWhiteSpace(combined))
                            {
                                clean = combined; // 放弃清洗，直接使用译文
                            }
                            else
                            {
                                // 译文本身就是空的
                            }
                        }

                        if (!string.IsNullOrEmpty(clean))
                        {
                            var cleanLines = clean.Replace("\r\n", "\n").Split('\n');
                            for (int i = 0; i < cleanLines.Length; i++)
                            {
                                if (i > 0) p.AppendChild(new Run(new Break()));
                                var run = new Run(new Text(cleanLines[i]));
                                if (baseRunProps != null) run.RunProperties = (RunProperties)baseRunProps.CloneNode(true);
                                p.AppendChild(run);
                            }
                        }
                        else if (count > 0)
                        {
                            // 译文为空，且原文有内容：回退逻辑
                            // 重新读取该段落的原文并写入（因为Runs已被移除）
                            // 简单的做法：写回 "[翻译为空]" 或者 尝试恢复
                            // 由于我们有 originalText，但那是整个单元格的。
                            // 这里的 p 是单元格中的一个段落。
                            // 更好的做法是：在删除Runs之前检查。
                        }
                    }
                    else
                    {
                         // 没有分配到译文：不删除Runs，保留原文
                         // do nothing (runs are NOT removed yet? No, I moved the removal logic inside hasTranslation check? No, I need to adjust the code structure.)
                    }
                    currentPIndex++;
                }
            }
        }

        /// <summary>
        /// 更新表格单元格文本，避免重复翻译
        /// </summary>
        private void UpdateTableCellText(List<Paragraph> paragraphs, string originalText, string translatedText)
        {
            if (!paragraphs.Any()) return;

            // 找到第一个有内容的段落
            var targetParagraph = paragraphs.FirstOrDefault(p => !string.IsNullOrWhiteSpace(p.InnerText)) ?? paragraphs[0];
            var clean = CleanTranslationTextWithOriginal(originalText, translatedText);

            // 捕获原始段落的RunProperties
            var runs = targetParagraph.Descendants<Run>().ToList();
            var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

            if (_outputFormat == "bilingual")
            {
                // 保留原段落，只在第一个非空段落后追加译文
                targetParagraph.AppendChild(new Run(new Break()));
                var translatedRun = new Run(new Text(clean));
                if (baseRunProps != null)
                {
                    translatedRun.RunProperties = baseRunProps;
                }
                else
                {
                    if (translatedRun.RunProperties == null)
                        translatedRun.RunProperties = new RunProperties();
                }
                translatedRun.RunProperties.Color = new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "0066CC" };
                targetParagraph.AppendChild(translatedRun);
            }
            else
            {
                // 仅翻译结果模式：清空所有段落，只在第一个段落写入译文
                foreach (var p in paragraphs)
                {
                    var pRuns = p.Descendants<Run>().ToList();
                    foreach (var run in pRuns) run.Remove();
                }
                var newRun = new Run(new Text(clean));
                if (baseRunProps != null)
                {
                    newRun.RunProperties = baseRunProps;
                }
                targetParagraph.AppendChild(newRun);
            }
        }


        /// <summary>
        /// 使用翻译模型检测语言
        /// </summary>
        private async Task<LanguageInfo> DetectLanguageWithModelAsync(string text)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(text))
                    return new LanguageInfo("other", "其他", 0.0);

                var translator = _translationService.CurrentTranslator;
                if (translator == null)
                {
                    _logger.LogWarning("翻译器未初始化，使用Unicode范围检测");
                    return LanguageDetector.Detect(text);
                }

                // 使用翻译模型的语言检测功能
                var detectedCode = await translator.DetectLanguageAsync(text);
                var langName = GetLanguageName(detectedCode);

                return new LanguageInfo(detectedCode, langName, 1.0);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "使用模型检测语言失败，使用Unicode范围检测");
                return LanguageDetector.Detect(text);
            }
        }



        /// <summary>
        /// 文档元素类型枚举
        /// </summary>
        public enum DocumentElementType
        {
            Title,
            Paragraph,
            Table,
            TableCell,
            Header,
            Footer,
            Footnote,
            Endnote,
            Image,
            Shape,
            Textbox
        }

        /// <summary>
        /// 文本格式信息类
        /// </summary>
        public class TextFormatInfo
        {
            public string FontName { get; set; }
            public double FontSize { get; set; }
            public bool IsBold { get; set; }
            public bool IsItalic { get; set; }
            public bool IsUnderline { get; set; }
            public string Color { get; set; }
            public string HighlightColor { get; set; }
            public string Alignment { get; set; }

            public TextFormatInfo()
            {
                FontName = "宋体";
                FontSize = 10.5;
                IsBold = false;
                IsItalic = false;
                IsUnderline = false;
                Color = "000000";
                HighlightColor = null;
                Alignment = "left";
            }

            public static TextFormatInfo FromRunProperties(RunProperties runProps)
            {
                var format = new TextFormatInfo();

                if (runProps != null)
            {
                var runFonts = runProps.RunFonts;
                if (runFonts != null)
                {
                    format.FontName = runFonts.Ascii ?? runFonts.HighAnsi ?? "宋体";
                }

                var fontSizeValue = runProps.FontSize?.Val?.Value;
                if (!string.IsNullOrEmpty(fontSizeValue) && int.TryParse(fontSizeValue, out var fontSize))
                {
                    format.FontSize = fontSize / 2.0;
                }

                format.IsBold = runProps.Bold != null;
                format.IsItalic = runProps.Italic != null;
                format.IsUnderline = runProps.Underline != null;

                var color = runProps.Color?.Val?.Value;
                if (color != null)
                {
                    format.Color = color;
                }

                var highlightValue = runProps.Highlight?.Val?.Value;
                if (highlightValue != null)
                {
                    format.HighlightColor = highlightValue.ToString();
                }
            }

                return format;
            }

            public RunProperties ToRunProperties()
            {
                var runProps = new RunProperties();

                if (!string.IsNullOrEmpty(FontName))
                {
                    var runFonts = new RunFonts();
                    runFonts.Ascii = FontName;
                    runFonts.HighAnsi = FontName;
                    runProps.AppendChild(runFonts);
                }

                runProps.FontSize = new FontSize { Val = new StringValue((FontSize * 2).ToString()) };

                if (IsBold)
                {
                    runProps.Bold = new Bold();
                }

                if (IsItalic)
                {
                    runProps.Italic = new Italic();
                }

                if (IsUnderline)
                {
                    runProps.Underline = new Underline();
                }

                if (!string.IsNullOrEmpty(Color))
                {
                    runProps.Color = new DocumentFormat.OpenXml.Wordprocessing.Color { Val = Color };
                }

                if (!string.IsNullOrEmpty(HighlightColor))
                {
                    if (Enum.TryParse<HighlightColorValues>(HighlightColor, true, out var highlightValue))
                    {
                        runProps.Highlight = new Highlight { Val = highlightValue };
                    }
                }

                return runProps;
            }
        }

        /// <summary>
        /// 文档结构信息类
        /// </summary>
        public class DocumentStructureInfo
        {
            public DocumentElementType ElementType { get; set; }
            public string Text { get; set; }
            public LanguageInfo Language { get; set; }
            public TextFormatInfo Format { get; set; }
            public int Position { get; set; }
            public string Location { get; set; }

            public DocumentStructureInfo()
            {
                Text = string.Empty;
                Language = new LanguageInfo("other", "其他", 0.0);
                Format = new TextFormatInfo();
                Position = 0;
                Location = string.Empty;
            }
        }

        /// <summary>
        /// Word文档结构解析器
        /// </summary>
        public class WordDocumentStructure
        {
            private readonly ILogger _logger;

            public WordDocumentStructure(ILogger logger)
            {
                _logger = logger;
            }

            public List<DocumentStructureInfo> ParseDocument(WordprocessingDocument document)
            {
                var structure = new List<DocumentStructureInfo>();
                var body = document.MainDocumentPart.Document.Body;

                ParseParagraphs(body, structure, "正文");
                ParseTables(body, structure, "正文");
                ParseHeaders(document, structure);
                ParseFooters(document, structure);
                ParseFootnotes(document, structure);
                ParseEndnotes(document, structure);

                return structure;
            }

            private void ParseParagraphs(Body body, List<DocumentStructureInfo> structure, string location)
            {
                var paragraphs = body.Descendants<Paragraph>().ToList();
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var para = paragraphs[i];
                    var text = para.InnerText?.Trim();

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    var runs = para.Descendants<Run>().ToList();
                    var runProps = runs.FirstOrDefault()?.RunProperties;

                    var info = new DocumentStructureInfo
                    {
                        ElementType = DocumentElementType.Paragraph,
                        Text = text,
                        Language = LanguageDetector.Detect(text),
                        Format = TextFormatInfo.FromRunProperties(runProps),
                        Position = i,
                        Location = $"{location}_段落{i + 1}"
                    };

                    structure.Add(info);
                }
            }

            private void ParseTables(Body body, List<DocumentStructureInfo> structure, string location)
            {
                var tables = body.Descendants<Table>().ToList();
                for (int tableIndex = 0; tableIndex < tables.Count; tableIndex++)
                {
                    var table = tables[tableIndex];
                    var rows = table.Descendants<TableRow>().ToList();

                    for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                    {
                        var cells = rows[rowIndex].Descendants<TableCell>().ToList();

                        for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                        {
                            var cell = cells[colIndex];
                            var paragraphs = cell.Descendants<Paragraph>().ToList();

                            foreach (var para in paragraphs)
                            {
                                var text = para.InnerText?.Trim();
                                if (string.IsNullOrWhiteSpace(text))
                                    continue;

                                var runs = para.Descendants<Run>().ToList();
                                var runProps = runs.FirstOrDefault()?.RunProperties;

                                var info = new DocumentStructureInfo
                                {
                                    ElementType = DocumentElementType.TableCell,
                                    Text = text,
                                    Language = LanguageDetector.Detect(text),
                                    Format = TextFormatInfo.FromRunProperties(runProps),
                                    Position = tableIndex * 1000 + rowIndex * 100 + colIndex,
                                    Location = $"{location}_表格{tableIndex + 1}_行{rowIndex + 1}_列{colIndex + 1}"
                                };

                                structure.Add(info);
                            }
                        }
                    }
                }
            }

            private void ParseHeaders(WordprocessingDocument document, List<DocumentStructureInfo> structure)
            {
                var headers = document.MainDocumentPart.HeaderParts;
                int headerIndex = 0;

                foreach (var headerPart in headers)
                {
                    var paragraphs = headerPart.Header.Descendants<Paragraph>().ToList();

                    foreach (var para in paragraphs)
                    {
                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        var runs = para.Descendants<Run>().ToList();
                        var runProps = runs.FirstOrDefault()?.RunProperties;

                        var info = new DocumentStructureInfo
                        {
                            ElementType = DocumentElementType.Header,
                            Text = text,
                            Language = LanguageDetector.Detect(text),
                            Format = TextFormatInfo.FromRunProperties(runProps),
                            Position = headerIndex * 10000 + structure.Count,
                            Location = $"页眉{headerIndex + 1}_段落"
                        };

                        structure.Add(info);
                    }
                    headerIndex++;
                }
            }

            private void ParseFooters(WordprocessingDocument document, List<DocumentStructureInfo> structure)
            {
                var footers = document.MainDocumentPart.FooterParts;
                int footerIndex = 0;

                foreach (var footerPart in footers)
                {
                    var paragraphs = footerPart.Footer.Descendants<Paragraph>().ToList();

                    foreach (var para in paragraphs)
                    {
                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text))
                            continue;

                        var runs = para.Descendants<Run>().ToList();
                        var runProps = runs.FirstOrDefault()?.RunProperties;

                        var info = new DocumentStructureInfo
                        {
                            ElementType = DocumentElementType.Footer,
                            Text = text,
                            Language = LanguageDetector.Detect(text),
                            Format = TextFormatInfo.FromRunProperties(runProps),
                            Position = footerIndex * 10000 + structure.Count,
                            Location = $"页脚{footerIndex + 1}_段落"
                        };

                        structure.Add(info);
                    }
                    footerIndex++;
                }
            }

            private void ParseFootnotes(WordprocessingDocument document, List<DocumentStructureInfo> structure)
            {
                var footnotesPart = document.MainDocumentPart.FootnotesPart;
                if (footnotesPart == null) return;

                var footnotes = footnotesPart.Footnotes.Descendants<Footnote>().ToList();
                foreach (var footnote in footnotes)
                {
                    var paragraphs = footnote.Descendants<Paragraph>().ToList();
                    foreach (var para in paragraphs)
                    {
                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        var runs = para.Descendants<Run>().ToList();
                        var runProps = runs.FirstOrDefault()?.RunProperties;

                        var info = new DocumentStructureInfo
                        {
                            ElementType = DocumentElementType.Footnote,
                            Text = text,
                            Language = LanguageDetector.Detect(text),
                            Format = TextFormatInfo.FromRunProperties(runProps),
                            Position = structure.Count,
                            Location = $"脚注_ID{footnote.Id}"
                        };
                        structure.Add(info);
                    }
                }
            }

            private void ParseEndnotes(WordprocessingDocument document, List<DocumentStructureInfo> structure)
            {
                var endnotesPart = document.MainDocumentPart.EndnotesPart;
                if (endnotesPart == null) return;

                var endnotes = endnotesPart.Endnotes.Descendants<Endnote>().ToList();
                foreach (var endnote in endnotes)
                {
                    var paragraphs = endnote.Descendants<Paragraph>().ToList();
                    foreach (var para in paragraphs)
                    {
                        var text = para.InnerText?.Trim();
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        var runs = para.Descendants<Run>().ToList();
                        var runProps = runs.FirstOrDefault()?.RunProperties;

                        var info = new DocumentStructureInfo
                        {
                            ElementType = DocumentElementType.Endnote,
                            Text = text,
                            Language = LanguageDetector.Detect(text),
                            Format = TextFormatInfo.FromRunProperties(runProps),
                            Position = structure.Count,
                            Location = $"尾注_ID{endnote.Id}"
                        };
                        structure.Add(info);
                    }
                }
            }
        }
    }
}
