using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Threading;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using System.Text;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// Excel文档处理器，负责处理Excel文档的翻译
    /// </summary>
    public class ExcelProcessor
    {
        private readonly TranslationService _translationService;
        private readonly ILogger<ExcelProcessor> _logger;
        private readonly TermExtractor _termExtractor;
        private bool _useTerminology = true;
        private bool _preprocessTerms = true;
        private string _sourceLang = "zh";
        private string _targetLang = "en";
        private string _terminologyLanguageName = "英语"; // 用于术语库检索的语言键（可由名称/代码归一化）
        private bool _isCnToForeign = true;
        private string _outputFormat = "bilingual"; // 输出格式：bilingual（双语对照）或 translation_only（仅翻译结果）
        private Action<double, string> _progressCallback;
        private int _retryCount = 1;
        private int _retryDelay = 1000; // 毫秒
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private const double _minProgressUpdateInterval = 100; // 最小进度更新间隔100ms
        private bool _useBOMTranslation = false; // BOM表专属翻译功能开关

        /// <summary>
        /// 是否使用BOM表专属翻译功能
        /// </summary>
        public bool UseBOMTranslation
        {
            get { return _useBOMTranslation; }
            set { _useBOMTranslation = value; }
        }

        /// <summary>
        /// 解析术语库使用的语言键：
        /// - 中文→外语：直接使用 UI 传入的目标语言名称（如"英语"）
        /// - 外语→中文：根据源语言代码映射到术语库顶层键（如 en→"英语"）
        /// </summary>
        private string ResolveTerminologyLanguageName(string targetLanguageName, string targetLangCode)
        {
            if (_isCnToForeign)
                return targetLanguageName;

            // 外语→中文：根据源语言代码映射
            var codeToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["en"] = "英语",
                ["ja"] = "日语",
                ["ko"] = "韩语",
                ["fr"] = "法语",
                ["de"] = "德语",
                ["es"] = "西班牙语",
                ["it"] = "意大利语",
                ["ru"] = "俄语",
            };
            if (codeToName.TryGetValue(_sourceLang ?? string.Empty, out var name))
                return name;

            // 兜底：返回传入名称以避免空值
            return targetLanguageName;
        }

        public ExcelProcessor(TranslationService translationService, ILogger<ExcelProcessor> logger)
        {
            _translationService = translationService ?? throw new ArgumentNullException(nameof(translationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _termExtractor = new TermExtractor();

            // 设置EPPlus许可证上下文为非商业用途
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// 设置进度回调函数
        /// </summary>
        public void SetProgressCallback(Action<double, string> callback)
        {
            _progressCallback = callback;
        }

        /// <summary>
        /// 设置翻译选项
        /// </summary>
        public void SetTranslationOptions(bool useTerminology = true, bool preprocessTerms = true,
            string sourceLang = "zh", string targetLang = "en", string outputFormat = "bilingual")
        {
            _useTerminology = useTerminology;
            _preprocessTerms = preprocessTerms;
            _sourceLang = sourceLang;
            _targetLang = targetLang;
            _outputFormat = outputFormat;
            _isCnToForeign = sourceLang == "zh";
        }

        /// <summary>
        /// 处理Excel文档
        /// </summary>
        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguageName)
        {
            return await ProcessDocumentAsync(filePath, targetLanguageName, null);
        }

        /// <summary>
        /// 处理Excel文档（带术语库）
        /// </summary>
        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguageName, Dictionary<string, string> terminology)
        {
            _logger.LogInformation($"开始处理Excel文档: {filePath}");
            
            try
            {
                // 解析目标语言名称和代码
                _terminologyLanguageName = ResolveTerminologyLanguageName(targetLanguageName, _targetLang);
                _logger.LogInformation($"解析术语库语言键: {_terminologyLanguageName}");

                // 加载术语库
                var finalTerminology = terminology ?? new Dictionary<string, string>();
                if (_useTerminology && finalTerminology.Count == 0)
                {
                    var terms = _termExtractor.GetTermsForLanguage(_terminologyLanguageName);
                    if (terms != null && terms.Count > 0)
                    {
                        finalTerminology = terms;
                        _logger.LogInformation($"已加载术语库（{_terminologyLanguageName}）：{terms.Count} 条");
                    }
                }

                // 生成输出文件路径
                string outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输出");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);
                string outputFilePath = Path.Combine(outputDir, $"{fileName}_带翻译_{DateTime.Now:yyyyMMddHHmmss}{extension}");

                // 复制原始文件到输出路径
                File.Copy(filePath, outputFilePath, true);
                _logger.LogInformation($"已创建输出文件: {outputFilePath}");

                // 处理Excel文档
                await ProcessExcelDocumentAsync(outputFilePath, finalTerminology);

                _logger.LogInformation($"Excel文档处理完成: {outputFilePath}");
                return outputFilePath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Excel文档处理失败");
                throw;
            }
        }

        /// <summary>
        /// 处理Excel文档
        /// </summary>
        private async Task ProcessExcelDocumentAsync(string filePath, Dictionary<string, string> terminology)
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheets = package.Workbook.Worksheets;
            var totalWorksheets = worksheets.Count;

            _logger.LogInformation($"开始处理 {totalWorksheets} 个工作表");

            // 准备 CSV 路径
            var csvDir = Path.Combine(Path.GetDirectoryName(filePath)!, "日志");
            if (!Directory.Exists(csvDir)) Directory.CreateDirectory(csvDir);
            var csvPath = Path.Combine(csvDir, $"Excel翻译日志_{DateTime.Now:yyyyMMddHHmmss}.csv");
            _logger.LogInformation($"将记录预处理与翻译到CSV: {csvPath}");

            // 写 CSV 头
            using var csvStream = new FileStream(csvPath, FileMode.Create, FileAccess.Write, FileShare.Read);
            using var csvWriter = new StreamWriter(csvStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
            await csvWriter.WriteLineAsync("工作表,位置,原文,预处理后,译文,状态,长度对比");
            await csvWriter.FlushAsync();
            var csvLock = new object();

            // 计算总单元格数，用于进度计算
            int totalCells = 0;
            foreach (var worksheet in worksheets)
            {
                var dimension = worksheet.Dimension;
                if (dimension != null)
                {
                    int rows = dimension.End.Row - dimension.Start.Row + 1;
                    int cols = dimension.End.Column - dimension.Start.Column + 1;
                    totalCells += rows * cols;
                }
            }

            _logger.LogInformation($"总单元格数: {totalCells}");

            // 进度跟踪
            int processedCells = 0;
            object progressLock = new object();

            // 受控并发处理各工作表：翻译阶段并发
            var tasks = new List<Task>();
            for (int i = 0; i < totalWorksheets; i++)
            {
                var index = i;
                var worksheet = worksheets[index];
                tasks.Add(Task.Run(async () =>
                {
                    try
                    {
                        if (_useBOMTranslation)
                        {
                            await ProcessBOMWorksheetAsync(worksheet, terminology, csvWriter, csvLock, () =>
                            {
                                lock (progressLock)
                                {
                                    processedCells++;
                                    double progress = totalCells > 0 ? (double)processedCells / totalCells : 0;
                                    UpdateProgress(progress, $"处理工作表 {worksheet.Name}");
                                }
                            });
                        }
                        else
                        {
                            await ProcessWorksheetAsync(worksheet, terminology, csvWriter, csvLock, () =>
                            {
                                lock (progressLock)
                                {
                                    processedCells++;
                                    double progress = totalCells > 0 ? (double)processedCells / totalCells : 0;
                                    UpdateProgress(progress, $"处理工作表 {worksheet.Name}");
                                }
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, $"处理工作表 {worksheet.Name} 失败");
                    }
                }));
            }

            await Task.WhenAll(tasks);
            _logger.LogInformation($"所有工作表处理完成");
            UpdateProgress(1.0, "处理完成");

            // 保存更改
            await package.SaveAsync();
        }

        /// <summary>
        /// 处理普通工作表
        /// </summary>
        private async Task ProcessWorksheetAsync(ExcelWorksheet worksheet, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock, Action onCellProcessed)
        {
            _logger.LogInformation($"处理工作表: {worksheet.Name}");

            // 获取工作表的使用范围
            var dimension = worksheet.Dimension;
            if (dimension == null)
            {
                _logger.LogInformation($"工作表 {worksheet.Name} 为空，跳过处理");
                return;
            }

            int startRow = dimension.Start.Row;
            int endRow = dimension.End.Row;
            int startCol = dimension.Start.Column;
            int endCol = dimension.End.Column;

            _logger.LogInformation($"处理工作表 {worksheet.Name}，范围: {startRow}-{endRow} 行, {startCol}-{endCol} 列");

            // 预处理合并单元格索引（仅记录左上角）
            var mergedCellMap = new Dictionary<long, ExcelAddress>();
            if (worksheet.MergedCells != null)
            {
                foreach (var addressStr in worksheet.MergedCells)
                {
                    try 
                    {
                        var addr = new ExcelAddress(addressStr);
                        long key = ((long)addr.Start.Row << 32) | (uint)addr.Start.Column;
                        if (!mergedCellMap.ContainsKey(key))
                        {
                            mergedCellMap[key] = addr;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"解析合并单元格地址失败: {addressStr}, 错误: {ex.Message}");
                    }
                }
            }

            var batchSize = 1000; // 批次大小
            var cellContexts = new List<CellProcessingContext>();

            // 遍历所有单元格
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    // 1. 准备上下文（主线程读取）
                    var cell = worksheet.Cells[row, col];
                    
                    // 合并单元格处理
                    if (cell.Merge)
                    {
                        long key = ((long)row << 32) | (uint)col;
                        if (mergedCellMap.TryGetValue(key, out var mergedAddr))
                        {
                            // 是左上角，继续处理
                            if (TryCreateCellContext(worksheet, row, col, terminology, out var ctx))
                            {
                                ctx.IsMerged = true;
                                ctx.MergedRange = mergedAddr;
                                cellContexts.Add(ctx);
                            }
                            else
                            {
                                onCellProcessed?.Invoke();
                            }
                        }
                        else
                        {
                            // 是合并单元格但非左上角，跳过
                            onCellProcessed?.Invoke();
                        }
                    }
                    else
                    {
                        // 普通单元格
                        if (TryCreateCellContext(worksheet, row, col, terminology, out var ctx))
                        {
                            cellContexts.Add(ctx);
                        }
                        else
                        {
                            onCellProcessed?.Invoke();
                        }
                    }

                    // 2. 批次处理
                    if (cellContexts.Count >= batchSize)
                    {
                        await ProcessBatchAsync(cellContexts, worksheet.Name, csvWriter, csvLock);
                        foreach(var c in cellContexts) 
                        {
                            ApplyContextToWorksheet(worksheet, c);
                            onCellProcessed?.Invoke();
                        }
                        cellContexts.Clear();
                    }
                }
            }

            // 处理剩余批次
            if (cellContexts.Count > 0)
            {
                await ProcessBatchAsync(cellContexts, worksheet.Name, csvWriter, csvLock);
                foreach(var c in cellContexts) 
                {
                    ApplyContextToWorksheet(worksheet, c);
                    onCellProcessed?.Invoke();
                }
                cellContexts.Clear();
            }
        }

        private bool TryCreateCellContext(ExcelWorksheet worksheet, int row, int col, Dictionary<string, string> terminology, out CellProcessingContext context)
        {
            context = null;
            var cell = worksheet.Cells[row, col];
            var cellText = cell.Text ?? string.Empty;

            if (string.IsNullOrWhiteSpace(cellText)) return false;

            context = new CellProcessingContext
            {
                Row = row,
                Col = col,
                OriginalText = cellText,
                Terminology = terminology,
                IsFormula = !string.IsNullOrEmpty(cell.Formula)
            };

            // 跳过检查
            if (ShouldSkipCellText(cellText))
            {
                context.ShouldSkip = true;
                context.SkipReason = "skip_token";
                return true;
            }

            if (IsAlreadyTranslated(cellText))
            {
                context.ShouldSkip = true;
                context.SkipReason = "already_translated";
                return true;
            }

            return true;
        }

        private async Task ProcessBatchAsync(List<CellProcessingContext> contexts, string worksheetName, StreamWriter csvWriter, object csvLock)
        {
            var options = new ParallelOptions 
            { 
                MaxDegreeOfParallelism = _translationService.GetEffectiveMaxParallelism() 
            };

            await Parallel.ForEachAsync(contexts, options, async (ctx, ct) =>
            {
                if (ctx.ShouldSkip)
                {
                    if (ctx.SkipReason == "skip_token")
                    {
                         LogSkippedTranslation(ctx.OriginalText);
                    }
                    else if (ctx.SkipReason == "already_translated")
                    {
                        // 记录已翻译的日志
                         LogTranslationToCSV(csvWriter, csvLock, worksheetName, $"{ctx.Row},{ctx.Col}", ctx.OriginalText, ctx.OriginalText, ctx.OriginalText, "跳过翻译", "");
                    }
                    return;
                }

                // 翻译逻辑
                try 
                {
                     // 术语预处理等逻辑... 复用 ProcessCellAsync 中的逻辑
                     // 为了代码复用，我们将核心翻译逻辑提取到这里
                     
                    var normalizedText = NormalizeUnitsNearNumbers(ctx.OriginalText);
                    ctx.PreprocessedText = normalizedText;

                    // 术语处理
                    Dictionary<string, string> termsToUse = new Dictionary<string, string>(ctx.Terminology ?? new Dictionary<string, string>());
                    if (_useTerminology)
                    {
                        // 注意：TermExtractor 应该是线程安全的，如果不安全，这里会有问题
                        // 假设它是无状态的
                        var extracted = _termExtractor.ExtractRelevantTerms(ctx.OriginalText, _terminologyLanguageName, _sourceLang, _targetLang);
                        if (extracted != null && extracted.Any())
                        {
                            foreach (var kv in extracted)
                            {
                                if (!termsToUse.ContainsKey(kv.Key)) termsToUse[kv.Key] = kv.Value;
                            }
                            if (_preprocessTerms)
                            {
                                termsToUse = _termExtractor.PreprocessTerms(termsToUse).ToDictionary(k => k.Key, v => v.Value);
                                ctx.PreprocessedText = _termExtractor.ReplaceTermsInText(normalizedText, termsToUse, _isCnToForeign);
                            }
                        }
                    }

                    var terminologyForTranslator = (_useTerminology && !_preprocessTerms) ? termsToUse : null;

                    // 翻译
                    // 这里可以记录开始时间
                    var translated = await TranslateTextWithRetryAsync(ctx.PreprocessedText, terminologyForTranslator, ctx.OriginalText);
                    
                    if (!string.IsNullOrEmpty(translated))
                    {
                        ctx.TranslatedText = translated;
                        ctx.TranslationStatus = translated.Contains("[长度异常]") ? "长度异常" : "翻译成功";
                        
                        // 记录 CSV
                        LogTranslationToCSV(csvWriter, csvLock, worksheetName, $"{ctx.Row},{ctx.Col}", ctx.OriginalText, ctx.PreprocessedText, translated, ctx.TranslationStatus);
                    }
                    else
                    {
                        ctx.TranslationStatus = "翻译失败";
                        LogTranslationToCSV(csvWriter, csvLock, worksheetName, $"{ctx.Row},{ctx.Col}", ctx.OriginalText, ctx.PreprocessedText, "翻译失败", "翻译失败");
                    }
                }
                catch (Exception ex)
                {
                    ctx.TranslationStatus = "异常失败";
                     _logger.LogWarning(ex, $"翻译单元格 [{ctx.Row},{ctx.Col}] 失败");
                     LogTranslationToCSV(csvWriter, csvLock, worksheetName, $"{ctx.Row},{ctx.Col}", ctx.OriginalText, "异常失败", "异常失败", "异常失败");
                }
            });
        }

        private void ApplyContextToWorksheet(ExcelWorksheet worksheet, CellProcessingContext ctx)
        {
            if (ctx.ShouldSkip || string.IsNullOrEmpty(ctx.TranslatedText)) return;
            if (ctx.IsFormula) return; // 公式不回写

            ExcelRange targetRange;
            if (ctx.IsMerged && ctx.MergedRange != null)
            {
                targetRange = worksheet.Cells[ctx.MergedRange.Start.Row, ctx.MergedRange.Start.Column, ctx.MergedRange.End.Row, ctx.MergedRange.End.Column];
            }
            else
            {
                targetRange = worksheet.Cells[ctx.Row, ctx.Col];
            }

            UpdateCellText(targetRange, ctx.OriginalText, ctx.TranslatedText);
        }

        /// <summary>
        /// 处理BOM表工作表
        /// </summary>
        private async Task ProcessBOMWorksheetAsync(ExcelWorksheet worksheet, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock, Action onCellProcessed)
        {
            _logger.LogInformation($"处理BOM表工作表: {worksheet.Name}");

            // 获取工作表的使用范围
            var dimension = worksheet.Dimension;
            if (dimension == null)
            {
                _logger.LogInformation($"工作表 {worksheet.Name} 为空，跳过处理");
                return;
            }

            int startRow = dimension.Start.Row;
            int endRow = dimension.End.Row;
            int startCol = dimension.Start.Column;
            int endCol = dimension.End.Column;

            _logger.LogInformation($"处理工作表 {worksheet.Name}，范围: {startRow}-{endRow} 行, {startCol}-{endCol} 列");

            // 分析列特征
            var columnTypes = new Dictionary<int, string>();
            var columnUniqueValues = new Dictionary<int, HashSet<string>>();

            for (int col = startCol; col <= endCol; col++)
            {
                var uniqueValues = new HashSet<string>();
                bool isNumericColumn = true;
                bool hasDuplicates = false;

                for (int row = startRow; row <= endRow; row++)
                {
                    var cell = worksheet.Cells[row, col];
                    var cellText = cell.Text ?? string.Empty;
                    var trimmedText = cellText.Trim();

                    if (!string.IsNullOrWhiteSpace(trimmedText))
                    {
                        uniqueValues.Add(trimmedText);

                        // 检查是否为纯数字
                        if (isNumericColumn && !System.Text.RegularExpressions.Regex.IsMatch(trimmedText, @"^[0-9]+(\.[0-9]+)?%?$"))
                        {
                            isNumericColumn = false;
                        }
                    }
                }

                // 检查是否有重复值（唯一值数量远小于总行数）
                hasDuplicates = uniqueValues.Count < (endRow - (startRow + 2) + 1) * 0.5;

                // 确定列类型
                string columnType;
                if (isNumericColumn)
                {
                    columnType = "numeric"; // 纯数字列，跳过翻译
                }
                else if (hasDuplicates && uniqueValues.Count <= 20) // 重复值列，统一翻译
                {
                    columnType = "duplicate_content";
                }
                else
                {
                    columnType = "unique_content"; // 唯一内容列，逐行翻译
                }

                columnUniqueValues[col] = uniqueValues;
                columnTypes[col] = columnType;
                _logger.LogInformation($"列 {col} 特征: 类型={columnType}, 唯一值数量={uniqueValues.Count}");
            }

            // 先处理重复内容列（统一翻译后替换）
            var duplicateColumnTranslations = new Dictionary<int, Dictionary<string, string>>();
            for (int col = startCol; col <= endCol; col++)
            {
                if (columnTypes[col] == "duplicate_content")
                {
                    _logger.LogInformation($"正在处理重复内容列 {col}，将统一翻译 {columnUniqueValues[col].Count} 个唯一值");
                    var translations = new Dictionary<string, string>();

                    // 创建临时翻译任务列表
                    var tempTranslationTasks = new List<Task>();
                    var duplicateSemaphore = new System.Threading.SemaphoreSlim(Math.Min(20, 10));

                    foreach (var value in columnUniqueValues[col])
                    {
                        // 过滤掉不需要翻译的内容
                        if (!string.IsNullOrWhiteSpace(value) && !ShouldSkipCellText(value))
                        {
                            tempTranslationTasks.Add(Task.Run(async () =>
                            {
                                await duplicateSemaphore.WaitAsync();
                                try
                                {
                                    // 从获取到信号量开始计时（即开始发送给模型时）
                                    var translationStartTime = DateTime.Now;
                                    var translated = await TranslateTextWithRetryAsync(value, terminology, value);
                                    var translationDuration = DateTime.Now - translationStartTime;
                                    _logger.LogDebug($"重复值 '{value}' 翻译耗时: {translationDuration.TotalSeconds:F2}秒");
                                    
                                    if (!string.IsNullOrWhiteSpace(translated))
                                    {
                                        lock (translations)
                                        {
                                            translations[value] = translated;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogWarning(ex, $"翻译重复值 '{value}' 失败");
                                }
                                finally
                                {
                                    duplicateSemaphore.Release();
                                }
                            }));
                        }
                    }

                    // 等待所有重复值翻译完成
                    await Task.WhenAll(tempTranslationTasks);
                    duplicateSemaphore.Dispose();
                    duplicateColumnTranslations[col] = translations;

                    // 应用翻译结果到所有数据行
                    for (int row = startRow + 2; row <= endRow; row++)
                    {
                        var cell = worksheet.Cells[row, col];
                        var cellText = cell.Text ?? string.Empty;
                        if (translations.TryGetValue(cellText.Trim(), out var translatedText))
                        {
                            UpdateCellText(cell, cellText, translatedText);
                            LogTranslationToCSV(csvWriter, csvLock, worksheet.Name, $"{row},{col}", cellText, cellText, translatedText);
                        }
                        onCellProcessed?.Invoke();
                    }
                }
            }

            // 处理唯一内容列（逐行翻译）
            for (int row = startRow + 2; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    var columnType = columnTypes[col];
                    if (columnType == "unique_content")
                    {
                        var cell = worksheet.Cells[row, col];
                        var cellText = cell.Text ?? string.Empty;
                        // 过滤掉不需要翻译的内容
                        if (!string.IsNullOrWhiteSpace(cellText) && !ShouldSkipCellText(cellText))
                        {
                            await TranslateAndUpdateCellAsync(worksheet, row, col, cellText, terminology, csvWriter, csvLock);
                        }
                    }
                    onCellProcessed?.Invoke();
                }
            }
        }

        /// <summary>
        /// 处理单元格
        /// </summary>
        private async Task ProcessCellAsync(ExcelWorksheet worksheet, int row, int col, Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock)
        {
            // 用于读取/写入内容的"单一单元格"（左上角）
            var contentCell = worksheet.Cells[row, col];
            // 用于样式应用的区域（合并单元格时为整块区域，否则为单一单元格）
            ExcelRange styleRange = contentCell;

            // 合并单元格处理：仅在合并区域左上角单元格执行翻译与写入
            if (contentCell.Merge && worksheet.MergedCells != null && worksheet.MergedCells.Count > 0)
            {
                ExcelAddress mergedRange = null;
                foreach (var mergedAddress in worksheet.MergedCells)
                {
                    try
                    {
                        var rng = new ExcelAddress(mergedAddress);
                        if (row >= rng.Start.Row && row <= rng.End.Row && col >= rng.Start.Column && col <= rng.End.Column)
                        {
                            mergedRange = rng;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"解析合并单元格地址失败: {mergedAddress}");
                    }
                }
                if (mergedRange != null)
                {
                    // 不是左上角的单元格就跳过，避免重复处理
                    if (!(row == mergedRange.Start.Row && col == mergedRange.Start.Column))
                    {
                        _logger.LogDebug($"跳过合并单元格非左上角单元格: {row},{col}");
                        return;
                    }
                    styleRange = worksheet.Cells[mergedRange.Start.Row, mergedRange.Start.Column, mergedRange.End.Row, mergedRange.End.Column];
                    _logger.LogInformation($"处理合并单元格区域: {mergedRange.Address}");
                }
                else
                {
                    // 合并单元格标记存在但未找到对应的合并区域，仍按普通单元格处理
                    _logger.LogWarning($"单元格 {row},{col} 标记为合并单元格，但未找到对应的合并区域");
                }
            }

            // 使用显示文本进行抽取，避免因类型导致遗漏（RichText/格式化文本/换行等）
            var cellText = contentCell.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cellText))
                return;

            // 记录是否为公式单元格：公式结果可记录到CSV，但为避免破坏工作簿，不写回公式单元格
            bool isFormulaCell = !string.IsNullOrEmpty(contentCell.Formula);

            // 短标点/代码类文本跳过（如"/", "-", "N/A"等）
            if (ShouldSkipCellText(cellText))
                return;

            // 检查是否已经包含翻译内容（自检过滤）
            if (IsAlreadyTranslated(cellText))
            {
                lock (csvLock)
                {
                    string esc(string s)
                    {
                        if (s == null) return "";
                        s = s.Replace("\r\n", "\n").Replace("\r", "\n");
                        if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
                            return '"' + s.Replace("\"", "\"\"") + '"';
                        return s;
                    }
                    string loc = $"{worksheet.Name}!{row},{col}";
                    csvWriter.WriteLine($"{esc(worksheet.Name)},{esc(loc)},{esc(cellText)},{esc(cellText)},{esc(cellText)}");
                    csvWriter.Flush();
                }
                return;
            }

            // 准备 CSV 写入的基本信息
            string location = $"{worksheet.Name}!{row},{col}";
            string preprocessed = string.Empty;
            string finalTranslated = string.Empty;

            // 术语预处理：根据全局设置与目标语言，提取本单元格相关术语
            // 规范术语映射方向：确保键总是"源语言术语"，值为"目标语言术语"
            Dictionary<string, string> termsToUse = terminology ?? new Dictionary<string, string>();
            if (_useTerminology && termsToUse.Count > 0 && !_isCnToForeign)
            {
                var reversed = new Dictionary<string, string>(StringComparer.Ordinal);
                foreach (var kv in termsToUse)
                {
                    var src = kv.Value; // 外语
                    var dst = kv.Key;   // 中文
                    if (string.IsNullOrWhiteSpace(src)) continue;
                    if (!reversed.ContainsKey(src)) reversed[src] = dst;
                }
                termsToUse = reversed;
            }
            if (_useTerminology)
            {
                var extracted = _termExtractor.ExtractRelevantTerms(cellText, _terminologyLanguageName, _sourceLang, _targetLang);
                if (extracted != null && extracted.Any())
                {
                    var samplePairs = string.Join(", ", extracted.Take(3).Select(kv => $"{kv.Key}->{kv.Value}"));
                    _logger.LogInformation($"术语提取：匹配 {extracted.Count} 个术语，示例: {samplePairs}");

                    foreach (var kv in extracted)
                    {
                        if (!termsToUse.ContainsKey(kv.Key)) termsToUse[kv.Key] = kv.Value;
                    }

                    if (_preprocessTerms)
                    {
                        termsToUse = _termExtractor.PreprocessTerms(termsToUse).ToDictionary(k => k.Key, v => v.Value);
                        _logger.LogInformation($"术语预处理已启用：将优先匹配长术语（当前可用术语 {termsToUse.Count} 个）");
                    }
                }
                else
                {
                    if (_preprocessTerms)
                    {
                        var fullDict = _termExtractor.GetTermsForLanguage(_terminologyLanguageName);
                        if (fullDict != null && fullDict.Count > 0)
                        {
                            foreach (var kv in fullDict)
                            {
                                if (!termsToUse.ContainsKey(kv.Key)) termsToUse[kv.Key] = kv.Value;
                            }
                            _logger.LogInformation($"术语预处理：使用全量术语库（{fullDict.Count} 条）进行预替换");
                        }
                    }
                }
            }

            // 翻译文本（优先处理中文单位 → 英文单位）
            string translatedText;
            if (TryTranslateStandaloneUnit(cellText, out var unitTranslation))
            {
                translatedText = unitTranslation;
                preprocessed = cellText;
            }
            else
            {
                var normalized = NormalizeUnitsNearNumbers(cellText);

                // 术语预替换：像 Word 一样在翻译前对文本进行术语预处理（直接替换为目标术语）
                string toTranslate = normalized;
                if (_useTerminology && _preprocessTerms && termsToUse != null && termsToUse.Any())
                {
                    preprocessed = _termExtractor.ReplaceTermsInText(normalized, termsToUse, _isCnToForeign);
                    toTranslate = preprocessed;

                    var beforeSnippet = normalized.Substring(0, Math.Min(80, normalized.Length));
                    var afterSnippet = preprocessed.Substring(0, Math.Min(80, preprocessed.Length));
                    _logger.LogInformation($"术语预替换已执行（{(_isCnToForeign ? "中文→外语" : "外语→中文")}）：'{beforeSnippet}' -> '{afterSnippet}'");
                }
                else
                {
                    preprocessed = normalized;
                }

                var terminologyForTranslator = (_useTerminology && !_preprocessTerms) ? termsToUse : null;

                translatedText = await TranslateTextWithRetryAsync(toTranslate, terminologyForTranslator, cellText);
                if (string.IsNullOrEmpty(translatedText))
                    return;
            }

            finalTranslated = translatedText;

            // 记录到 CSV：工作表,位置,原文,预处理后,译文
            lock (csvLock)
            {
                // 基础CSV转义
                string esc(string s)
                {
                    if (s == null) return "";
                    s = s.Replace("\r\n", "\n").Replace("\r", "\n");
                    if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
                        return '"' + s.Replace("\"", "\"\"") + '"';
                    return s;
                }

                csvWriter.WriteLine($"{esc(worksheet.Name)},{esc(location)},{esc(cellText)},{esc(preprocessed)},{esc(finalTranslated)}");
                csvWriter.Flush();
            }

            // 更新单元格（若是合并区域，则对整个区域开启WrapText；仅在左上角写入内容）。若是公式单元格，仅记录CSV，不写回，避免破坏公式
            if (!isFormulaCell)
            {
                UpdateCellText(styleRange, cellText, finalTranslated);
            }
        }

        /// <summary>
        /// 翻译并更新单元格
        /// </summary>
        private async Task TranslateAndUpdateCellAsync(ExcelWorksheet worksheet, int row, int col, string cellText, 
            Dictionary<string, string> terminology, StreamWriter csvWriter, object csvLock)
        {
            try
            {
                // 再次检查是否需要跳过翻译（双重保障）
                if (ShouldSkipCellText(cellText))
                {
                    // 记录被跳过的文本到CSV
                    LogTranslationToCSV(csvWriter, csvLock, worksheet.Name, $"{row},{col}", cellText, "跳过翻译", "", "跳过翻译");
                    _logger.LogDebug($"已记录跳过翻译的文本到CSV: {cellText}");
                    _logger.LogDebug($"跳过翻译单元格 [{row},{col}]：{cellText.Substring(0, Math.Min(50, cellText.Length))}...");
                    return;
                }

                // 使用显示文本进行抽取，避免因类型导致遗漏
                var normalizedText = NormalizeUnitsNearNumbers(cellText);
                string preprocessed = normalizedText;

                // 术语预处理
                Dictionary<string, string> termsToUse = new Dictionary<string, string>(terminology ?? new Dictionary<string, string>());
                if (_useTerminology)
                {
                    var extracted = _termExtractor.ExtractRelevantTerms(cellText, _terminologyLanguageName, _sourceLang, _targetLang);
                    if (extracted != null && extracted.Any())
                    {
                        foreach (var kv in extracted)
                        {
                            if (!termsToUse.ContainsKey(kv.Key)) termsToUse[kv.Key] = kv.Value;
                        }

                        if (_preprocessTerms)
                        {
                            termsToUse = _termExtractor.PreprocessTerms(termsToUse).ToDictionary(k => k.Key, v => v.Value);
                            preprocessed = _termExtractor.ReplaceTermsInText(normalizedText, termsToUse, _isCnToForeign);
                        }
                    }
                }

                var terminologyForTranslator = (_useTerminology && !_preprocessTerms) ? termsToUse : null;
                
                // 开始计时（从发送到模型时开始）
                var translationStartTime = DateTime.Now;
                _logger.LogInformation($"开始翻译单元格 [{row},{col}]，预处理后文本：{preprocessed.Substring(0, Math.Min(50, preprocessed.Length))}...");
                var translatedText = await TranslateTextWithRetryAsync(preprocessed, terminologyForTranslator, cellText);
                var translationDuration = DateTime.Now - translationStartTime;
                _logger.LogInformation($"单元格 [{row},{col}] 翻译完成，耗时: {translationDuration.TotalSeconds:F2}秒");

                if (!string.IsNullOrWhiteSpace(translatedText))
                {
                    _logger.LogDebug($"翻译结果：{translatedText.Substring(0, Math.Min(50, translatedText.Length))}...");
                    // 更新单元格
                    var cell = worksheet.Cells[row, col];
                    UpdateCellText(cell, cellText, translatedText);
                    _logger.LogDebug($"已更新单元格 [{row},{col}] 的内容");

                    // 记录到CSV
                    LogTranslationToCSV(csvWriter, csvLock, worksheet.Name, $"{row},{col}", cellText, preprocessed, translatedText);
                    _logger.LogDebug($"已记录单元格 [{row},{col}] 的翻译结果到CSV");
                    _logger.LogDebug($"翻译结果状态: {(translatedText.Contains("[长度异常]") ? "长度异常" : "翻译成功")}");
                }
                else
                {
                    _logger.LogDebug($"单元格 [{row},{col}] 翻译失败，未获得有效译文");
                    // 记录翻译失败的文本到CSV
                    LogTranslationToCSV(csvWriter, csvLock, worksheet.Name, $"{row},{col}", cellText, preprocessed, "翻译失败", "翻译失败");
                    _logger.LogDebug($"已记录翻译失败的文本到CSV: {cellText}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"翻译单元格 [{row},{col}] 失败: {cellText.Substring(0, Math.Min(50, cellText.Length))}...");
                // 记录异常到CSV
                LogTranslationToCSV(csvWriter, csvLock, worksheet.Name, $"{row},{col}", cellText, "异常失败", "异常失败");
                _logger.LogDebug($"已记录异常的文本到CSV: {cellText}");
            }
        }

        /// <summary>
        /// 记录翻译结果到CSV（辅助方法）
        /// </summary>
        private void LogTranslationToCSV(StreamWriter csvWriter, object csvLock, string worksheetName, string location, 
            string originalText, string preprocessedText, string translatedText, string status = "", string lengthComparison = "")
        {
            try
            {
                lock (csvLock)
                {
                    string esc(string s)
                    {
                        if (s == null) return "";
                        s = s.Replace("\r\n", "\n").Replace("\r", "\n");
                        if (s.Contains('"') || s.Contains(',') || s.Contains('\n'))
                            return '"' + s.Replace("\"", "\"\"") + '"';
                        return s;
                    }

                    // 计算长度对比
                    if (string.IsNullOrEmpty(lengthComparison) && !string.IsNullOrEmpty(originalText) && !string.IsNullOrEmpty(translatedText))
                    {
                        int originalLength = originalText.Length;
                        int translatedLength = translatedText.Length;
                        double ratio = (double)translatedLength / originalLength;
                        lengthComparison = $"{originalLength}→{translatedLength} ({ratio:F2}x)";
                    }

                    // 确定状态
                    string finalStatus = status;
                    if (string.IsNullOrEmpty(finalStatus))
                    {
                        if (string.IsNullOrEmpty(translatedText))
                        {
                            finalStatus = "翻译失败";
                        }
                        else if (preprocessedText == "跳过翻译")
                        {
                            finalStatus = "跳过翻译";
                        }
                        else if (IsTranslationLengthAbnormal(originalText, translatedText))
                        {
                            finalStatus = "长度异常";
                        }
                        else
                        {
                            finalStatus = "翻译成功";
                        }
                    }

                    csvWriter.WriteLine($"{esc(worksheetName)},{esc(location)},{esc(originalText)},{esc(preprocessedText)},{esc(translatedText)},{esc(finalStatus)},{esc(lengthComparison)}");
                    csvWriter.Flush();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "记录翻译结果到CSV失败");
            }
        }

        /// <summary>
        /// 清洗模型返回的译文，移除前后缀、代码块、无用标记与多余空白，返回"纯净译文"
        /// </summary>
        private string CleanTranslationText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            var s = text.Trim();

            // 1) 去除常见前缀
            // 如："译文:", "翻译:", "Translation:", "Translated:", "Result:" 等
            s = Regex.Replace(s, @"^(译文|翻译|Translation|Translated|Result|输出|Output)\s*[:：]-?\s*", string.Empty, RegexOptions.IgnoreCase);

            // 2) 去除Markdown代码块标记
            // ```xxx ... ``` 或 ``` ... ```
            s = Regex.Replace(s, @"^```[a-zA-Z0-9_-]*\s*", string.Empty);
            s = Regex.Replace(s, @"```\s*$", string.Empty);

            // 3) 去除多余的行首尾引号和括号
            s = s.Trim().Trim('"', '\'', '“', '”', '‘', '’', '(', ')', '[', ']', '{', '}');

            // 4) 去除重复的空格和换行
            s = Regex.Replace(s, @"\s+\n", "\n");
            s = Regex.Replace(s, @"\n+\s+", "\n");
            s = Regex.Replace(s, @"\s{2,}", " ");

            // 5) 去除翻译器可能添加的额外注释
            s = Regex.Replace(s, @"\s*\([^)]*\)\s*$", string.Empty);
            s = Regex.Replace(s, @"\s*\[[^\]]*\]\s*$", string.Empty);

            // 6) 保持数字、单位和特殊符号的一致性
            // 确保常见单位符号不被翻译或修改
            var unitPatterns = new[] { "%", "°C", "°F", "kg", "g", "mg", "t", "m", "cm", "mm", "km", "L", "mL", "h", "min", "s", "A", "V", "W", "Ω" };
            foreach (var unit in unitPatterns)
            {
                // 确保单位前后有适当的空格
                string replacement1 = "$1" + unit;
                string replacement2 = unit + " $1";
                s = Regex.Replace(s, $"(\\d+)\\s*{Regex.Escape(unit)}", replacement1);
                s = Regex.Replace(s, $"{Regex.Escape(unit)}\\s*(\\d+)", replacement2);
            }

            return s.Trim();
        }

        /// <summary>
        /// 更新单元格文本
        /// </summary>
        private void UpdateCellText(ExcelRange cell, string originalText, string translatedText)
        {
            // 纯净化译文，去除模型可能添加的多余标记
            var clean = CleanTranslationText(translatedText);

            // 获取顶端单元格（若为合并区域则取左上角单元格），用于写入内容
            var ws = cell.Worksheet;
            var topLeftCell = ws.Cells[cell.Start.Row, cell.Start.Column];

            // 检查译文长度是否异常
            bool isLengthAbnormal = IsTranslationLengthAbnormal(originalText, clean);

            if (_outputFormat == "bilingual")
            {
                // 双语对照模式：使用富文本分行显示，增强可读性
                // 原文保留原样；译文使用斜体和颜色区分
                try
                {
                    topLeftCell.RichText.Clear();
                    var rt1 = topLeftCell.RichText.Add(originalText ?? string.Empty);
                    // 颜色与字体保持默认，减少对原始视觉的干扰
                    topLeftCell.RichText.Add(Environment.NewLine);
                    var rt2 = topLeftCell.RichText.Add(clean);
                    rt2.Italic = true;
                    // 轻微区分色（深蓝），避免打印可读性问题
                    rt2.Color = isLengthAbnormal ? System.Drawing.Color.Red : System.Drawing.Color.FromArgb(0, 102, 204);
                }
                catch
                {
                    // 若富文本写入失败，则退回到简单换行文本
                    topLeftCell.Value = $"{originalText}{Environment.NewLine}{clean}";
                    // 若译文长度异常，设置单元格文本颜色为红色
                    if (isLengthAbnormal)
                    {
                        topLeftCell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                    }
                }
            }
            else
            {
                // 仅翻译结果模式：只写入纯净译文
                topLeftCell.Value = clean;
                // 若译文长度异常，设置单元格文本颜色为红色
                if (isLengthAbnormal)
                {
                    topLeftCell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                }
            }

            // 开启自动换行，确保"译文换行"在单元格内显示为新行
            var style = cell.Style; // 对整个区域（含合并区域）生效
            style.WrapText = true;
            style.ShrinkToFit = false;
            style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        }

        /// <summary>
        /// 更新进度
        /// </summary>
        private void UpdateProgress(double progress, string message)
        {
            var now = DateTime.Now;
            if ((now - _lastProgressUpdate).TotalMilliseconds >= _minProgressUpdateInterval)
            {
                _progressCallback?.Invoke(progress, message);
                _lastProgressUpdate = now;
            }
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
        /// 简单启发式：跳过无需翻译的短文本/符号/代码样式内容
        /// </summary>
        private bool ShouldSkipCellText(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return true;
            var t = s.Trim();

            // 仅符号/装饰性字符：跳过
            string[] skipTokens = { "N/A", "NA", "--", "-", "/", "\\", "#", "*", "√", "×", "—" };
            if (Array.Exists(skipTokens, k => string.Equals(k, t, StringComparison.OrdinalIgnoreCase)))
            {
                // 记录跳过的文本
                LogSkippedTranslation(s);
                return true;
            }

            // 常见单位和缩写：跳过翻译
            string[] commonAbbreviations = { "min", "max", "avg", "sum", "count", "total", "unit", "pcs", "kg", "g", "mg", "t", "m", "cm", "mm", "km", "L", "mL", "h", "s", "ms", "A", "V", "W", "Ω", "%" };
            if (Array.Exists(commonAbbreviations, k => string.Equals(k, t, StringComparison.OrdinalIgnoreCase)))
            {
                // 记录跳过的文本
                LogSkippedTranslation(s);
                return true;
            }

            // 纯数字：跳过
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"^[0-9]+(\.[0-9]+)?%?$"))
            {
                // 记录跳过的文本
                LogSkippedTranslation(s);
                return true;
            }

            // 核心逻辑：如果文本不包含原语种字符，则跳过
            // 这意味着纯英文（在中译外时）或纯中文（在外译中时）会被跳过
            // 这符合用户"若为纯英文文本则跳过翻译，若不是则需要翻译"的要求
            if (!ContainsSourceLanguage(t))
            {
                // 记录跳过的文本
                LogSkippedTranslation(s);
                return true;
            }

            // 单字符：若为中文或单位名，不跳过，避免漏译"台""件""米"等
            if (t.Length == 1)
            {
                if (ZhUnitMap.ContainsKey(t)) return false; // 单位
                if (System.Text.RegularExpressions.Regex.IsMatch(t, @"^[\u4e00-\u9fff]$") || System.Text.RegularExpressions.Regex.IsMatch(t, @"^[A-Za-z]$") || System.Text.RegularExpressions.Regex.IsMatch(t, @"^[0-9]$") || System.Text.RegularExpressions.Regex.IsMatch(t, @"^[%°]$"))
                    return false;
                // 其它单字符多为符号，记录并跳过
                LogSkippedTranslation(s);
                return true;
            }

            return false;
        }

        /// <summary>
        /// 检查文本是否需要跳过翻译
        /// </summary>
        private bool ShouldSkipTextForTranslation(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return true;
            var t = text.Trim();

            // 长度检查：过短的文本可能无法翻译
            if (t.Length <= 2)
            {
                // 检查是否为常见缩写
                string[] commonAbbreviations = { "min", "max", "avg", "sum", "count", "total", "unit", "pcs", "kg", "g", "mg", "t", "m", "cm", "mm", "km", "L", "mL", "h", "s", "ms", "A", "V", "W", "Ω", "%" };
                if (Array.Exists(commonAbbreviations, k => string.Equals(k, t, StringComparison.OrdinalIgnoreCase)))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 记录跳过的翻译
        /// </summary>
        private void LogSkippedTranslation(string text)
        {
            try
            {
                var currentTranslator = _translationService.CurrentTranslator;
                if (currentTranslator != null)
                {
                    currentTranslator.LogFailureTranslation(text, "跳过翻译", _sourceLang, _targetLang);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "记录跳过翻译失败");
            }
        }

        /// <summary>
        /// 检查译文长度是否异常
        /// </summary>
        private bool IsTranslationLengthAbnormal(string originalText, string translatedText)
        {
            if (string.IsNullOrWhiteSpace(originalText) || string.IsNullOrWhiteSpace(translatedText))
                return false;

            int originalLength = originalText.Length;
            int translatedLength = translatedText.Length;

            // 根据原文长度调整阈值
            if (originalLength <= 10)
            {
                // 对于短文本，使用更宽松的阈值
                return translatedLength < originalLength * 0.2 || translatedLength > originalLength * 5;
            }
            else
            {
                // 对于长文本，使用更严格的阈值，当英文长度是中文的4倍时标红
                return translatedLength < originalLength * 0.3 || translatedLength > originalLength * 4;
            }
        }



        // 简单单位词典（可扩展）
        private static readonly Dictionary<string, string> ZhUnitMap = new(StringComparer.Ordinal)
        {
            ["台"] = "units",
            ["只"] = "pcs",
            ["件"] = "pcs",
            ["个"] = "pcs",
            ["套"] = "sets",
            ["桶"] = "barrels",
            ["吨"] = "tons",
            ["千克"] = "kg",
            ["公斤"] = "kg",
            ["克"] = "g",
            ["米"] = "m",
            ["厘米"] = "cm",
            ["毫米"] = "mm",
            ["升"] = "L",
            ["毫升"] = "mL",
            ["小时"] = "h",
            ["分钟"] = "min",
            ["秒"] = "s"
        };

        // 若单元格为"纯单位词"或"纯数字+单位"，直接进行单位翻译
        private static bool TryTranslateStandaloneUnit(string text, out string translated)
        {
            translated = null;
            if (string.IsNullOrWhiteSpace(text)) return false;
            var t = text.Trim();
            // 纯数字直接返回 false（不翻译）
            if (System.Text.RegularExpressions.Regex.IsMatch(t, @"^[0-9]+(\.[0-9]+)?$")) return false;

            // 纯单位词
            if (ZhUnitMap.TryGetValue(t, out var unit))
            {
                translated = unit;
                return true;
            }

            // 数字 + 单位（可包含空格）
            var m = System.Text.RegularExpressions.Regex.Match(t, @"^([0-9]+(?:\.[0-9]+)?)\s*([一二三四五六七八九十百千万亿]?)([台只件个套桶吨千克公斤克米厘米毫米升毫升小时分钟秒])$");
            if (m.Success)
            {
                var number = m.Groups[1].Value;
                var zhUnit = m.Groups[3].Value;
                if (ZhUnitMap.TryGetValue(zhUnit, out var enUnit))
                {
                    translated = $"{number} {enUnit}";
                    return true;
                }
            }
            return false;
        }

        // 规范化：将"数字+中文单位"附近加空格，便于LLM理解，不改变纯数字
        private static string NormalizeUnitsNearNumbers(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            var s = text;
            // 例："10台设备" → "10 台设备"；"5吨" → "5 吨"
            s = System.Text.RegularExpressions.Regex.Replace(s, @"([0-9]+(?:\.[0-9]+)?)([台只件个套桶吨千克公斤克米厘米毫米升毫升小时分钟秒])", "$1 $2");
            return s;
        }

        // 移除所有空格
        private static string RemoveSpaces(string s) => string.IsNullOrEmpty(s) ? s : System.Text.RegularExpressions.Regex.Replace(s, @"\s+", string.Empty);

        /// <summary>
        /// 带重试机制的翻译
        /// </summary>
        private async Task<string> TranslateTextWithRetryAsync(string text, Dictionary<string, string> terminology, string originalText = null)
        {
            // 用户要求：无论文本是否被翻译，都要再走一遍翻译流程
            // 因此移除此处的跳过检查，依赖 ShouldSkipCellText 在提取阶段的过滤
            /*
            if (ShouldSkipTextForTranslation(text))
            {
                _logger.LogDebug($"跳过翻译文本：{text}");
                // ...
                return string.Empty;
            }
            */

            for (int attempt = 0; attempt < _retryCount; attempt++)
            {
                try
                {
                    // 添加超时处理
                    var cts = new CancellationTokenSource(TimeSpan.FromSeconds(60)); // 60秒超时
                    // 传递 null 作为 prompt，让 BaseTranslator 或具体的 Translator 决定是否需要默认 prompt (通常不需要)
                    var result = await _translationService.TranslateTextAsync(text, terminology, _sourceLang, _targetLang, null, originalText);

                    // 译文净化与质量校验
                    var clean = CleanTranslationText(result);
                    
                    /* 用户要求不再进行质量检测
                    if (string.IsNullOrWhiteSpace(clean))
                        throw new Exception("译文为空");

                    // 额外检查：如果译文与预处理后的文本相同，但与原始文本不同，不视为无效
                    if (!string.Equals(RemoveSpaces(originalText ?? text), RemoveSpaces(text), StringComparison.Ordinal) && 
                        string.Equals(RemoveSpaces(text), RemoveSpaces(clean), StringComparison.Ordinal))
                    {
                        _logger.LogInformation($"译文与预处理后的文本相同，但与原始文本不同，视为有效翻译");
                    }
                    // 若译文与原始文本在去空白后一致，则视为无效
                    else if (string.Equals(RemoveSpaces(originalText ?? text), RemoveSpaces(clean), StringComparison.Ordinal))
                        throw new Exception("译文与原文近似一致");
                    */

                    // 额外的质量检查：确保译文长度合理
                    bool isLengthAbnormal = IsTranslationLengthAbnormal(text, clean);
                    if (isLengthAbnormal)
                    {
                        _logger.LogWarning($"译文长度异常：原文长度 {text.Length}，译文长度 {clean.Length}");
                        // 记录长度异常的译文
                        var currentTranslator = _translationService.CurrentTranslator;
                        if (currentTranslator != null)
                        {
                            try
                            {
                                currentTranslator.LogAbnormalTranslation(originalText ?? text, clean, _sourceLang, _targetLang);
                            }
                            catch (Exception ex)
                            {
                                _logger.LogWarning(ex, "记录译文长度异常失败");
                            }
                        }
                        // 不再在译文中添加异常标记，改为在更新单元格时标红
                    }

                    return clean;
                }
                catch (TaskCanceledException ex)
                {
                    _logger.LogDebug(ex, $"翻译超时，尝试 {attempt + 1}/{_retryCount}，文本：{text.Substring(0, Math.Min(50, text.Length))}...");

                    if (attempt == _retryCount - 1)
                    {
                        _logger.LogDebug($"翻译重试次数已用完，跳过该文本：{text.Substring(0, Math.Min(100, text.Length))}...");
                        // 记录超时到TimeOut文件夹
                        var currentTranslator = _translationService.CurrentTranslator;
                        if (currentTranslator != null)
                        {
                            try
                            {
                                currentTranslator.LogTimeoutTranslation(originalText ?? text, 60, _sourceLang, _targetLang);
                            }
                            catch (Exception logEx)
                            {
                                _logger.LogWarning(logEx, "记录翻译超时失败");
                            }
                        }
                        return string.Empty;
                    }

                    // 递增延迟，避免频繁请求导致的失败
                    int currentDelay = _retryDelay * (attempt + 1);
                    _logger.LogDebug($"等待 {currentDelay}ms 后重试");
                    await Task.Delay(currentDelay);
                }
                catch (Exception ex)
                {
                    // 检查是否是质量不佳错误
                    bool isQualityError = ex.Message.Contains("译文质量不佳");
                    
                    if (isQualityError)
                    {
                        // 质量不佳错误，不再重试，直接返回空字符串
                        _logger.LogDebug($"译文质量不佳，不再重试，文本：{text.Substring(0, Math.Min(50, text.Length))}...");
                        // 记录失败到Failure文件夹
                        var currentTranslator = _translationService.CurrentTranslator;
                        if (currentTranslator != null)
                        {
                            try
                            {
                                currentTranslator.LogFailureTranslation(originalText ?? text, ex.Message, _sourceLang, _targetLang);
                            }
                            catch (Exception logEx)
                            {
                                _logger.LogWarning(logEx, "记录翻译失败失败");
                            }
                        }
                        return string.Empty;
                    }
                    
                    _logger.LogDebug(ex, $"翻译失败，尝试 {attempt + 1}/{_retryCount}，文本：{text.Substring(0, Math.Min(50, text.Length))}...");

                    if (attempt == _retryCount - 1)
                    {
                        _logger.LogDebug($"翻译重试次数已用完，跳过该文本：{text.Substring(0, Math.Min(100, text.Length))}...");
                        // 记录失败到Failure文件夹
                        var currentTranslator = _translationService.CurrentTranslator;
                        if (currentTranslator != null)
                        {
                            try
                            {
                                currentTranslator.LogFailureTranslation(originalText ?? text, ex.Message, _sourceLang, _targetLang);
                            }
                            catch (Exception logEx)
                            {
                                _logger.LogWarning(logEx, "记录翻译失败失败");
                            }
                        }
                        return string.Empty;
                    }

                    // 递增延迟，避免频繁请求导致的失败
                    int currentDelay = _retryDelay * (attempt + 1);
                    _logger.LogDebug($"等待 {currentDelay}ms 后重试");
                    await Task.Delay(currentDelay);
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// 清理语言，保留指定的语种
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="keepLanguage">要保留的语言代码，zh为中文，en为英文</param>
        /// <returns>处理后的文件路径</returns>
        public async Task<string> CleanLanguageAsync(string filePath, string keepLanguage)
        {
            _logger.LogInformation($"开始清理语言，文件: {filePath}, 保留语言: {keepLanguage}");
            
            try
            {
                // 生成输出文件路径
                string outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输出");
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                string extension = Path.GetExtension(filePath);
                string outputFilePath = Path.Combine(outputDir, $"{fileName}_清理语言_{DateTime.Now:yyyyMMddHHmmss}{extension}");

                // 复制原始文件到输出路径
                File.Copy(filePath, outputFilePath, true);
                _logger.LogInformation($"已创建输出文件: {outputFilePath}");

                // 打开Excel文件
                using var package = new ExcelPackage(new FileInfo(outputFilePath));
                int totalWorksheets = package.Workbook.Worksheets.Count;
                int processedWorksheets = 0;

                // 遍历所有工作表
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    processedWorksheets++;
                    double worksheetProgress = (double)processedWorksheets / totalWorksheets * 100;
                    _progressCallback?.Invoke(worksheetProgress, $"处理工作表 {worksheet.Name} ({processedWorksheets}/{totalWorksheets})");

                    _logger.LogInformation($"处理工作表: {worksheet.Name}");

                    // 获取工作表的使用范围
                    var usedRange = worksheet.Cells[worksheet.Dimension?.Address ?? "A1:A1"];
                    int startRow = usedRange.Start.Row;
                    int endRow = usedRange.End.Row;
                    int startCol = usedRange.Start.Column;
                    int endCol = usedRange.End.Column;

                    // 遍历所有单元格
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            await ProcessCellForCleanupAsync(cell, keepLanguage);
                        }
                    }
                }

                // 保存处理后的文件
                await package.SaveAsAsync(new FileInfo(outputFilePath));
                _logger.LogInformation($"清理完成，输出文件: {outputFilePath}");
                
                _progressCallback?.Invoke(100, "清理完成");
                return outputFilePath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "清理语言失败");
                throw;
            }
        }

        /// <summary>
        /// 处理单元格进行语言清理
        /// </summary>
        /// <param name="cell">要处理的单元格</param>
        /// <param name="keepLanguage">要保留的语言代码</param>
        private async Task ProcessCellForCleanupAsync(ExcelRange cell, string keepLanguage)
        {
            try
            {
                // 检查是否需要跳过
                if (ShouldSkipCellForCleanup(cell))
                {
                    _logger.LogDebug($"跳过单元格 [{cell.Start.Row},{cell.Start.Column}]：{cell.Text}");
                    return;
                }

                // 获取单元格文本
                var cellText = cell.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(cellText))
                    return;

                // 使用智能清理（模型提取）
                var cleanedText = await _translationService.SmartCleanTextAsync(cellText, keepLanguage);

                // 如果结果有变化且不为空，则更新单元格
                if (cleanedText != cellText && !string.IsNullOrEmpty(cleanedText))
                {
                    // 直接设置Value会清除RichText格式，符合清理要求
                    cell.Value = cleanedText;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"处理单元格 [{cell.Start.Row},{cell.Start.Column}] 失败");
            }
        }

        /// <summary>
        /// 检查是否需要跳过单元格
        /// </summary>
        /// <param name="cell">要检查的单元格</param>
        /// <returns>是否需要跳过</returns>
        private bool ShouldSkipCellForCleanup(ExcelRange cell)
        {
            // 跳过合并单元格的非左上角单元格
            if (cell.Merge)
            {
                var ws = cell.Worksheet;
                if (ws.MergedCells != null && ws.MergedCells.Count > 0)
                {
                    foreach (var mergedAddress in ws.MergedCells)
                    {
                        try
                        {
                            var rng = new ExcelAddress(mergedAddress);
                            if (cell.Start.Row >= rng.Start.Row && cell.Start.Row <= rng.End.Row && 
                                cell.Start.Column >= rng.Start.Column && cell.Start.Column <= rng.End.Column)
                            {
                                // 不是左上角的单元格就跳过
                                if (!(cell.Start.Row == rng.Start.Row && cell.Start.Column == rng.Start.Column))
                                {
                                    return true;
                                }
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogWarning(ex, $"解析合并单元格地址失败: {mergedAddress}");
                        }
                    }
                }
            }

            // 跳过公式单元格
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                return true;
            }

            // 跳过纯数字单元格
            var cellText = cell.Text ?? string.Empty;
            if (System.Text.RegularExpressions.Regex.IsMatch(cellText.Trim(), @"^[0-9]+(\.[0-9]+)?%?$"))
            {
                return true;
            }

            // 跳过纯编码单元格
            if (System.Text.RegularExpressions.Regex.IsMatch(cellText.Trim(), @"^[A-Za-z0-9_]+$"))
            {
                return true;
            }

            return false;
        }


        private class CellProcessingContext
        {
            public int Row { get; set; }
            public int Col { get; set; }
            public string OriginalText { get; set; }
            public string PreprocessedText { get; set; }
            public string TranslatedText { get; set; }
            public bool IsFormula { get; set; }
            public bool ShouldSkip { get; set; }
            public string SkipReason { get; set; }
            public Dictionary<string, string> Terminology { get; set; }
            public ExcelAddress MergedRange { get; set; }
            public bool IsMerged { get; set; }
            public string TranslationStatus { get; set; }
            public string LengthComparison { get; set; }
        }
    }
}