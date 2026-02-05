using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Threading.Tasks;
using System.Threading;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Tesseract;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// PDF文档处理器，支持无损翻译和排版保留
    /// 参考PDFMathTranslate项目的实现方式
    /// </summary>
    public class PDFProcessor
    {
        private readonly TranslationService _translationService;
        private readonly ILogger<PDFProcessor> _logger;
        private readonly TermExtractor _termExtractor;
        private bool _useTerminology = true;
        private bool _preprocessTerms = true;
        private bool _exportPdf = false;
        private string _sourceLang = "zh";
        private string _targetLang = "en";
        private bool _isCnToForeign = true;
        private string _outputFormat = "bilingual";
        private Action<double, string> _progressCallback;
        private int _retryCount = 1;
        private int _retryDelay = 1000;
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private const double _minProgressUpdateInterval = 100;

        private bool _useOcr = true;
        private bool _useAdobeAcrobatPreprocess = false;
        private string _tesseractDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
        private TesseractEngine _ocrEngine;
        private TranslationOutputConfig _outputConfig = new TranslationOutputConfig();

        private readonly List<Regex> _latexPatterns = new List<Regex>
        {
            new Regex(@"\$\$(.*?)\$\$", RegexOptions.Singleline),
            new Regex(@"\$(.*?)\$", RegexOptions.Singleline),
            new Regex(@"\\begin\{equation\*?\}(.*?)\\end\{equation\*?\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{align\*?\}(.*?)\\end\{align\*?\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{eqnarray\*?\}(.*?)\\end\{eqnarray\*?\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{gather\*?\}(.*?)\\end\{gather\*?\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{multline\*?\}(.*?)\\end\{multline\*?\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{math\}(.*?)\\end\{math\}", RegexOptions.Singleline),
            new Regex(@"\\begin\{displaymath\}(.*?)\\end\{displaymath\}", RegexOptions.Singleline),
            new Regex(@"\\[(.*?)\\]", RegexOptions.Singleline),
            new Regex(@"\\((.*?)\\)", RegexOptions.Singleline),
            new Regex(@"\\begin\{cases\}(.*?)\\end\{cases\}", RegexOptions.Singleline)
        };

        public PDFProcessor(TranslationService translationService, ILogger<PDFProcessor> logger, bool ocrAvailable = true)
        {
            _translationService = translationService ?? throw new ArgumentNullException(nameof(translationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _termExtractor = new TermExtractor();

            _useOcr = ocrAvailable;

            if (_useOcr)
            {
                InitializeOcrEngine();
            }
            else
            {
                _logger.LogInformation("OCR功能不可用，跳过OCR引擎初始化");
            }
        }

        private void InitializeOcrEngine()
        {
            try
            {
                if (!Directory.Exists(_tesseractDataPath))
                {
                    Directory.CreateDirectory(_tesseractDataPath);
                    _logger.LogWarning($"创建tessdata目录: {_tesseractDataPath}");
                }

                var chiSimPath = Path.Combine(_tesseractDataPath, "chi_sim.traineddata");
                if (!File.Exists(chiSimPath))
                {
                    _logger.LogWarning($"未找到中文OCR数据文件: {chiSimPath}，OCR功能可能无法正常工作");
                    _logger.LogWarning("请从 https://github.com/tesseract-ocr/tessdata 下载 chi_sim.traineddata 并放置到 tessdata 目录");
                }

                _ocrEngine = new TesseractEngine(_tesseractDataPath, "chi_sim", EngineMode.Default);
                _logger.LogInformation("Tesseract OCR引擎初始化成功");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Tesseract OCR引擎初始化失败，将禁用OCR功能");
                _useOcr = false;
            }
        }

        public void SetProgressCallback(Action<double, string> callback)
        {
            _progressCallback = callback;
        }

        public void SetTranslationOptions(bool useTerminology = true, bool preprocessTerms = true,
            bool exportPdf = false, string sourceLang = "zh", string targetLang = "en", string outputFormat = "bilingual",
            bool useOcr = true, bool useAdobeAcrobatPreprocess = false)
        {
            _useTerminology = useTerminology;
            _preprocessTerms = preprocessTerms;
            _exportPdf = exportPdf;
            _sourceLang = sourceLang;
            _targetLang = targetLang;
            _outputFormat = outputFormat;
            _isCnToForeign = sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh");
            _useOcr = useOcr;
            _useAdobeAcrobatPreprocess = useAdobeAcrobatPreprocess;

            _logger.LogInformation($"翻译方向: {(_isCnToForeign ? "中文→外语" : "外语→中文")}");
            _logger.LogInformation($"输出格式: {GetOutputFormatName(_outputConfig.Format)}");
            _logger.LogInformation($"分段模式: {(_outputConfig.Segmentation == TranslationOutputConfig.SegmentationMode.Paragraph ? "按段落" : "按句子")}");
            _logger.LogInformation($"术语高亮: {(_outputConfig.HighlightTerms ? "启用" : "禁用")}");
            _logger.LogInformation($"OCR功能: {(_useOcr ? "启用" : "禁用")}");
            _logger.LogInformation($"Adobe Acrobat预处理: {(_useAdobeAcrobatPreprocess ? "启用" : "禁用")}");
        }

        private string GetOutputFormatName(TranslationOutputConfig.OutputFormat format)
        {
            return format switch
            {
                TranslationOutputConfig.OutputFormat.BilingualVertical => "双语对照（垂直）",
                TranslationOutputConfig.OutputFormat.BilingualSideBySide => "并排对照",
                TranslationOutputConfig.OutputFormat.Inline => "行内对照",
                TranslationOutputConfig.OutputFormat.OriginalOnly => "仅原文",
                TranslationOutputConfig.OutputFormat.TranslatedOnly => "仅译文",
                _ => "双语对照（垂直）"
            };
        }

        public void SetOcrOptions(bool useOcr, bool useAdobeAcrobatPreprocess = false, string tesseractDataPath = null)
        {
            _useOcr = useOcr;
            _useAdobeAcrobatPreprocess = useAdobeAcrobatPreprocess;
            
            if (!string.IsNullOrEmpty(tesseractDataPath))
            {
                _tesseractDataPath = tesseractDataPath;
            }

            _logger.LogInformation($"OCR选项更新: useOcr={_useOcr}, useAdobeAcrobatPreprocess={_useAdobeAcrobatPreprocess}");
        }

        public void SetOutputConfig(TranslationOutputConfig config)
        {
            _outputConfig = config ?? new TranslationOutputConfig();
            _logger.LogInformation($"输出配置更新: Format={_outputConfig.Format}, Segmentation={_outputConfig.Segmentation}, HighlightTerms={_outputConfig.HighlightTerms}");
        }

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguage,
            Dictionary<string, string> terminology)
        {
            UpdateProgress(0.01, "开始处理PDF文档...");

            if (!File.Exists(filePath))
            {
                _logger.LogError($"文件不存在: {filePath}");
                throw new FileNotFoundException("文件不存在");
            }

            var outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输出");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var outputFileName = $"{fileName}_带翻译_{timeStamp}.docx";
            var outputPath = Path.Combine(outputDir, outputFileName);

            _logger.LogInformation($"PDF处理器输出配置:");
            _logger.LogInformation($"  原始文件: {filePath}");
            _logger.LogInformation($"  输出路径: {outputPath}");

            try
            {
                string processedFilePath = filePath;

                if (_useAdobeAcrobatPreprocess)
                {
                    UpdateProgress(0.02, "使用Adobe Acrobat预处理PDF...");
                    processedFilePath = await PreprocessWithAdobeAcrobatAsync(filePath);
                    if (processedFilePath != filePath)
                    {
                        _logger.LogInformation($"Adobe Acrobat预处理完成，临时文件: {processedFilePath}");
                    }
                }

                await ProcessPDFDocumentAsync(processedFilePath, outputPath, terminology);

                if (processedFilePath != filePath && File.Exists(processedFilePath))
                {
                    try
                    {
                        File.Delete(processedFilePath);
                        _logger.LogInformation($"已删除临时预处理文件: {processedFilePath}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"删除临时文件失败: {processedFilePath}");
                    }
                }

                if (_exportPdf)
                {
                    UpdateProgress(0.95, "正在导出为PDF...");
                    var pdfPath = await ConvertWordToPdfAsync(outputPath);
                    if (!string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath))
                    {
                        outputPath = pdfPath;
                        _logger.LogInformation($"已导出PDF: {pdfPath}");
                    }
                }

                UpdateProgress(1.0, "翻译完成");

                _logger.LogInformation($"PDF文档翻译完成，输出路径: {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "PDF文档处理失败");
                throw;
            }
        }

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private async Task ProcessPDFDocumentAsync(string inputPath, string outputPath,
            Dictionary<string, string> terminology)
        {
            UpdateProgress(0.05, "打开PDF文件...");

            using var pdfReader = new PdfReader(inputPath);
            using var pdfDoc = new PdfDocument(pdfReader);
            var totalPages = pdfDoc.GetNumberOfPages();

            _logger.LogInformation($"PDF文档共 {totalPages} 页");

            UpdateProgress(0.1, "初始化输出...");

            var outputDir = Path.GetDirectoryName(outputPath)!;
            var logDir = Path.Combine(outputDir, "日志");
            var imagesDir = Path.Combine(logDir, "图片");
            if (!Directory.Exists(logDir)) Directory.CreateDirectory(logDir);
            if (!Directory.Exists(imagesDir)) Directory.CreateDirectory(imagesDir);

            var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var csvPath = Path.Combine(logDir, $"PDF翻译日志_{timeStamp}.csv");
            var jsonlPath = Path.Combine(imagesDir, $"{Path.GetFileNameWithoutExtension(inputPath)}_layout.jsonl");

            using var csvWriter = new StreamWriter(csvPath, false, Encoding.UTF8);
            await csvWriter.WriteLineAsync("页码,LayoutID,段落索引,分段索引,原文,译文,位置X,位置Y,宽,高,字体,字号");

            var pageLayouts = new Dictionary<int, List<LayoutItem>>();

            UpdateProgress(0.15, "提取内容与分析布局...");

            for (int i = 1; i <= totalPages; i++)
            {
                var strategy = new CompositeLayoutStrategy(i, imagesDir);
                var parser = new PdfCanvasProcessor(strategy);
                parser.ProcessPageContent(pdfDoc.GetPage(i));
                
                var rawItems = strategy.GetItems();
                
                // Check for garbage text (encoding issues)
                bool needsOcr = !rawItems.Any(item => item is TextLayoutItem);
                if (!needsOcr && _useOcr)
                {
                    var allText = string.Join("", rawItems.OfType<TextLayoutItem>().Select(t => t.Content));
                    if (IsGarbageText(allText))
                    {
                        needsOcr = true;
                        rawItems.RemoveAll(item => item is TextLayoutItem); // Clear garbage text
                        _logger.LogWarning($"第 {i} 页文本提取结果疑似乱码（包含大量PUA字符或未知符号），将强制使用OCR");
                    }
                }

                // OCR Fallback if needed
                if (needsOcr && _useOcr)
                {
                    UpdateProgress(0.15 + (0.15 * i / totalPages), $"正在OCR第 {i}/{totalPages} 页...");
                    var ocrItems = PerformOcrOnImages(rawItems.OfType<ImageLayoutItem>().ToList(), i, imagesDir);
                    rawItems.AddRange(ocrItems);
                }

                var analyzedItems = AnalyzePageLayout(rawItems);
                pageLayouts[i] = MergeTextItems(analyzedItems);

                UpdateProgress(0.15 + (0.15 * i / totalPages), $"正在分析第 {i}/{totalPages} 页布局...");
            }

            // Save Layout to JSONL
            using (var jsonlWriter = new StreamWriter(jsonlPath, false, Encoding.UTF8))
            {
                foreach (var kvp in pageLayouts)
                {
                    foreach (var item in kvp.Value)
                    {
                        var json = JsonSerializer.Serialize(item);
                        await jsonlWriter.WriteLineAsync(json);
                    }
                }
            }

            UpdateProgress(0.3, "准备翻译文本...");

            var textSegmentsToTranslate = new List<(int Page, string LayoutId, int SegIndex, string Text)>();

            foreach (var kvp in pageLayouts)
            {
                var pageNum = kvp.Key;
                foreach (var item in kvp.Value)
                {
                    if (item is not TextLayoutItem textItem)
                        continue;

                    textItem.Segments.Clear();

                    var pieces = SplitParagraphToSentences(textItem.Content);
                    if (pieces.Count == 0)
                        continue;

                    for (int s = 0; s < pieces.Count; s++)
                    {
                        var sentence = pieces[s].Text;
                        if (string.IsNullOrWhiteSpace(sentence))
                            continue;

                        textItem.Segments.Add(new TextSegment
                        {
                            Original = sentence,
                            Separator = pieces[s].Separator,
                            LogIndex = s
                        });
                        textSegmentsToTranslate.Add((pageNum, textItem.Id, textItem.Segments.Count - 1, sentence));
                    }

                    if (textItem.Segments.Count > 0)
                    {
                        textItem.Content = string.Concat(textItem.Segments.Select(seg => (seg.Original ?? string.Empty) + (seg.Separator ?? string.Empty)));
                    }
                }
            }

            if (_useTerminology && _preprocessTerms && terminology != null && terminology.Count > 0)
            {
                 UpdateProgress(0.35, "术语预处理...");
                 var allTexts = textSegmentsToTranslate.Select(t => t.Text).ToList();
                 var preReplaceCsvPath = Path.Combine(logDir, $"PDF术语预替换_{timeStamp}.csv");
                 var preprocessed = _termExtractor.PreprocessTextsWithMirroredTerms(allTexts, terminology, preReplaceCsvPath);
                 
                 for(int k=0; k<textSegmentsToTranslate.Count; k++)
                 {
                     var t = textSegmentsToTranslate[k];
                     var item = pageLayouts[t.Page].FirstOrDefault(i => i.Id == t.LayoutId) as TextLayoutItem;
                     if(item != null && k < preprocessed.Count)
                     {
                         item.Segments[t.SegIndex].Preprocessed = preprocessed[k];
                     }
                 }
            }

            UpdateProgress(0.4, "开始并发翻译...");

            var maxConcurrency = Math.Max(1, _translationService.GetEffectiveMaxParallelism());
            var finished = 0;
            var totalSegments = textSegmentsToTranslate.Count;
            var csvLock = new object();

            var textItemMapByPage = pageLayouts.ToDictionary(
                kvp => kvp.Key,
                kvp => kvp.Value.OfType<TextLayoutItem>().ToDictionary(t => t.Id, t => t));

            await Parallel.ForEachAsync(
                textSegmentsToTranslate,
                new ParallelOptions { MaxDegreeOfParallelism = maxConcurrency },
                async (segment, ct) =>
                {
                    if (!textItemMapByPage.TryGetValue(segment.Page, out var map)) return;
                    if (!map.TryGetValue(segment.LayoutId, out var item)) return;
                    if (segment.SegIndex < 0 || segment.SegIndex >= item.Segments.Count) return;

                    var segmentModel = item.Segments[segment.SegIndex];
                    var textToTranslate = segmentModel.Preprocessed ?? segmentModel.Original ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(textToTranslate))
                        return;

                    string translated;
                    try
                    {
                        var (protectedText, placeholders) = ProtectMathFormulas(textToTranslate);
                        var termsForTranslator = (_useTerminology && _preprocessTerms) ? null : terminology;
                        translated = await TranslateTextWithRetryAsync(protectedText, termsForTranslator, textToTranslate);
                        translated = RestoreMathFormulas(translated, placeholders);
                    }
                    catch
                    {
                        translated = textToTranslate;
                    }

                    segmentModel.Translated = translated;

                    lock (csvLock)
                    {
                        csvWriter.WriteLine($"{segment.Page},{segment.LayoutId},{segment.SegIndex},0,{EscapeCsv(segment.Text)},{EscapeCsv(translated)},{item.X},{item.Y},{item.Width},{item.Height},{EscapeCsv(item.FontName)},{item.FontSize}");
                    }

                    var done = Interlocked.Increment(ref finished);
                    if (done % 50 == 0 || done == totalSegments)
                    {
                        UpdateProgress(0.4 + (0.5 * done / Math.Max(1, totalSegments)), $"翻译进度 {done}/{totalSegments}...");
                    }
                });

            await csvWriter.FlushAsync();

            UpdateProgress(0.9, "生成输出文档...");

            await GenerateWordOutputAsync(outputPath, pageLayouts);

            if (_exportPdf)
            {
                UpdateProgress(0.95, "生成PDF文档...");
                await GeneratePdfOutputAsync(outputPath.Replace(".docx", ".pdf"), pageLayouts, inputPath);
            }
            
             UpdateProgress(1.0, "完成");
             _logger.LogInformation($"PDF文档处理完成: {outputPath}");
        }

        private bool IsGarbageText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            int garbageCount = 0;
            int totalCount = 0;

            foreach (char c in text)
            {
                if (char.IsWhiteSpace(c)) continue;
                totalCount++;

                // PUA (Private Use Area) E000-F8FF often indicates missing font mapping
                if (c >= 0xE000 && c <= 0xF8FF)
                {
                    garbageCount++;
                }
                // Replacement Character
                else if (c == '\uFFFD')
                {
                    garbageCount++;
                }
                // Unexpected control characters (excluding standard whitespace)
                else if (char.IsControl(c) && c != '\n' && c != '\r' && c != '\t')
                {
                    garbageCount++;
                }
            }

            if (totalCount == 0) return false;

            // If > 20% characters are suspicious, treat as garbage
            return (double)garbageCount / totalCount > 0.2;
        }

        private async Task GenerateWordOutputAsync(string outputPath, Dictionary<int, List<LayoutItem>> pageLayouts)
        {
            await Task.Run(() =>
            {
                using var wordDoc = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());

                AddBilingualStyles(mainPart);

                foreach (var kvp in pageLayouts)
                {
                    var pageNum = kvp.Key;
                    var pageHeading = new Paragraph(new Run(new Text($"第 {pageNum} 页")))
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            Justification = new Justification { Val = JustificationValues.Center },
                            ParagraphBorders = new ParagraphBorders
                            {
                                BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 12 }
                            }
                        }
                    };
                    body.AppendChild(pageHeading);

                    // 合并相邻的文本项以重建段落
                    var mergedItems = MergeTextItems(kvp.Value);

                    if (_outputConfig.Format == TranslationOutputConfig.OutputFormat.BilingualSideBySide)
                    {
                        // 双语对照模式：使用表格布局整个页面
                        var table = new Table();
                        
                        // 设置表格属性：宽度100%
                        var tblPr = new TableProperties(
                            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                            new TableBorders(
                                new TopBorder { Val = BorderValues.Single, Size = 4 },
                                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                new InsideHorizontalBorder { Val = BorderValues.Dotted, Size = 4 },
                                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                            )
                        );
                        table.AppendChild(tblPr);

                        // 添加表头
                        var headerTr = new TableRow();
                        headerTr.Append(
                            new TableCell(new Paragraph(new Run(new Text("原文")) { RunProperties = new RunProperties { Bold = new Bold() } })),
                            new TableCell(new Paragraph(new Run(new Text("译文")) { RunProperties = new RunProperties { Bold = new Bold() } }))
                        );
                        table.AppendChild(headerTr);

                        foreach (var item in mergedItems)
                        {
                            if (item is ImageLayoutItem imgItem)
                            {
                                // 图片单独占一行，合并单元格
                                var tr = new TableRow();
                                var tc = new TableCell(new Paragraph(new Run(new Text($"[图片: {Path.GetFileName(imgItem.Path)}]"))));
                                var tcPr = new TableCellProperties(new GridSpan { Val = 2 });
                                tc.Append(tcPr);
                                tr.Append(tc);
                                table.Append(tr);
                            }
                            else if (item is TextLayoutItem textItem)
                            {
                                var original = textItem.Content;
                                var translated = string.Concat(textItem.Segments.Select(s => (s.Translated ?? string.Empty) + (s.Separator ?? string.Empty)));

                                var tr = new TableRow();
                                var tc1 = new TableCell(CreateStyledParagraph(original, "OriginalText", textItem));
                                var tc2 = new TableCell(CreateStyledParagraph(translated, "TranslatedText", textItem));
                                tr.Append(tc1, tc2);
                                table.Append(tr);
                            }
                        }
                        body.AppendChild(table);
                    }
                    else
                    {
                        // 垂直对照或仅译文模式
                        foreach (var item in mergedItems)
                        {
                            if (item is ImageLayoutItem imgItem)
                            {
                                var para = new Paragraph();
                                var run = new Run();
                                run.AppendChild(new Text($"[图片: {Path.GetFileName(imgItem.Path)}]"));
                                para.AppendChild(run);
                                body.AppendChild(para);
                            }
                            else if (item is TextLayoutItem textItem)
                            {
                                var original = textItem.Content;
                                var translated = string.Concat(textItem.Segments.Select(s => (s.Translated ?? string.Empty) + (s.Separator ?? string.Empty)));

                                // 应用样式
                                body.AppendChild(CreateStyledParagraph(original, "OriginalText", textItem));
                                
                                if (!string.IsNullOrWhiteSpace(translated))
                                {
                                    var transPara = CreateStyledParagraph(translated, "TranslatedText", textItem);
                                    // 译文颜色设为蓝色以区分
                                    if (transPara.Descendants<Run>().Any())
                                    {
                                        var run = transPara.Descendants<Run>().First();
                                        run.RunProperties ??= new RunProperties();
                                        run.RunProperties.Color = new Color { Val = "0000FF" };
                                    }
                                    body.AppendChild(transPara);
                                }
                                
                                body.AppendChild(new Paragraph()); // 段间距
                            }
                        }
                    }
                    
                    // 分页符
                    body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                }
                mainPart.Document.Save();
            });
        }

        private List<LayoutItem> MergeTextItems(List<LayoutItem> items)
        {
            var result = new List<LayoutItem>();
            TextLayoutItem currentText = null;

            foreach (var item in items)
            {
                if (item is TextLayoutItem textItem)
                {
                    if (currentText == null)
                    {
                        currentText = textItem;
                    }
                    else
                    {
                        // 判断是否应该合并：
                        // 1. 字体大小相近
                        // 2. 垂直距离在一定范围内（说明是同一段落的下一行）
                        // 3. 字体名称相同
                        bool sameFont = currentText.FontName == textItem.FontName;
                        bool sameSize = Math.Abs(currentText.FontSize - textItem.FontSize) < 2;
                        float verticalDist = currentText.Y - textItem.Y; // Y是Top Y，下一行Y更小
                        bool isNextLine = verticalDist > 0 && verticalDist < currentText.FontSize * 2.5; // 允许2.5倍行距
                        
                        if (sameFont && sameSize && isNextLine)
                        {
                            // 合并
                            currentText.Content += " " + textItem.Content;
                            currentText.Segments.AddRange(textItem.Segments);
                            // 更新Y坐标为最新的（虽然对于段落来说通常取第一行Y，但这里不重要）
                        }
                        else
                        {
                            // 结束当前段落
                            result.Add(currentText);
                            currentText = textItem;
                        }
                    }
                }
                else
                {
                    // 遇到非文本（如图片），先结束当前文本段落
                    if (currentText != null)
                    {
                        result.Add(currentText);
                        currentText = null;
                    }
                    result.Add(item);
                }
            }

            if (currentText != null)
            {
                result.Add(currentText);
            }

            return result;
        }

        private Paragraph CreateStyledParagraph(string text, string styleId, TextLayoutItem layoutItem = null)
        {
             var run = new Run(new Text(text));
             var runProps = new RunProperties();

             if (layoutItem != null)
             {
                 // 设置字号 (OpenXML使用半点，所以 * 2)
                 if (layoutItem.FontSize > 0)
                 {
                     runProps.FontSize = new FontSize { Val = (layoutItem.FontSize * 2).ToString("0") };
                 }

                 // 粗体检测
                 if (!string.IsNullOrEmpty(layoutItem.FontName) && 
                    (layoutItem.FontName.ToLower().Contains("bold") || layoutItem.FontName.Contains("黑体")))
                 {
                     runProps.Bold = new Bold();
                 }
             }

             run.RunProperties = runProps;

             return new Paragraph(run)
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = styleId }
                    }
                };
        }

        private async Task GeneratePdfOutputAsync(string outputPath, Dictionary<int, List<LayoutItem>> pageLayouts, string originalPdfPath)
        {
             await Task.Run(() =>
             {
                 try
                 {
                     using var pdfWriter = new PdfWriter(outputPath);
                     using var pdfDoc = new PdfDocument(pdfWriter);
                     using var originalPdfReader = new PdfReader(originalPdfPath);
                     using var originalPdfDoc = new PdfDocument(originalPdfReader);
                     
                     int totalPages = originalPdfDoc.GetNumberOfPages();
                     
                     var fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
                     var msyhPath = Path.Combine(fontsFolder, "msyh.ttc");
                     var simheiPath = Path.Combine(fontsFolder, "simhei.ttf");
                     
                     string fontToLoad = null;
                     if (File.Exists(msyhPath))
                     {
                         fontToLoad = msyhPath + ",0";
                     }
                     else if (File.Exists(simheiPath))
                     {
                         fontToLoad = simheiPath;
                     }

                     iText.Kernel.Font.PdfFont font = null;
                     if (fontToLoad != null)
                     {
                         font = iText.Kernel.Font.PdfFontFactory.CreateFont(fontToLoad, iText.IO.Font.PdfEncodings.IDENTITY_H);
                     }
                     
                     for (int i = 1; i <= totalPages; i++)
                     {
                         var originalPage = originalPdfDoc.GetPage(i);
                         var pageSize = originalPage.GetPageSize();
                         
                         var newPage = pdfDoc.AddNewPage(new iText.Kernel.Geom.PageSize(pageSize));
                         var canvas = new iText.Kernel.Pdf.Canvas.PdfCanvas(newPage);
                         
                         if (pageLayouts.ContainsKey(i))
                         {
                             foreach(var item in pageLayouts[i])
                             {
                                 if (item is TextLayoutItem textItem)
                                 {
                                     var text = string.Concat(textItem.Segments.Select(s => (s.Translated ?? string.Empty) + (s.Separator ?? string.Empty)));
                                     if (string.IsNullOrEmpty(text)) text = textItem.Content;
                                     
                                     if (font != null)
                                    {
                                        canvas.BeginText();
                                        float fontSize = textItem.FontSize > 0 ? textItem.FontSize : 9;
                                        canvas.SetFontAndSize(font, fontSize);
                                        
                                        var cleanText = text.Replace("\r\n", "\n").Replace("\r", "\n");
                                        var lines = cleanText.Split('\n');
                                        float lineHeight = fontSize * 1.2f;
                                        float startY = pageSize.GetHeight() - textItem.Y - fontSize + 2; // +2 adjustment

                                        for(int lineIdx = 0; lineIdx < lines.Length; lineIdx++)
                                        {
                                            canvas.SetTextMatrix(textItem.X, startY - (lineIdx * lineHeight));
                                            canvas.ShowText(lines[lineIdx]);
                                        }
                                        canvas.EndText();
                                    }
                                 }
                                 else if (item is ImageLayoutItem imgItem)
                                 {
                                     var imgPath = Path.Combine(Path.GetDirectoryName(outputPath), "日志", "图片", imgItem.Path);
                                     if (File.Exists(imgPath))
                                     {
                                         try {
                                             var imgData = iText.IO.Image.ImageDataFactory.Create(imgPath);
                                             float bottomY = pageSize.GetHeight() - imgItem.Y; // Y is Top Y
                                              // Correct: Top Y is from Top. 
                                              // Wait, CompositeLayoutStrategy: Y = y + h. (y is bottom). So Y is Top.
                                              // So Bottom Y = Y - Height.
                                              // But if Y is measured from Bottom (which is how ascent works), then Y is Top Y from Bottom.
                                              // So Bottom Y = Y - Height.
                                              
                                             canvas.AddImageAt(imgData, imgItem.X, bottomY - imgItem.Height, true);
                                         } catch {}
                                     }
                                 }
                             }
                         }
                     }
                 }
                 catch(Exception ex)
                 {
                     _logger.LogWarning(ex, "PDF generation failed");
                 }
             });
        }

        private List<LayoutItem> PerformOcrOnImages(List<ImageLayoutItem> images, int pageNum, string imagesDir)
        {
            var items = new List<LayoutItem>();
            if (_ocrEngine == null) return items;

            foreach(var img in images)
            {
                try
                {
                    var fullPath = Path.Combine(imagesDir, img.Path);
                    if (!File.Exists(fullPath)) continue;

                    using var pix = Pix.LoadFromFile(fullPath);
                    using var page = _ocrEngine.Process(pix, PageSegMode.Auto);
                    var text = page.GetText();
                    var meanConfidence = page.GetMeanConfidence();

                    if (!string.IsNullOrWhiteSpace(text) && meanConfidence > 0.6)
                    {
                        items.Add(new TextLayoutItem
                        {
                            Page = pageNum,
                            Content = text.Trim(),
                            X = img.X,
                            Y = img.Y,
                            Width = img.Width,
                            Height = img.Height,
                            FontName = "OCR_Fallback",
                            FontSize = 10 // Default size
                        });
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"OCR failed for image {img.Path}");
                }
            }
            return items;
        }

        private void AddBilingualVerticalContent(Body body, List<string> originalParagraphs, List<string> translatedParagraphs)
        {
            for (int paraIdx = 0; paraIdx < originalParagraphs.Count; paraIdx++)
            {
                var originalPara = originalParagraphs[paraIdx];
                var translatedPara = paraIdx < translatedParagraphs.Count ? translatedParagraphs[paraIdx] : string.Empty;

                var originalParaElement = new Paragraph(new Run(new Text(originalPara)))
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "OriginalText" }
                    }
                };
                body.AppendChild(originalParaElement);

                if (!string.IsNullOrWhiteSpace(translatedPara))
                {
                    var translatedParaElement = new Paragraph(new Run(new Text(translatedPara)))
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            ParagraphStyleId = new ParagraphStyleId { Val = "TranslatedText" }
                        }
                    };
                    body.AppendChild(translatedParaElement);
                }
            }
        }

        private void AddBilingualSideBySideContent(Body body, List<string> originalParagraphs, List<string> translatedParagraphs)
        {
            var table = new Table();
            var tableProperties = new TableProperties();
            var tableWidth = new TableWidth { Width = "100%", Type = TableWidthUnitValues.Pct };
            tableProperties.AppendChild(tableWidth);
            tableProperties.AppendChild(new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 1 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 1 }
            ));
            table.AppendChild(tableProperties);

            var tableGrid = new TableGrid();
            tableGrid.AppendChild(new GridColumn { Width = "50%" });
            tableGrid.AppendChild(new GridColumn { Width = "50%" });
            table.AppendChild(tableGrid);

            for (int paraIdx = 0; paraIdx < originalParagraphs.Count; paraIdx++)
            {
                var originalPara = originalParagraphs[paraIdx];
                var translatedPara = paraIdx < translatedParagraphs.Count ? translatedParagraphs[paraIdx] : string.Empty;

                var tr = new TableRow();

                var originalCell = new TableCell();
                var originalParaElement = new Paragraph(new Run(new Text(originalPara)))
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "OriginalText" }
                    }
                };
                originalCell.AppendChild(originalParaElement);
                tr.AppendChild(originalCell);

                var translatedCell = new TableCell();
                if (!string.IsNullOrWhiteSpace(translatedPara))
                {
                    var translatedParaElement = new Paragraph(new Run(new Text(translatedPara)))
                    {
                        ParagraphProperties = new ParagraphProperties
                        {
                            ParagraphStyleId = new ParagraphStyleId { Val = "TranslatedText" }
                        }
                    };
                    translatedCell.AppendChild(translatedParaElement);
                }
                tr.AppendChild(translatedCell);

                table.AppendChild(tr);
            }

            body.AppendChild(table);
        }

        private void AddInlineContent(Body body, List<string> originalParagraphs, List<string> translatedParagraphs)
        {
            for (int paraIdx = 0; paraIdx < originalParagraphs.Count; paraIdx++)
            {
                var originalPara = originalParagraphs[paraIdx];
                var translatedPara = paraIdx < translatedParagraphs.Count ? translatedParagraphs[paraIdx] : string.Empty;

                var runs = new List<Run>();
                runs.Add(new Run(new Text(originalPara)));

                if (!string.IsNullOrWhiteSpace(translatedPara))
                {
                    runs.Add(new Run(new Text(" (")));
                    var translatedRun = new Run(new Text(translatedPara))
                    {
                        RunProperties = new RunProperties
                        {
                            Color = new Color { Val = _outputConfig.TranslatedStyle.Color }
                        }
                    };
                    if (_outputConfig.TranslatedStyle.Italic)
                    {
                        translatedRun.RunProperties.Append(new Italic());
                    }
                    runs.Add(translatedRun);
                    runs.Add(new Run(new Text(")")));
                }

                var paraElement = new Paragraph();
                foreach (var run in runs)
                {
                    paraElement.AppendChild(run);
                }
                body.AppendChild(paraElement);
            }
        }

        private void AddOriginalOnlyContent(Body body, List<string> originalParagraphs)
        {
            foreach (var originalPara in originalParagraphs)
            {
                var paraElement = new Paragraph(new Run(new Text(originalPara)))
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "OriginalText" }
                    }
                };
                body.AppendChild(paraElement);
            }
        }

        private void AddTranslatedOnlyContent(Body body, List<string> translatedParagraphs)
        {
            foreach (var translatedPara in translatedParagraphs)
            {
                var paraElement = new Paragraph(new Run(new Text(translatedPara)))
                {
                    ParagraphProperties = new ParagraphProperties
                    {
                        ParagraphStyleId = new ParagraphStyleId { Val = "TranslatedText" }
                    }
                };
                body.AppendChild(paraElement);
            }
        }

        private void AddBilingualStyles(MainDocumentPart mainPart)
        {
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();

            var normalStyle = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal"
            };
            normalStyle.Append(new Name { Val = "Normal" });
            stylesPart.Styles.Append(normalStyle);

            var originalStyle = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "OriginalText"
            };
            originalStyle.Append(new Name { Val = "Original Text" });
            originalStyle.Append(new BasedOn { Val = "Normal" });
            originalStyle.Append(new NextParagraphStyle { Val = "TranslatedText" });
            
            var originalRunProperties = new RunProperties
            {
                Color = new Color { Val = _outputConfig.OriginalStyle.Color },
                FontSize = new FontSize { Val = (_outputConfig.OriginalStyle.FontSize * 2).ToString() }
            };
            if (_outputConfig.OriginalStyle.Bold)
            {
                originalRunProperties.Append(new Bold());
            }
            if (_outputConfig.OriginalStyle.Italic)
            {
                originalRunProperties.Append(new Italic());
            }
            originalStyle.Append(originalRunProperties);
            stylesPart.Styles.Append(originalStyle);

            var translatedStyle = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "TranslatedText"
            };
            translatedStyle.Append(new Name { Val = "Translated Text" });
            translatedStyle.Append(new BasedOn { Val = "Normal" });
            translatedStyle.Append(new NextParagraphStyle { Val = "OriginalText" });
            
            var translatedRunProperties = new RunProperties
            {
                Color = new Color { Val = _outputConfig.TranslatedStyle.Color },
                FontSize = new FontSize { Val = (_outputConfig.TranslatedStyle.FontSize * 2).ToString() }
            };
            if (_outputConfig.TranslatedStyle.Bold)
            {
                translatedRunProperties.Append(new Bold());
            }
            if (_outputConfig.TranslatedStyle.Italic)
            {
                translatedRunProperties.Append(new Italic());
            }
            translatedStyle.Append(translatedRunProperties);
            stylesPart.Styles.Append(translatedStyle);

            var highlightStyle = new Style
            {
                Type = StyleValues.Character,
                StyleId = "TermHighlight"
            };
            highlightStyle.Append(new Name { Val = "Term Highlight" });
            
            var highlightRunProperties = new RunProperties
            {
                Color = new Color { Val = _outputConfig.HighlightStyle.Color }
            };
            highlightRunProperties.Append(new Bold());
            if (_outputConfig.HighlightStyle.Underline)
            {
                highlightRunProperties.Append(new Underline { Val = UnderlineValues.Single });
            }
            highlightStyle.Append(highlightRunProperties);
            stylesPart.Styles.Append(highlightStyle);

            stylesPart.Styles.Save();
        }

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private string ExtractTextFromPage(PdfPage page, int pageNum)
        {
            var strategy = new LocationTextExtractionStrategy();
            var text = PdfTextExtractor.GetTextFromPage(page, strategy);

            text = PostProcessExtractedText(text);

            if (_useOcr && string.IsNullOrWhiteSpace(text))
            {
                _logger.LogInformation($"第 {pageNum} 页文本为空，尝试使用OCR识别");
                text = ExtractTextFromPageWithOcr(page, pageNum);
            }
            else if (_useOcr && text.Length < 100)
            {
                _logger.LogInformation($"第 {pageNum} 页文本过短（{text.Length}字符），可能为扫描版PDF，尝试使用OCR识别");
                var ocrText = ExtractTextFromPageWithOcr(page, pageNum);
                if (ocrText.Length > text.Length)
                {
                    text = ocrText;
                    _logger.LogInformation($"第 {pageNum} 页使用OCR识别结果（{ocrText.Length}字符）");
                }
            }

            _logger.LogDebug($"第 {pageNum} 页提取文本: {text.Length} 个字符");
            return text;
        }

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private string ExtractTextFromPageWithOcr(PdfPage page, int pageNum)
        {
            try
            {
                if (_ocrEngine == null)
                {
                    _logger.LogWarning("OCR引擎未初始化");
                    return string.Empty;
                }

                var ocrText = new StringBuilder();

                var images = ExtractImagesFromPage(page, pageNum);
                foreach (var imageInfo in images)
                {
                    try
                    {
                        using var image = System.Drawing.Image.FromStream(new MemoryStream(imageInfo.ImageData));
                        
                        using var bmp = new System.Drawing.Bitmap(image.Width, image.Height);
                        using var g = System.Drawing.Graphics.FromImage(bmp);
                        g.DrawImage(image, 0, 0);

                        var tempDir = Path.Combine(Path.GetTempPath(), "OCR_Temp");
                        if (!Directory.Exists(tempDir))
                        {
                            Directory.CreateDirectory(tempDir);
                        }

                        var tempImagePath = Path.Combine(tempDir, $"temp_{Guid.NewGuid()}.png");
                        bmp.Save(tempImagePath, System.Drawing.Imaging.ImageFormat.Png);

                        try
                        {
                            using var pix = Pix.LoadFromFile(tempImagePath);
                            using var pageResult = _ocrEngine.Process(pix, PageSegMode.Auto);
                            var text = pageResult.GetText();

                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                ocrText.AppendLine(text.Trim());
                            }
                        }
                        finally
                        {
                            if (File.Exists(tempImagePath))
                            {
                                File.Delete(tempImagePath);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"OCR处理图片失败: 第{pageNum}页图片{imageInfo.Index}");
                    }
                }

                var result = ocrText.ToString().Trim();
                _logger.LogInformation($"第 {pageNum} 页OCR识别结果: {result.Length} 个字符");
                return result;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"第 {pageNum} 页OCR识别失败");
                return string.Empty;
            }
        }

        private string PostProcessExtractedText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            text = Regex.Replace(text, @"\r\n", "\n");
            text = Regex.Replace(text, @"\n{3,}", "\n\n");
            text = Regex.Replace(text, @"\s{4,}", "    ");
            text = text.Trim();

            return text;
        }

        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private List<string> ExtractTextWithFormatting(PdfPage page, int pageNum)
        {
            var strategy = new LocationTextExtractionStrategy();
            var text = PdfTextExtractor.GetTextFromPage(page, strategy);

            text = PostProcessExtractedText(text);

            var paragraphs = new List<string>();
            var lines = text.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            var currentParagraph = new StringBuilder();
            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                
                if (string.IsNullOrWhiteSpace(trimmedLine))
                {
                    if (currentParagraph.Length > 0)
                    {
                        paragraphs.Add(currentParagraph.ToString().Trim());
                        currentParagraph.Clear();
                    }
                }
                else
                {
                    bool isHeading = IsHeading(trimmedLine);
                    
                    if (isHeading)
                    {
                        if (currentParagraph.Length > 0)
                        {
                            paragraphs.Add(currentParagraph.ToString().Trim());
                            currentParagraph.Clear();
                        }
                        currentParagraph.Append(trimmedLine);
                    }
                    else if (currentParagraph.Length > 0)
                    {
                        currentParagraph.Append(" ");
                        currentParagraph.Append(trimmedLine);
                    }
                    else
                    {
                        currentParagraph.Append(trimmedLine);
                    }
                }
            }

            if (currentParagraph.Length > 0)
            {
                paragraphs.Add(currentParagraph.ToString().Trim());
            }

            if (paragraphs.Count == 0)
            {
                _logger.LogWarning($"第 {pageNum} 页未提取到段落，尝试使用整页文本");
                if (!string.IsNullOrWhiteSpace(text))
                {
                    paragraphs.Add(text.Trim());
                }
            }

            _logger.LogDebug($"第 {pageNum} 页提取到 {paragraphs.Count} 个段落");
            return paragraphs;
        }

        private bool IsHeading(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return false;

            var trimmed = text.Trim();

            if (Regex.IsMatch(trimmed, @"^[一二三四五六七八九十]+、"))
                return true;

            if (Regex.IsMatch(trimmed, @"^\d+\.\s+\S+"))
                return true;

            if (Regex.IsMatch(trimmed, @"^[IVX]+\."))
                return true;

            if (Regex.IsMatch(trimmed, @"^\d+\s+"))
                return true;

            if (Regex.IsMatch(trimmed, @"^[IVX]+\s+"))
                return true;

            return false;
        }

        private string HighlightTerms(string text, Dictionary<string, string> terminology)
        {
            if (!_outputConfig.HighlightTerms || terminology == null || terminology.Count == 0)
            {
                return text;
            }

            var result = text;
            foreach (var term in terminology.OrderByDescending(t => t.Key.Length))
            {
                if (result.Contains(term.Key))
                {
                    result = result.Replace(term.Key, $"[TERM_HIGHLIGHT_{term.Key}]");
                }
            }

            return result;
        }

        private string RestoreTermHighlights(string text, Dictionary<string, string> terminology)
        {
            var result = text;
            foreach (var term in terminology)
            {
                var placeholder = $"[TERM_HIGHLIGHT_{term.Key}]";
                if (result.Contains(placeholder))
                {
                    result = result.Replace(placeholder, term.Key);
                }
            }

            return result;
        }

        private List<string> SplitIntoSegments(string text)
        {
            // 如果配置为按句子分段，直接使用句子分段
            if (_outputConfig.Segmentation == TranslationOutputConfig.SegmentationMode.Sentence)
            {
                return SplitIntoSentences(text);
            }
            
            // 如果文本过长（超过1500字符），强制进行智能分段，避免模型处理失败
            // 即使在"按段落"模式下，过长的段落也需要拆分
            const int MaxSegmentLength = 1500;
            if (text.Length > MaxSegmentLength)
            {
                _logger.LogInformation($"文本长度({text.Length})超过限制({MaxSegmentLength})，执行智能分段");
                return SmartSplitText(text, MaxSegmentLength);
            }

            return new List<string> { text };
        }

        private List<string> SmartSplitText(string text, int maxLength)
        {
            var segments = new List<string>();
            if (string.IsNullOrEmpty(text)) return segments;

            var currentStart = 0;
            while (currentStart < text.Length)
            {
                // 如果剩余文本小于最大长度，直接添加
                if (text.Length - currentStart <= maxLength)
                {
                    segments.Add(text.Substring(currentStart));
                    break;
                }

                // 寻找最佳分割点
                // 从 maxLength 处向前寻找句子结束符
                var splitIndex = -1;
                var searchEnd = Math.Min(currentStart + maxLength, text.Length);
                var searchStart = Math.Max(currentStart + maxLength / 2, currentStart); // 只在后半部分寻找

                // 优先级1: 换行符
                var lastNewLine = text.LastIndexOf('\n', searchEnd - 1, searchEnd - searchStart);
                if (lastNewLine != -1)
                {
                    splitIndex = lastNewLine + 1;
                }
                else
                {
                    // 优先级2: 句子结束符
                    for (int i = searchEnd - 1; i >= searchStart; i--)
                    {
                        if (IsSentenceEnd(text, i))
                        {
                            splitIndex = i + 1;
                            break;
                        }
                    }
                }

                // 优先级3: 逗号或空格 (如果没有找到句子结束符)
                if (splitIndex == -1)
                {
                    var lastComma = text.LastIndexOfAny(new[] { '，', ',', ' ', '\t' }, searchEnd - 1, searchEnd - searchStart);
                    if (lastComma != -1)
                    {
                        splitIndex = lastComma + 1;
                    }
                }

                // 如果实在找不到合适的分割点，强制分割
                if (splitIndex == -1)
                {
                    splitIndex = searchEnd;
                }

                segments.Add(text.Substring(currentStart, splitIndex - currentStart).Trim());
                currentStart = splitIndex;
            }

            return segments;
        }

        private List<string> SplitIntoSentences(string text)
        {
            var sentences = new List<string>();
            var currentSentence = new StringBuilder();

            for (int i = 0; i < text.Length; i++)
            {
                currentSentence.Append(text[i]);

                if (IsSentenceEnd(text, i))
                {
                    var sentence = currentSentence.ToString().Trim();
                    if (!string.IsNullOrWhiteSpace(sentence))
                    {
                        sentences.Add(sentence);
                    }
                    currentSentence.Clear();
                }
            }

            if (currentSentence.Length > 0)
            {
                var remaining = currentSentence.ToString().Trim();
                if (!string.IsNullOrWhiteSpace(remaining))
                {
                    sentences.Add(remaining);
                }
            }

            return sentences;
        }

        private bool IsSentenceEnd(string text, int index)
        {
            var currentChar = text[index];

            if (currentChar == '。' || currentChar == '！' || currentChar == '？' || 
                currentChar == '.' || currentChar == '!' || currentChar == '?')
            {
                return true;
            }

            if (currentChar == '；' || currentChar == ';')
            {
                return true;
            }

            return false;
        }

        private List<ImageInfo> ExtractImagesFromPage(PdfPage page, int pageNum)
        {
            var images = new List<ImageInfo>();
            var listener = new ImageRenderListener();

            var canvasProcessor = new PdfCanvasProcessor(listener);
            canvasProcessor.ProcessPageContent(page);

            images = listener.GetImages().Select((img, idx) => new ImageInfo
            {
                PageNum = pageNum,
                Index = idx + 1,
                ImageData = img
            }).ToList();

            if (images.Count > 0)
            {
                _logger.LogInformation($"第 {pageNum} 页提取到 {images.Count} 张图片");
            }

            return images;
        }

        private (string text, List<(string placeholder, string formula)> placeholders) ProtectMathFormulas(string text)
        {
            var placeholders = new List<(string placeholder, string formula)>();
            var processedText = text;
            var placeholderIndex = 0;

            foreach (var pattern in _latexPatterns)
            {
                var matches = pattern.Matches(processedText);
                foreach (Match match in matches)
                {
                    var formula = match.Value;
                    var placeholder = $"[MATH_FORMULA_{placeholderIndex++}]";

                    placeholders.Add((placeholder, formula));
                    processedText = processedText.Replace(formula, placeholder);
                }
            }

            _logger.LogInformation($"保护了 {placeholders.Count} 个数学公式");
            return (processedText, placeholders);
        }

        private string RestoreMathFormulas(string text, List<(string placeholder, string formula)> placeholders)
        {
            var result = text;

            foreach (var (placeholder, formula) in placeholders.OrderByDescending(p => p.placeholder.Length))
            {
                result = result.Replace(placeholder, formula);
            }

            _logger.LogInformation($"恢复了 {placeholders.Count} 个数学公式");
            return result;
        }

        private async Task<string> TranslateTextWithRetryAsync(string text, Dictionary<string, string> terminology, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var src = string.IsNullOrWhiteSpace(_sourceLang) ? "auto" : _sourceLang;
            var tgt = string.IsNullOrWhiteSpace(_targetLang) ? "en" : _targetLang;

            for (int attempt = 1; attempt <= _retryCount; attempt++)
            {
                try
                {
                    var translated = await _translationService.TranslateTextAsync(text, terminology, src, tgt, null, originalText);
                    if (!string.IsNullOrEmpty(translated))
                    {
                        return translated;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"翻译尝试 {attempt}/{_retryCount} 失败: {ex.Message}");
                }

                if (attempt < _retryCount)
                {
                    await Task.Delay(_retryDelay);
                }
            }

            _logger.LogError($"翻译失败，已重试 {_retryCount} 次");
            
            // 如果提供了原文，在翻译失败时返回原文
            if (!string.IsNullOrEmpty(originalText))
            {
                _logger.LogWarning("翻译失败，使用原文作为回退结果");
                return originalText;
            }
            
            return string.Empty;
        }

        private void AddImageToDocument(Body body, ImageInfo imageInfo)
        {
            try
            {
                if (body == null)
                {
                    _logger.LogWarning("Body为null，无法添加图片");
                    return;
                }

                var para = new Paragraph();
                var run = new Run();

                var imageText = $"[图片 {imageInfo.PageNum}-{imageInfo.Index}]";
                run.AppendChild(new Text(imageText));
                para.AppendChild(run);
                body.AppendChild(para);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"添加图片失败: 页面 {imageInfo.PageNum}, 图片 {imageInfo.Index}");
            }
        }

        private void UpdateProgress(double progress, string message)
        {
            try
            {
                var now = DateTime.Now;
                var elapsedSinceLastUpdate = (now - _lastProgressUpdate).TotalMilliseconds;

                if (elapsedSinceLastUpdate >= _minProgressUpdateInterval || progress >= 1.0)
                {
                    _lastProgressUpdate = now;
                    _progressCallback?.Invoke(progress, message);
                    _logger.LogInformation($"进度: {progress * 100:F1}% - {message}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "更新进度失败");
            }
        }

        private async Task<string> PreprocessWithAdobeAcrobatAsync(string pdfPath)
        {
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), "AdobeAcrobatPreprocess");
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                var fileName = Path.GetFileNameWithoutExtension(pdfPath);
                var wordPath = Path.Combine(tempDir, $"{fileName}_temp.docx");

                _logger.LogInformation($"尝试使用Adobe Acrobat将PDF转换为Word: {pdfPath}");

                var acrobatPath = FindAdobeAcrobatPath();
                if (string.IsNullOrEmpty(acrobatPath))
                {
                    _logger.LogWarning("未找到Adobe Acrobat，跳过预处理");
                    return pdfPath;
                }

                var scriptPath = Path.Combine(tempDir, $"convert_{Guid.NewGuid()}.vbs");
                var scriptContent = $@"
Set AcroApp = CreateObject(""AcroExch.App"")
Set AVDoc = CreateObject(""AcroExch.AVDoc"")
Set PDDoc = CreateObject(""AcroExch.PDDoc"")

If AVDoc.Open(""{pdfPath}"", ""Acrobat"") Then
    Set PDDoc = AVDoc.GetPDDoc
    If PDDoc.Save(1, ""{wordPath}"") Then
        WScript.Echo ""SUCCESS""
    Else
        WScript.Echo ""FAILED""
    End If
    PDDoc.Close
    AVDoc.Close True
Else
    WScript.Echo ""FAILED""
End If

AcroApp.Exit
Set AcroApp = Nothing
Set AVDoc = Nothing
Set PDDoc = Nothing
";

                File.WriteAllText(scriptPath, scriptContent, Encoding.Default);

                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "wscript.exe",
                    Arguments = $"\"{scriptPath}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var process = System.Diagnostics.Process.Start(processStartInfo);
                if (process != null)
                {
                    var output = await process.StandardOutput.ReadToEndAsync();
                    var error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    if (process.ExitCode == 0 && output.Contains("SUCCESS"))
                    {
                        _logger.LogInformation($"Adobe Acrobat转换成功: {wordPath}");
                        return wordPath;
                    }
                    else
                    {
                        _logger.LogWarning($"Adobe Acrobat转换失败: {error}");
                    }
                }

                try
                {
                    File.Delete(scriptPath);
                }
                catch { }

                return pdfPath;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Adobe Acrobat预处理失败，使用原始PDF");
                return pdfPath;
            }
        }

        private string FindAdobeAcrobatPath()
        {
            var possiblePaths = new[]
            {
                @"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                @"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                @"C:\Program Files\Adobe\Acrobat 2020\Acrobat\Acrobat.exe",
                @"C:\Program Files (x86)\Adobe\Acrobat 2020\Acrobat\Acrobat.exe",
                @"C:\Program Files\Adobe\Acrobat 2017\Acrobat\Acrobat.exe",
                @"C:\Program Files (x86)\Adobe\Acrobat 2017\Acrobat\Acrobat.exe"
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    _logger.LogInformation($"找到Adobe Acrobat: {path}");
                    return path;
                }
            }

            return null;
        }

        private async Task<string> ConvertWordToPdfAsync(string wordPath)
        {
            try
            {
                var pdfPath = Path.ChangeExtension(wordPath, ".pdf");
                _logger.LogInformation($"尝试将Word转换为PDF: {wordPath} -> {pdfPath}");

                var tempDir = Path.Combine(Path.GetTempPath(), "WordToPdfConversion");
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }

                var scriptPath = Path.Combine(tempDir, $"convert_word_{Guid.NewGuid()}.vbs");
                
                // VBScript to open Word and save as PDF
                // wdFormatPDF = 17
                var scriptContent = $@"
Set objWord = CreateObject(""Word.Application"")
objWord.Visible = False
objWord.DisplayAlerts = 0

Set objDoc = objWord.Documents.Open(""{wordPath}"")
objDoc.SaveAs2 ""{pdfPath}"", 17

objDoc.Close False
objWord.Quit

Set objDoc = Nothing
Set objWord = Nothing
WScript.Echo ""SUCCESS""
";

                File.WriteAllText(scriptPath, scriptContent, Encoding.Default);

                var processStartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "wscript.exe",
                    Arguments = $"\"{scriptPath}\"",
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using var process = System.Diagnostics.Process.Start(processStartInfo);
                if (process != null)
                {
                    var output = await process.StandardOutput.ReadToEndAsync();
                    var error = await process.StandardError.ReadToEndAsync();
                    await process.WaitForExitAsync();

                    if (process.ExitCode == 0 && output.Contains("SUCCESS"))
                    {
                        _logger.LogInformation($"Word转PDF成功: {pdfPath}");
                        
                        try { File.Delete(scriptPath); } catch { }
                        
                        return pdfPath;
                    }
                    else
                    {
                        _logger.LogWarning($"Word转PDF失败: {error}");
                        return null;
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Word转PDF转换异常");
                return null;
            }
        }

        private string EscapeCsv(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            text = text.Replace("\r\n", "\n").Replace("\r", "\n");
            if (text.Contains('"') || text.Contains(',') || text.Contains('\n'))
            {
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }
            return text;
        }

        private List<LayoutItem> AnalyzePageLayout(List<LayoutItem> rawItems)
        {
            if (rawItems == null || rawItems.Count == 0) return new List<LayoutItem>();

            // Detect two-column layout
            var textItems = rawItems.OfType<TextLayoutItem>().ToList();
            bool isTwoColumn = false;
            float pageCenter = 0;
            float headerThreshold = float.MaxValue;
            float footerThreshold = float.MinValue;

            if (textItems.Count > 10)
            {
                var minX = textItems.Min(i => i.X);
                var maxX = textItems.Max(i => i.X + i.Width);
                var pageWidth = maxX - minX;
                pageCenter = (minX + maxX) / 2;

                var leftCount = textItems.Count(i => (i.X + i.Width) < pageCenter);
                var rightCount = textItems.Count(i => i.X > pageCenter);

                // If significant content on both sides
                if (leftCount > textItems.Count * 0.3 && rightCount > textItems.Count * 0.3)
                {
                    isTwoColumn = true;
                    var minY = textItems.Min(i => i.Y);
                    var maxY = textItems.Max(i => i.Y);
                    // Top 10% and Bottom 10% as Header/Footer areas
                    headerThreshold = maxY - (maxY - minY) * 0.1f;
                    footerThreshold = minY + (maxY - minY) * 0.1f;
                }
            }

            List<LayoutItem> sortedItems;
            if (isTwoColumn)
            {
                var allItems = new List<LayoutItem>(rawItems);
                // Note: Y increases upwards in PDF. Max Y is Top.
                var headers = allItems.Where(i => i.Y > headerThreshold).OrderByDescending(i => i.Y).ThenBy(i => i.X).ToList();
                var footers = allItems.Where(i => i.Y < footerThreshold).OrderByDescending(i => i.Y).ThenBy(i => i.X).ToList();
                var body = allItems.Where(i => i.Y <= headerThreshold && i.Y >= footerThreshold).ToList();

                var bodyLeft = body.Where(i => (i.X + i.Width / 2) < pageCenter).OrderByDescending(i => i.Y).ThenBy(i => i.X).ToList();
                var bodyRight = body.Where(i => (i.X + i.Width / 2) >= pageCenter).OrderByDescending(i => i.Y).ThenBy(i => i.X).ToList();

                sortedItems = new List<LayoutItem>();
                sortedItems.AddRange(headers);
                sortedItems.AddRange(bodyLeft);
                sortedItems.AddRange(bodyRight);
                sortedItems.AddRange(footers);
            }
            else
            {
                // Standard sort: Top to Bottom, Left to Right
                sortedItems = rawItems.OrderByDescending(i => i.Y).ThenBy(i => i.X).ToList();
            }

            var mergedItems = new List<LayoutItem>();
            TextLayoutItem currentTextItem = null;

            foreach (var item in sortedItems)
            {
                if (item is ImageLayoutItem imgItem)
                {
                    if (currentTextItem != null)
                    {
                        mergedItems.Add(currentTextItem);
                        currentTextItem = null;
                    }
                    mergedItems.Add(imgItem);
                }
                else if (item is TextLayoutItem textItem)
                {
                    if (currentTextItem == null)
                    {
                        currentTextItem = textItem;
                    }
                    else
                    {
                        // Merge logic
                        float verticalGap = currentTextItem.Y - textItem.Y; // Positive if textItem is below
                        float height = currentTextItem.Height > 0 ? currentTextItem.Height : 10;

                        bool isSameLine = Math.Abs(currentTextItem.Y - textItem.Y) < (height * 0.5);
                        bool isNextLine = verticalGap > 0 && verticalGap < (height * 2.5);
                        bool sameFont = Math.Abs(currentTextItem.FontSize - textItem.FontSize) < 2;

                        if ((isSameLine || isNextLine) && sameFont)
                        {
                            if (isSameLine)
                            {
                                float horizontalGap = textItem.X - (currentTextItem.X + currentTextItem.Width);
                                if (horizontalGap > 2) 
                                {
                                    char lastChar = currentTextItem.Content.LastOrDefault();
                                    char nextChar = textItem.Content.FirstOrDefault();
                                    bool isChinese = (lastChar >= 0x4E00 && lastChar <= 0x9FFF) || (nextChar >= 0x4E00 && nextChar <= 0x9FFF);
                                    
                                    if (!isChinese) currentTextItem.Content += " ";
                                }
                                currentTextItem.Content += textItem.Content;
                                currentTextItem.Width = (textItem.X + textItem.Width) - currentTextItem.X;
                            }
                            else
                            {
                                // New line merge
                                char lastChar = currentTextItem.Content.LastOrDefault();
                                char nextChar = textItem.Content.FirstOrDefault();

                                if (currentTextItem.Content.EndsWith("-"))
                                {
                                    currentTextItem.Content = currentTextItem.Content.TrimEnd('-');
                                    currentTextItem.Content += textItem.Content;
                                }
                                else
                                {
                                     bool isChinese = (lastChar >= 0x4E00 && lastChar <= 0x9FFF) || (nextChar >= 0x4E00 && nextChar <= 0x9FFF);
                                     if (isChinese)
                                     {
                                         currentTextItem.Content += textItem.Content;
                                     }
                                     else
                                     {
                                         currentTextItem.Content += " " + textItem.Content;
                                     }
                                }
                                currentTextItem.Height += verticalGap;
                            }
                        }
                        else
                        {
                            mergedItems.Add(currentTextItem);
                            currentTextItem = textItem;
                        }
                    }
                }
            }

            if (currentTextItem != null)
            {
                mergedItems.Add(currentTextItem);
            }

            return mergedItems;
        }

        private List<string> SplitByMainPunctuation(string text)
        {
            var segments = new List<string>();
            if (string.IsNullOrWhiteSpace(text)) return segments;

            var sb = new StringBuilder();
            // Main punctuation marks: . ! ? ; 。 ！？ ；
            // Also consider \n as a hard break
            
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                sb.Append(c);

                if (IsMainPunctuation(c) || c == '\n')
                {
                    segments.Add(sb.ToString());
                    sb.Clear();
                }
            }

            if (sb.Length > 0)
            {
                var remaining = sb.ToString();
                if (!string.IsNullOrEmpty(remaining))
                {
                    segments.Add(remaining);
                }
            }

            return segments;
        }

        private readonly struct SentencePiece
        {
            public string Text { get; }
            public string Separator { get; }

            public SentencePiece(string text, string separator)
            {
                Text = text;
                Separator = separator;
            }
        }

        private List<SentencePiece> SplitParagraphToSentences(string text)
        {
            var result = new List<SentencePiece>();
            if (string.IsNullOrWhiteSpace(text)) return result;

            var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            var sb = new StringBuilder();

            int i = 0;
            while (i < normalized.Length)
            {
                var c = normalized[i];
                sb.Append(c);

                if (IsSentenceBoundary(normalized, i))
                {
                    int j = i + 1;
                    var sep = new StringBuilder();
                    while (j < normalized.Length && char.IsWhiteSpace(normalized[j]))
                    {
                        sep.Append(normalized[j]);
                        j++;
                    }

                    var sentence = sb.ToString().Trim();
                    if (sentence.Length > 0)
                    {
                        result.Add(new SentencePiece(sentence, sep.ToString()));
                    }
                    else if (result.Count > 0)
                    {
                        var prev = result[^1];
                        result[^1] = new SentencePiece(prev.Text, prev.Separator + sep);
                    }

                    sb.Clear();
                    i = j;
                    continue;
                }

                i++;
            }

            var tail = sb.ToString().Trim();
            if (tail.Length > 0)
            {
                result.Add(new SentencePiece(tail, ""));
            }

            return result;
        }

        private bool IsSentenceBoundary(string text, int index)
        {
            if (index < 0 || index >= text.Length) return false;
            char c = text[index];

            if (c == '。' || c == '！' || c == '？' || c == '；')
                return true;

            if (c == '!' || c == '?' || c == ';')
                return true;

            if (c != '.')
                return false;

            char prev = index > 0 ? text[index - 1] : '\0';
            char next = index + 1 < text.Length ? text[index + 1] : '\0';

            if (char.IsDigit(prev) && char.IsDigit(next))
                return false;

            var start = Math.Max(0, index - 6);
            var window = text.Substring(start, index - start + 1).ToLowerInvariant();
            if (window.EndsWith("e.g.") || window.EndsWith("i.e.") || window.EndsWith("fig.") || window.EndsWith("no.") ||
                window.EndsWith("dr.") || window.EndsWith("mr.") || window.EndsWith("ms.") || window.EndsWith("vs."))
                return false;

            return true;
        }

        private bool IsMainPunctuation(char c)
        {
            return c == '.' || c == '!' || c == '?' || c == ';' || 
                   c == '。' || c == '！' || c == '？' || c == '；';
        }

        private class ImageRenderListener : IEventListener
        {
            private readonly List<byte[]> _images = new List<byte[]>();

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_IMAGE)
                {
                    var imageData = data as ImageRenderInfo;
                    if (imageData != null)
                    {
                        try
                        {
                            var imageBytes = imageData.GetImage().GetImageBytes(true);
                            _images.Add(imageBytes);
                        }
                        catch
                        {
                        }
                    }
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new List<EventType> { EventType.RENDER_IMAGE };
            }

            public List<byte[]> GetImages()
            {
                return _images;
            }
        }

        private class ImageInfo
        {
            public int PageNum { get; set; }
            public int Index { get; set; }
            public byte[] ImageData { get; set; }
        }

        #region Layout Analysis Classes

        private class LayoutItem
        {
            public string Id { get; set; } = Guid.NewGuid().ToString("N");
            public string Type { get; set; } // "text" or "image"
            public int Page { get; set; }
            public float X { get; set; }
            public float Y { get; set; } // Top Y
            public float Width { get; set; }
            public float Height { get; set; }
        }

        private class TextLayoutItem : LayoutItem
        {
            public string Content { get; set; }
            public string FontName { get; set; }
            public float FontSize { get; set; }
            public List<TextSegment> Segments { get; set; } = new List<TextSegment>();

            public TextLayoutItem() { Type = "text"; }
        }

        private class TextSegment
        {
            public string Original { get; set; }
            public string Preprocessed { get; set; }
            public string Translated { get; set; }
            public string Separator { get; set; }
            public int LogIndex { get; set; }
        }

        private class ImageLayoutItem : LayoutItem
        {
            public string Path { get; set; } // Relative path in log folder
            public ImageLayoutItem() { Type = "image"; }
        }

        private class CompositeLayoutStrategy : IEventListener
        {
            private readonly List<LayoutItem> _items = new List<LayoutItem>();
            private readonly int _pageNum;
            private readonly string _imageOutputDir;

            public CompositeLayoutStrategy(int pageNum, string imageOutputDir)
            {
                _pageNum = pageNum;
                _imageOutputDir = imageOutputDir;
            }

            public List<LayoutItem> GetItems() => _items;

            public void EventOccurred(IEventData data, EventType type)
            {
                if (type == EventType.RENDER_TEXT)
                {
                    var textInfo = data as TextRenderInfo;
                    if (textInfo != null)
                    {
                        var text = textInfo.GetText();
                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            var baseline = textInfo.GetBaseline();
                            var ascent = textInfo.GetAscentLine();
                            var descent = textInfo.GetDescentLine();
                            
                            var start = baseline.GetStartPoint();
                            var end = baseline.GetEndPoint();
                            
                            var x = start.Get(0);
                            var y = ascent?.GetStartPoint().Get(1) ?? start.Get(1); // Top Y
                            var h = (ascent?.GetStartPoint().Get(1) ?? start.Get(1)) - (descent?.GetStartPoint().Get(1) ?? start.Get(1));
                            var w = end.Get(0) - start.Get(0);

                            var fontName = textInfo.GetFont()?.GetFontProgram()?.GetFontNames()?.GetFontName() ?? "Unknown";
                            var fontSize = textInfo.GetFontSize();

                            // Initial granular item, will be merged later
                            _items.Add(new TextLayoutItem
                            {
                                Page = _pageNum,
                                Content = text,
                                X = x,
                                Y = y,
                                Width = w,
                                Height = h,
                                FontName = fontName,
                                FontSize = fontSize
                            });
                        }
                    }
                }
                else if (type == EventType.RENDER_IMAGE)
                {
                    var imgInfo = data as ImageRenderInfo;
                    if (imgInfo != null)
                    {
                        try
                        {
                            var ctm = imgInfo.GetImageCtm();
                            // Extract coordinates from CTM
                            // [a b 0]
                            // [c d 0]
                            // [e f 1]
                            // e = x, f = y (bottom-left)
                            var x = ctm.Get(6);
                            var y = ctm.Get(7);
                            var w = ctm.Get(0);
                            var h = ctm.Get(4); // Or appropriate scaling

                            var imageBytes = imgInfo.GetImage().GetImageBytes(true);
                            if (imageBytes != null && imageBytes.Length > 0)
                            {
                                var fileName = $"Image_{_pageNum}_{_items.Count + 1}.png";
                                var fullPath = Path.Combine(_imageOutputDir, fileName);
                                File.WriteAllBytes(fullPath, imageBytes);

                                _items.Add(new ImageLayoutItem
                                {
                                    Page = _pageNum,
                                    Path = fileName,
                                    X = x,
                                    Y = y + h, // Convert to Top Y
                                    Width = w,
                                    Height = h
                                });
                            }
                        }
                        catch (Exception) { /* Ignore image errors */ }
                    }
                }
            }

            public ICollection<EventType> GetSupportedEvents()
            {
                return new List<EventType> { EventType.RENDER_TEXT, EventType.RENDER_IMAGE };
            }
        }
        #endregion
    }
}
