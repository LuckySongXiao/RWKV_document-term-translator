#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Shape = DocumentFormat.OpenXml.Presentation.Shape;
using GraphicFrame = DocumentFormat.OpenXml.Presentation.GraphicFrame;
using GroupShape = DocumentFormat.OpenXml.Presentation.GroupShape;
using Text = DocumentFormat.OpenXml.Presentation.Text;
using Path = System.IO.Path;
using System.Collections.Concurrent;
using System.Threading;
using System.Text;
using System.Globalization;
using ClosedXML.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// PowerPoint文档处理器，负责处理PPT文档的翻译
    /// </summary>
    public class PPTProcessor
    {
        private readonly TranslationService _translationService;
        private readonly ILogger<PPTProcessor> _logger;
        private readonly TermExtractor _termExtractor;
        private bool _useTerminology = true;
        private bool _preprocessTerms = true;
        private string _sourceLang = "zh";
        private string _targetLang = "en";
        private string _targetLanguageName = "英语";
        // 移除未使用的字段
        private bool _isCnToForeign = true;
        private string _outputFormat = "bilingual";
        private Action<double, string> _progressCallback;
        private int _retryCount = 1;
        private int _retryDelay = 1000; // 毫秒
        private bool _preserveFormatting = true;
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private const double _minProgressUpdateInterval = 100; // 最小进度更新间隔100ms
        private bool _translateNotes = true;
        private bool _translateCharts = true;
        private bool _translateSmartArt = true;
        private bool _autoAdjustLayout = true;
        private bool _useImmersiveStyle = false;

        // 数学公式正则表达式模式
        private readonly List<Regex> _latexPatterns = new List<Regex>
        {
            new Regex(@"\$\$(.*?)\$\$", RegexOptions.Singleline), // 行间公式 $$...$$
            new Regex(@"\$(.*?)\$", RegexOptions.Singleline),      // 行内公式 $...$
            new Regex(@"\\begin\{equation\}(.*?)\\end\{equation\}", RegexOptions.Singleline), // equation环境
            new Regex(@"\\begin\{align\}(.*?)\\end\{align\}", RegexOptions.Singleline),       // align环境
            new Regex(@"\\begin\{eqnarray\}(.*?)\\end\{eqnarray\}", RegexOptions.Singleline)  // eqnarray环境
        };

        // PPT特定的翻译提示词
        private readonly Dictionary<string, string> _pptPrompts = new Dictionary<string, string>
        {
            ["title"] = "这是一个PPT标题，请提供简洁明了的翻译，保持专业术语的准确性，确保标题简洁有力。",
            ["content"] = "这是一个PPT演示文稿的内容，请提供简洁明了的翻译，保持专业术语的准确性，并确保翻译后的文本长度适合在幻灯片上显示。",
            ["bullet"] = "这是PPT项目符号列表，请提供简洁的翻译，保持格式一致性。",
            ["chart"] = "这是PPT图表中的文本标签，请提供准确的翻译，保持专业术语一致性。",
            ["note"] = "这是PPT备注内容，请提供完整的翻译，保持专业性和可读性。",
            ["smartart"] = "这是PPT SmartArt图形中的文本，请提供简洁明了的翻译，保持专业术语准确性。"
        };

        // 沉浸式翻译风格的提示词
        private readonly Dictionary<string, string> _immersivePrompts = new Dictionary<string, string>
        {
            ["title"] = "PPT标题翻译要求：简洁有力、术语准确、符合目标语言表达习惯。确保翻译后的标题在幻灯片上清晰易读。",
            ["content"] = "PPT内容翻译要求：准确传达原文含义、保持专业术语一致性、译文长度适中、符合目标语言表达习惯。注意保持原文的格式和段落结构。",
            ["bullet"] = "PPT项目符号翻译要求：保持列表格式、术语一致、简洁明了。确保每个项目符号的翻译风格统一。",
            ["chart"] = "PPT图表标签翻译要求：术语准确、简洁明了、保持图表的可读性。确保标签翻译后不会影响图表的理解。",
            ["note"] = "PPT备注翻译要求：完整传达原文含义、保持专业性、语言流畅自然。备注通常包含演讲者的补充说明，需要准确翻译。",
            ["smartart"] = "PPT SmartArt文本翻译要求：术语准确、简洁明了、保持图形结构的清晰性。SmartArt通常用于展示概念或流程，翻译需要准确传达这些概念。"
        };

        public PPTProcessor(TranslationService translationService, ILogger<PPTProcessor> logger)
        {
            _translationService = translationService ?? throw new ArgumentNullException(nameof(translationService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _termExtractor = new TermExtractor();
            _progressCallback = (p, m) => { }; // 初始化为 no-op，避免可空警告
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
            string sourceLang = "zh", string targetLang = "en", string outputFormat = "bilingual",
            bool preserveFormatting = true, bool translateNotes = true, bool translateCharts = true,
            bool translateSmartArt = true, bool autoAdjustLayout = true, bool useImmersiveStyle = false)
        {
            _useTerminology = useTerminology;
            _preprocessTerms = preprocessTerms;
            _sourceLang = sourceLang;
            _targetLang = targetLang;
            _outputFormat = outputFormat;
            _preserveFormatting = preserveFormatting;
            _translateNotes = translateNotes;
            _translateCharts = translateCharts;
            _translateSmartArt = translateSmartArt;
            _autoAdjustLayout = autoAdjustLayout;
            _useImmersiveStyle = useImmersiveStyle;
            _isCnToForeign = sourceLang == "zh" || (sourceLang == "auto" && targetLang != "zh");

            _logger.LogInformation($"PPT翻译选项已设置: 术语库={useTerminology}, 预处理={preprocessTerms}, 输出格式={outputFormat}, 沉浸式风格={useImmersiveStyle}");
            _logger.LogInformation($"翻译方向: {(_isCnToForeign ? "中文→外语" : "外语→中文")}");
            _logger.LogInformation($"输出格式: {(_outputFormat == "bilingual" ? "双语对照" : "仅翻译结果")}");

            if (_useImmersiveStyle)
            {
                _logger.LogInformation("🌟 沉浸式翻译风格已启用 - 将使用沉浸式翻译提示词");
            }

            // 验证术语预处理设置
            if (_useTerminology && _preprocessTerms)
            {
                _logger.LogInformation("✅ 术语预处理已启用 - 将在翻译前直接替换术语");
            }
            else if (_useTerminology && !_preprocessTerms)
            {
                _logger.LogInformation("⚡ 术语库已启用但预处理关闭 - 将使用占位符策略");
            }
            else
            {
                _logger.LogInformation("❌ 术语库功能已关闭");
            }
        }

        /// <summary>
        /// 从双语PPT文件生成纯翻译版本（删除中文内容）
        /// </summary>
        public string GenerateTranslationOnlyFromBilingual(string bilingualFilePath)
        {
            try
            {
                _logger.LogInformation($"开始从双语文件生成纯翻译版本: {bilingualFilePath}");

                var fileInfo = new FileInfo(bilingualFilePath);
                var outputDir = fileInfo.DirectoryName;
                var fileName = Path.GetFileNameWithoutExtension(fileInfo?.Name ?? string.Empty);
                var ext = fileInfo?.Extension ?? ".pptx";
                var timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                // 生成纯翻译版本的文件名
                var baseDir = string.IsNullOrEmpty(outputDir) ? AppDomain.CurrentDomain.BaseDirectory : outputDir;
                var translationOnlyPath = Path.Combine(baseDir!, $"{fileName}_纯翻译版_{timeStamp}{ext}");

                // 复制双语文件到新位置
                File.Copy(bilingualFilePath, translationOnlyPath, true);

                // 打开文件并删除中文内容
                using (var presentation = PresentationDocument.Open(translationOnlyPath, true))
                {
                    var presentationPart = presentation.PresentationPart;
                    if (presentationPart == null)
                    {
                        throw new InvalidOperationException("无法打开PPT文档");
                    }

                    var slides = presentationPart.SlideParts.ToList();

                    // 处理每张幻灯片
                    for (int i = 0; i < slides.Count; i++)
                    {
                        var slide = slides[i];
                        RemoveChineseTextFromSlide(slide);
                    }

                    // 处理备注中的中文
                    var notesParts = slides.Select(s => s.NotesSlidePart).Where(n => n != null);
                    foreach (var notesPart in notesParts)
                    {
                        if (notesPart != null) RemoveChineseTextFromNotes(notesPart);
                    }

                    presentation.Save();
                }

                _logger.LogInformation($"纯翻译版本已生成: {translationOnlyPath}");
                return translationOnlyPath;
            }
            catch (Exception ex)
            {
                _logger.LogError($"生成纯翻译版本失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 清理文档语言，保留指定语言的内容
        /// </summary>
        public async Task<string> CleanLanguageAsync(string filePath, string keepLanguage)
        {
            _logger.LogInformation($"开始清理语言，文件: {filePath}, 保留语言: {keepLanguage}");
            
            var startTime = DateTime.Now;
            var processedSlides = 0;
            var stats = new Dictionary<string, int>
            {
                ["processedTextShapes"] = 0,
                ["processedTables"] = 0,
                ["skippedSmartArt"] = 0,
                ["skippedCharts"] = 0,
                ["skippedMedia"] = 0,
                ["processedNotes"] = 0
            };
            
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

                // 打开PPT文档
                using var presentation = PresentationDocument.Open(outputFilePath, true);
                var presentationPart = presentation.PresentationPart;
                if (presentationPart == null)
                {
                    throw new InvalidOperationException("无法打开PPT文档");
                }

                // 获取所有幻灯片
                var slides = presentationPart.SlideParts.ToList();
                var totalSlides = slides.Count;
                _logger.LogInformation($"开始处理 {totalSlides} 张幻灯片");

                // 处理幻灯片
                for (int i = 0; i < slides.Count; i++)
                {
                    var slide = slides[i];
                    double progress = (double)(i + 1) / totalSlides * 100;
                    _progressCallback?.Invoke(progress, $"处理幻灯片 {i + 1}/{totalSlides}");

                    await ProcessSlideForCleanupAsync(slide, keepLanguage, stats);
                    processedSlides++;
                }

                // 处理备注
                if (_translateNotes)
                {
                    double progress = 95.0;
                    _progressCallback?.Invoke(progress, "正在处理幻灯片备注...");
                    await ProcessNotesForCleanupAsync(presentationPart, keepLanguage, stats);
                }

                // 保存文档
                presentation.Save();
                
                var endTime = DateTime.Now;
                var duration = endTime - startTime;
                
                // 记录统计信息
                _logger.LogInformation($"清理完成，输出文件: {outputFilePath}");
                _logger.LogInformation($"处理统计：");
                _logger.LogInformation($"  - 处理幻灯片: {processedSlides}/{totalSlides}");
                _logger.LogInformation($"  - 处理文本框: {stats["processedTextShapes"]}");
                _logger.LogInformation($"  - 处理表格: {stats["processedTables"]}");
                _logger.LogInformation($"  - 跳过SmartArt: {stats["skippedSmartArt"]}");
                _logger.LogInformation($"  - 跳过图表: {stats["skippedCharts"]}");
                _logger.LogInformation($"  - 跳过多媒体: {stats["skippedMedia"]}");
                _logger.LogInformation($"  - 处理备注: {stats["processedNotes"]}");
                _logger.LogInformation($"  - 处理耗时: {duration.TotalSeconds:F2}秒");
                
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
        /// 处理幻灯片进行语言清理
        /// </summary>
        private async Task ProcessSlideForCleanupAsync(SlidePart slidePart, string keepLanguage, Dictionary<string, int> stats)
        {
            try
            {
                var slide = slidePart?.Slide;
                if (slide == null) return;

                // 递归处理幻灯片中的所有元素
                if (slide.CommonSlideData?.ShapeTree != null)
                {
                    await ProcessCompositeElementForCleanupAsync(slide.CommonSlideData.ShapeTree, keepLanguage, stats);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理幻灯片时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 递归处理组合元素进行语言清理
        /// </summary>
        private async Task ProcessCompositeElementForCleanupAsync(OpenXmlElement composite, string keepLanguage, Dictionary<string, int> stats)
        {
            foreach (var element in composite.Elements<OpenXmlElement>())
            {
                if (element is Shape shape)
                {
                    // 处理文本框
                    if (shape.TextBody != null)
                    {
                        await ProcessTextShapeForCleanupAsync(shape, keepLanguage);
                        stats["processedTextShapes"]++;
                    }

                    // 处理多媒体内容
                    if (shape.Descendants().Any(e => e.LocalName == "video" || e.LocalName == "audio"))
                    {
                        _logger.LogDebug("跳过多媒体内容（保留原样）");
                        stats["skippedMedia"]++;
                    }
                }
                else if (element is GraphicFrame graphicFrame)
                {
                    // 处理表格
                    if (graphicFrame.Descendants().Any(e => e.LocalName == "tbl"))
                    {
                        await ProcessTableForCleanupAsync(graphicFrame, keepLanguage);
                        stats["processedTables"]++;
                    }
                    // 处理图表
                    else if (graphicFrame.Descendants().Any(e => e.LocalName == "chart"))
                    {
                        _logger.LogDebug("跳过图表（保留原样）");
                        stats["skippedCharts"]++;
                    }
                }
                else if (element is GroupShape groupShape)
                {
                    await ProcessCompositeElementForCleanupAsync(groupShape, keepLanguage, stats);
                }
            }
        }

        /// <summary>
        /// 处理文本形状进行语言清理
        /// </summary>
        private async Task ProcessTextShapeForCleanupAsync(Shape shape, string keepLanguage)
        {
            try
            {
                var textBody = shape.TextBody;
                if (textBody == null) return;

                var paragraphs = textBody.Elements<A.Paragraph>().ToList();
                foreach (var paragraph in paragraphs)
                {
                    var paragraphText = GetParagraphText(paragraph).Trim();
                    if (string.IsNullOrEmpty(paragraphText)) continue;

                    // 处理文本
                    await ProcessPlainTextForCleanupAsync(paragraph, paragraphText, keepLanguage);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理文本形状时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取段落文本（包含换行符）
        /// </summary>
        private string GetParagraphText(A.Paragraph paragraph)
        {
            var sb = new StringBuilder();
            foreach (var child in paragraph.ChildElements)
            {
                if (child is A.Run run)
                {
                    var text = run.GetFirstChild<A.Text>();
                    if (text != null) sb.Append(text.Text);
                }
                else if (child is A.Break)
                {
                    sb.Append("\n");
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// 处理表格进行语言清理
        /// </summary>
        private async Task ProcessTableForCleanupAsync(OpenXmlElement element, string keepLanguage)
        {
            try
            {
                // 查找表格元素
                var tables = element.Descendants().Where(e => e.LocalName == "tbl");
                foreach (var table in tables)
                {
                    // 查找行元素
                    var rows = table.Elements().Where(e => e.LocalName == "tr");
                    foreach (var row in rows)
                    {
                        // 查找单元格元素
                        var cells = row.Elements().Where(e => e.LocalName == "tc");
                        foreach (var cell in cells)
                        {
                            // 查找文本体
                            var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                            if (textBody == null) continue;

                            // 查找段落
                            var paragraphs = textBody.Elements().Where(e => e.LocalName == "p");
                            foreach (var pObj in paragraphs)
                            {
                                if (pObj is A.Paragraph paragraph)
                                {
                                    var paragraphText = GetParagraphText(paragraph).Trim();
                                    if (string.IsNullOrEmpty(paragraphText)) continue;

                                    // 处理文本
                                    await ProcessPlainTextForCleanupAsync(paragraph, paragraphText, keepLanguage);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理幻灯片备注进行语言清理
        /// </summary>
        private async Task ProcessNotesForCleanupAsync(PresentationPart presentationPart, string keepLanguage, Dictionary<string, int> stats)
        {
            try
            {
                var notesMasterPart = presentationPart.NotesMasterPart;
                if (notesMasterPart == null) return;

                // 获取所有备注页
                var notesSlides = presentationPart.GetPartsOfType<NotesSlidePart>();
                foreach (var notesSlide in notesSlides)
                {
                    var notesSlide1 = notesSlide.NotesSlide;
                    var shapes = notesSlide1.Descendants<Shape>();

                    foreach (var shape in shapes)
                    {
                        if (shape.TextBody != null)
                        {
                            await ProcessTextShapeForCleanupAsync(shape, keepLanguage);
                            stats["processedNotes"]++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理幻灯片备注时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理普通文本进行语言清理
        /// </summary>
        private async Task ProcessPlainTextForCleanupAsync(A.Paragraph paragraph, string text, string keepLanguage)
        {
            // 使用智能清理（模型提取）
            var cleanedText = await _translationService.SmartCleanTextAsync(text, keepLanguage);

            if (cleanedText == text) return; // 无变化

            // 获取Run属性用于重建
            var runs = paragraph.Elements<A.Run>().ToList();
            var firstRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties;
            var lastRunProps = runs.LastOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties;

            // 清除旧内容
            paragraph.RemoveAllChildren<A.Run>();
            paragraph.RemoveAllChildren<A.Break>();

            if (string.IsNullOrWhiteSpace(cleanedText)) return; // 被删除

            // 智能应用格式
            // 如果保留后的文本出现在原文末尾，尝试应用最后一个Run的格式
            // 否则应用第一个Run的格式
            var newRunProps = (text.EndsWith(cleanedText) && lastRunProps != null) ? lastRunProps : firstRunProps;
            if (newRunProps == null) newRunProps = new A.RunProperties();

            var newRun = new A.Run { RunProperties = newRunProps };
            newRun.AppendChild(new A.Text { Text = cleanedText });
            paragraph.AppendChild(newRun);
        }



        /// <summary>
        /// 处理PPT文档翻译
        /// </summary>
        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguage,
            Dictionary<string, string> terminology)
        {
            // 直接调用批量处理方法，该方法已实现了并行翻译和优化的处理流程
            return await BatchProcessPPTDocumentAsync(filePath, targetLanguage, terminology, createCopy: true);
        }

        /// <summary>
        /// 处理单张幻灯片
        /// </summary>
        private async Task ProcessSlideAsync(SlidePart slidePart, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber, PresentationPart presentationPart, bool useBilingualMode = true)
        {
            try
            {
                var slide = slidePart?.Slide;
                if (slide == null) return;

                // 递归处理幻灯片中的所有元素
                if (slide.CommonSlideData?.ShapeTree != null)
                {
                    await ProcessCompositeElementAsync(slide.CommonSlideData.ShapeTree, terminology, translationResults, slideNumber, presentationPart, useBilingualMode);
                }
                if (_translateSmartArt && slidePart != null)
                {
                    await ProcessSmartArtPartsAsync(slidePart, terminology, translationResults, slideNumber, useBilingualMode);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理第 {slideNumber} 张幻灯片时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 递归处理组合元素（如幻灯片根节点或组合形状）
        /// </summary>
        private async Task ProcessCompositeElementAsync(OpenXmlElement composite, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber, PresentationPart presentationPart, bool useBilingualMode)
        {
            foreach (var element in composite.Elements<OpenXmlElement>())
            {
                if (element is Shape shape)
                {
                    // 处理普通文本框
                    if (shape.TextBody != null)
                    {
                        await ProcessTextShapeAsync(shape, terminology, translationResults, slideNumber, "文本", useBilingualMode);
                    }
                }
                else if (element is GraphicFrame graphicFrame)
                {
                    // 处理表格
                    if (graphicFrame.Descendants().Any(e => e.LocalName == "tbl"))
                    {
                        await ProcessTableAsync(graphicFrame, terminology, translationResults, slideNumber);
                    }
                    // 处理图表
                    else if (_translateCharts && graphicFrame.Descendants().Any(e => e.LocalName == "chart"))
                    {
                        await ProcessChartAsync(graphicFrame, terminology, translationResults, slideNumber, presentationPart);
                    }
                }
                else if (element is GroupShape groupShape)
                {
                    // 处理组合形状（包括可能被误认为是SmartArt的普通组合）
                    // 直接递归处理子元素，确保所有包含文本的形状都被处理
                    await ProcessCompositeElementAsync(groupShape, terminology, translationResults, slideNumber, presentationPart, useBilingualMode);
                }
            }
        }

        private async Task ProcessSmartArtPartsAsync(SlidePart slidePart, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber, bool useBilingualMode)
        {
            try
            {
                var diagramParts = slidePart.DiagramDataParts.ToList();
                if (diagramParts.Count == 0) return;

                foreach (var diagramPart in diagramParts)
                {
                    var textElements = GetSmartArtTextElements(diagramPart);
                    if (textElements.Count == 0) continue;

                    foreach (var textElement in textElements)
                    {
                        var originalText = textElement.Text?.Trim();
                        if (string.IsNullOrEmpty(originalText)) continue;

                        var (text, formulas) = ExtractLatexFormulas(originalText);
                        var translationPrompt = _useImmersiveStyle
                            ? _immersivePrompts.GetValueOrDefault("smartart", _immersivePrompts["content"])
                            : _pptPrompts.GetValueOrDefault("smartart", _pptPrompts["content"]);

                        var translatedText = await TranslateTextWithRetryAsync(text, terminology, translationPrompt, originalText);
                        if (string.IsNullOrEmpty(translatedText)) continue;

                        if (formulas.Any())
                        {
                            translatedText = RestoreLatexFormulas(translatedText, formulas);
                        }

                        translationResults.Add(new TranslationResult
                        {
                            OriginalText = originalText,
                            TranslatedText = translatedText,
                            Position = $"幻灯片 {slideNumber}，SmartArt",
                            ElementType = "smartart"
                        });

                        textElement.Text = useBilingualMode
                            ? $"{originalText}##-----##{translatedText}"
                            : translatedText;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理SmartArt时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理文本形状
        /// </summary>
        private async Task ProcessTextShapeAsync(Shape shape, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber, string elementType, bool useBilingualMode = true)
        {
            try
            {
                var textBody = shape.TextBody;
                if (textBody == null) return;

                var paragraphs = textBody.Elements<A.Paragraph>().ToList();
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var paragraph = paragraphs[i];
                    var runs = paragraph.Elements<A.Run>().ToList();

                    // 聚合段落文本
                    var sb = new StringBuilder();
                    foreach (var r in runs)
                    {
                        var t = r.GetFirstChild<A.Text>();
                        if (t != null && !string.IsNullOrEmpty(t.Text)) sb.Append(t.Text);
                    }

                    var originalText = sb.ToString().Trim();
                    if (string.IsNullOrEmpty(originalText)) continue;

                    // 检查文本中是否包含数学公式
                    var (text, formulas) = ExtractLatexFormulas(originalText);

                    // 确定文本类型以选择合适的提示词
                    var textType = DetermineTextType(shape, originalText);
                    var translationPrompt = _useImmersiveStyle 
                        ? _immersivePrompts.GetValueOrDefault(textType, _immersivePrompts["content"])
                        : _pptPrompts.GetValueOrDefault(textType, _pptPrompts["content"]);

                    // 翻译文本
                    var translatedText = await TranslateTextWithRetryAsync(text, terminology, translationPrompt, originalText);
                    if (string.IsNullOrEmpty(translatedText)) continue;

                    // 将公式重新插入到翻译后的文本中
                    if (formulas.Any())
                    {
                        translatedText = RestoreLatexFormulas(translatedText, formulas);
                    }

                    // 记录翻译结果
                    translationResults.Add(new TranslationResult
                    {
                        OriginalText = originalText,
                        TranslatedText = translatedText,
                        Position = $"幻灯片 {slideNumber}，{elementType}",
                        ElementType = textType
                    });

                    if (useBilingualMode)
                    {
                        // 双语：将原文和译文合并到同一个段落中，用##-----##分隔
                        var baseRunProps = (runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties) ?? new A.RunProperties();

                        // 清空所有run
                        foreach (var r in runs) r.Remove();

                        // 添加原文
                        var originalRun = new A.Run { RunProperties = baseRunProps };
                        originalRun.AppendChild(new A.Text { Text = originalText });
                        paragraph.AppendChild(originalRun);

                        // 添加分隔符##-----##
                        var separatorRun = new A.Run { RunProperties = baseRunProps };
                        separatorRun.AppendChild(new A.Text { Text = "##-----##" });
                        paragraph.AppendChild(separatorRun);

                        // 添加译文
                        var translatedRun = new A.Run { RunProperties = baseRunProps };
                        translatedRun.AppendChild(new A.Text { Text = translatedText });
                        paragraph.AppendChild(translatedRun);

                        _logger.LogDebug($"双语模式：原文和译文已合并到同一段落，使用##-----##分隔");
                    }
                    else
                    {
                        // 仅译文：替换该段落中的文本，但尽量保留第一个run的样式，并用 <a:br/> 处理换行
                        var baseRunProps = (runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties) ?? new A.RunProperties();

                        foreach (var r in runs) r.Remove();

                        // 使用统一的换行写入方法，避免 PowerPoint 修复
                        AppendTextWithLineBreaks(paragraph, translatedText, baseRunProps);
                        EnsureEndParaRPr(paragraph);
                        ValidateAndFixTextShape(shape);
                    }
                }

                // 自动调整布局
                if (_autoAdjustLayout)
                {
                    AdjustTextShapeLayout(shape);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理文本形状时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理表格
        /// </summary>
        private async Task ProcessTableAsync(OpenXmlElement element, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber)
        {
            try
            {
                // 查找表格元素
                var tables = element.Descendants().Where(e => e.LocalName == "tbl");
                foreach (var table in tables)
                {
                    // 查找行元素
                    var rows = table.Elements().Where(e => e.LocalName == "tr");
                    foreach (var row in rows)
                    {
                        // 查找单元格元素
                        var cells = row.Elements().Where(e => e.LocalName == "tc");
                        foreach (var cell in cells)
                        {
                            // 查找文本体
                            var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                            if (textBody == null) continue;

                            // 查找段落
                            var paragraphs = textBody.Elements().Where(e => e.LocalName == "p");
                            foreach (var paragraph in paragraphs)
                            {
                                // 查找文本运行
                                var runs = paragraph.Elements().Where(e => e.LocalName == "r");
                                foreach (var run in runs)
                                {
                                    // 查找文本元素
                                    var textElement = run.Elements().FirstOrDefault(e => e.LocalName == "t");
                                    if (textElement == null) continue;

                                    var originalText = textElement.InnerText?.Trim();
                                    if (string.IsNullOrEmpty(originalText)) continue;

                                    // 检查文本中是否包含数学公式
                                    var (text, formulas) = ExtractLatexFormulas(originalText);

                                    // 翻译文本
                                    var translationPrompt = _useImmersiveStyle ? _immersivePrompts["content"] : _pptPrompts["content"];
                                    var translatedText = await TranslateTextWithRetryAsync(text, terminology, translationPrompt, originalText);
                                    if (string.IsNullOrEmpty(translatedText)) continue;

                                    // 将公式重新插入到翻译后的文本中
                                    if (formulas.Any())
                                    {
                                        translatedText = RestoreLatexFormulas(translatedText, formulas);
                                    }

                                    // 记录翻译结果
                                    translationResults.Add(new TranslationResult
                                    {
                                        OriginalText = originalText,
                                        TranslatedText = translatedText,
                                        Position = $"幻灯片 {slideNumber}，表格单元格",
                                        ElementType = "table"
                                    });

                                    // 根据输出格式设置文本内容
                                    if (_outputFormat == "bilingual")
                                    {
                                        // 表格内用 Drawing 的 A.Text，不要混用 Presentation.Text
                                        var aText = textElement as A.Text ?? textElement.GetFirstChild<A.Text>();
                                        if (aText != null)
                                        {
                                            aText.Text = $"{originalText}##-----##{translatedText}";
                                        }
                                    }
                                    else
                                    {
                                        var aText = textElement as A.Text ?? textElement.GetFirstChild<A.Text>();
                                        if (aText != null)
                                        {
                                            aText.Text = translatedText;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理表格时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理图表
        /// </summary>
        private Task ProcessChartAsync(OpenXmlElement element, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber, PresentationPart presentationPart)
        {
            try
            {
                var graphicFrames = element is GraphicFrame gf 
                    ? new[] { gf } 
                    : element.Descendants<GraphicFrame>();

                int chartIndex = 0;
                
                foreach (var graphicFrame in graphicFrames)
                {
                    try
                    {
                        var frameId = "未知";
                        try
                        {
                            var idAttribute = graphicFrame.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            if (idAttribute.Value != null)
                            {
                                frameId = idAttribute.Value;
                            }
                        }
                        catch
                        {
                            frameId = "未知";
                        }

                        // 尝试获取图表部分
                        ChartPart? chartPart = null;
                        try
                        {
                            var chartReference = graphicFrame.GetFirstChild<ChartReference>();
                            if (chartReference != null && presentationPart != null)
                            {
                                var relationshipId = chartReference.Id;
                                if (!string.IsNullOrEmpty(relationshipId))
                                {
                                    var part = presentationPart.GetPartById(relationshipId!);
                                    if (part != null)
                                    {
                                        chartPart = part as ChartPart;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            chartPart = null;
                        }
                        
                        if (chartPart != null)
                        {
                            _logger.LogInformation($"发现图表元素，图表ID: {frameId}");
                            
                            // 提取图表文本
                            var chartTextEntries = ExtractChartText(chartPart, slideNumber, frameId, chartIndex);
                            
                            // 翻译图表文本
                            if (chartTextEntries.Any())
                            {
                                var translatedEntries = TranslateTextEntries(chartTextEntries, terminology, slideNumber, "chart");
                                
                                // 将翻译结果写入图表
                                WriteChartText(chartPart, translatedEntries, slideNumber);
                                
                                _logger.LogInformation($"已处理图表文本: {translatedEntries.Count} 个条目");
                            }
                        }
                        
                        chartIndex++;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"处理单个图表时出错: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理图表时出错: {ex.Message}");
            }

            return Task.CompletedTask;
        }

        /// <summary>
        /// 处理SmartArt
        /// </summary>
        private async Task ProcessSmartArtAsync(OpenXmlElement element, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, int slideNumber)
        {
            try
            {
                // 获取所有包含文本的子形状
                var shapes = element.Descendants<Shape>();
                foreach (var subShape in shapes)
                {
                    if (subShape.TextBody != null)
                    {
                        await ProcessTextShapeAsync(subShape, terminology, translationResults, slideNumber, "SmartArt");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理SmartArt时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理幻灯片备注
        /// </summary>
        private Task ProcessNotesAsync(PresentationPart presentationPart, Dictionary<string, string> terminology,
            List<TranslationResult> translationResults, SlidePart? specificSlidePart = null, bool useBilingualMode = true)
        {
            try
            {
                var notesMasterPart = presentationPart.NotesMasterPart;
                if (notesMasterPart == null) return Task.CompletedTask;

                // 获取所有备注页或特定幻灯片的备注页
                IEnumerable<NotesSlidePart> notesSlides;
                if (specificSlidePart != null)
                {
                    // 只处理特定幻灯片的备注
                    var notesPart = specificSlidePart.NotesSlidePart;
                    notesSlides = notesPart != null ? new[] { notesPart } : Enumerable.Empty<NotesSlidePart>();
                }
                else
                {
                    // 处理所有备注页
                    notesSlides = presentationPart.GetPartsOfType<NotesSlidePart>();
                }

                foreach (var notesSlide in notesSlides)
                {
                    var notesSlide1 = notesSlide.NotesSlide;
                    var shapes = notesSlide1.Descendants<Shape>();

                    foreach (var shape in shapes)
                    {
                        if (shape.TextBody != null)
                        {
                            // 同步调用异步方法
                            ProcessTextShapeAsync(shape, terminology, translationResults, 0, "备注", useBilingualMode).GetAwaiter().GetResult();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"处理幻灯片备注时出错: {ex.Message}");
            }

            // 由于此方法目前是同步的，直接返回已完成的任务
            return Task.CompletedTask;
        }

        /// <summary>
        /// 从文本中提取LaTeX公式
        /// </summary>
        private (string text, List<(string placeholder, string formula)> formulas) ExtractLatexFormulas(string text)
        {
            var formulas = new List<(string placeholder, string formula)>();
            var processedText = text;

            for (int i = 0; i < _latexPatterns.Count; i++)
            {
                var pattern = _latexPatterns[i];
                var matches = pattern.Matches(processedText);

                for (int j = 0; j < matches.Count; j++)
                {
                    var match = matches[j];
                    var formula = match.Value;
                    var placeholder = $"[FORMULA_{i}_{j}]";
                    processedText = processedText.Replace(formula, placeholder);
                    formulas.Add((placeholder, formula));
                }
            }

            return (processedText, formulas);
        }

        /// <summary>
        /// 将公式占位符替换回原始公式
        /// </summary>
        private string RestoreLatexFormulas(string text, List<(string placeholder, string formula)> formulas)
        {
            var result = text;
            foreach (var (placeholder, formula) in formulas)
            {
                result = result.Replace(placeholder, formula);
            }
            return result;
        }

        /// <summary>
        /// 确定文本类型
        /// </summary>
        private string DetermineTextType(Shape shape, string text)
        {
            // 检查占位符类型
            var placeholder = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape;
            if (placeholder != null)
            {
                if (placeholder.Type != null)
                {
                    if (placeholder.Type == P.PlaceholderValues.Title || placeholder.Type == P.PlaceholderValues.CenteredTitle)
                    {
                        return "title";
                    }
                    if (placeholder.Type == P.PlaceholderValues.Body || placeholder.Type == P.PlaceholderValues.Object)
                    {
                        return "content";
                    }
                }
            }

            // 检查是否是标题
            var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
            if (!string.IsNullOrEmpty(name) && (name.Contains("Title") || name.Contains("标题")))
            {
                return "title";
            }

            // 检查是否是项目符号列表
            if (text.Contains('\n') && text.Split('\n').Any(line => line.Trim().StartsWith("•") || line.Trim().StartsWith("-")))
            {
                return "bullet";
            }

            // 默认返回内容类型
            return "content";
        }

        /// <summary>
        /// 自动调整文本形状布局
        /// </summary>
        private void AdjustTextShapeLayout(Shape shape)
        {
            try
            {
                if (shape?.TextBody == null) return;

                var textBody = shape.TextBody;

                // 设置文本框属性
                var bodyProps = textBody.BodyProperties ??= new A.BodyProperties();

                // 设置垂直对齐为顶部
                bodyProps.Anchor = A.TextAnchoringTypeValues.Top;

                // 设置文本自动适应
                var noAutoFit = bodyProps.GetFirstChild<A.NoAutoFit>();
                if (noAutoFit != null) noAutoFit.Remove();

                // 对于双语模式，使用更宽松的自动适应
                if (_outputFormat == "bilingual")
                {
                    // 双语模式：允许文本框扩展
                    if (bodyProps.GetFirstChild<A.ShapeAutoFit>() == null)
                    {
                        bodyProps.AppendChild(new A.ShapeAutoFit());
                    }
                }
                else
                {
                    // 单语模式：使用正常自动适应
                    if (bodyProps.GetFirstChild<A.NormalAutoFit>() == null)
                    {
                        bodyProps.AppendChild(new A.NormalAutoFit());
                    }
                }

                _logger.LogDebug($"已应用自动布局调整，形状ID: {shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"调整文本形状布局失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查文本是否应该跳过翻译（纯序号或纯数字）
        /// </summary>
        private bool ShouldSkipTranslation(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return true;

            text = text.Trim();

            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[\d\.\,\-\+%\s]+$"))
            {
                return true;
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[\d①②③④⑤⑥⑦⑧⑨⑩ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ][\.\)、]*$") ||
                System.Text.RegularExpressions.Regex.IsMatch(text, @"^[a-zA-Z][\.\)]*$"))
            {
                return true;
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[\(\[\{][\d①②③④⑤⑥⑦⑧⑨⑩ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩA-zA-Z]+[\)\]\}]$"))
            {
                return true;
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[≥≤<>＜＞\-\+\=\(\)\[\]\{\}\s、。，；：！？""''…—·]+$"))
            {
                return true;
            }

            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[≥≤<>＜＞]*[\d\.\,\-\+]+[μΩa-zA-Z]{1,5}$") &&
                !System.Text.RegularExpressions.Regex.IsMatch(text, @"[\u4e00-\u9fff]"))
            {
                return true;
            }

            if (text.Length <= 1)
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// 过滤翻译结果中的提示词和无关内容
        /// </summary>
        private string FilterTranslationResult(string translatedText)
        {
            if (string.IsNullOrWhiteSpace(translatedText))
                return translatedText;

            // 过滤常见的提示词
            var filtersToRemove = new[]
            {
                "请提供纯中文翻译：",
                "请提供纯中文翻译:",
                "纯中文翻译：",
                "纯中文翻译:",
                "中文翻译：",
                "中文翻译:",
                "翻译：",
                "翻译:",
                "Translation:",
                "Translated:",
                "中文：",
                "中文:",
                "Chinese:",
                "以下是翻译：",
                "以下是翻译:",
                "翻译结果：",
                "翻译结果:",
                "Translation result:",
                "Result:",
                "答案：",
                "答案:",
                "Answer:"
            };

            string result = translatedText;

            // 移除开头的提示词
            foreach (var filter in filtersToRemove)
            {
                if (result.StartsWith(filter, StringComparison.OrdinalIgnoreCase))
                {
                    result = result.Substring(filter.Length).Trim();
                    _logger.LogDebug($"已过滤翻译结果开头的提示词: '{filter}'");
                }
            }

            // 移除多余的换行和空白
            result = result.Trim();

            // 如果结果为空，返回原文
            if (string.IsNullOrWhiteSpace(result))
            {
                _logger.LogWarning("过滤后翻译结果为空，返回原始翻译结果");
                return translatedText.Trim();
            }

            return result;
        }

        /// <summary>
        /// 将同一段落中被空白分行分隔的内容拆分为独立分段
        /// 规则：以两个及以上连续换行（或空白行）作为段落边界；单个换行保留在分段内部
        /// </summary>
        private List<string> SplitSegmentsByBlankLines(string text)
        {
            var segments = new List<string>();
            if (string.IsNullOrEmpty(text)) return segments;

            // 统一换行符
            var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            var lines = normalized.Split('\n');

            var sb = new StringBuilder();
            int blankStreak = 0;

            foreach (var raw in lines)
            {
                var line = raw; // 保留行内内容（不TrimStart，避免破坏缩进）
                bool isBlank = string.IsNullOrWhiteSpace(line);

                if (isBlank)
                {
                    blankStreak++;
                    if (blankStreak >= 1)
                    {
                        // 遇到空白行：结束当前段，并开始新段（跳过连续空白）
                        if (sb.Length > 0)
                        {
                            segments.Add(sb.ToString().Trim());
                            sb.Clear();
                        }
                        continue; // 连续空白行全部跳过
                    }
                }
                else
                {
                    // 非空白行
                    if (sb.Length > 0)
                        sb.Append('\n'); // 段内使用单换行
                    sb.Append(line.TrimEnd());
                    blankStreak = 0;
                }
            }

            if (sb.Length > 0)
                segments.Add(sb.ToString().Trim());

            // 如果没有分出任何段且原文不为空，则以原文为一个段
            if (segments.Count == 0 && !string.IsNullOrWhiteSpace(text))
                segments.Add(text.Trim());

            return segments;
        }

        /// <summary>
        /// 带重试的文本翻译
        /// </summary>
        private async Task<string> TranslateTextWithRetryAsync(string text, Dictionary<string, string> terminology, string prompt, string originalText = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            // 前置过滤：检查是否应该跳过翻译
            if (ShouldSkipTranslation(text))
            {
                return text;
            }

            // 规范术语映射方向：确保键总是"源语言术语"，值为"目标语言术语"
            // - 中文→外语：术语库本身是 {中文: 外语}，可直接使用
            // - 外语→中文：MainWindow已经直接加载了外译中术语库，无需反转
            Dictionary<string, string> termsToUse = terminology;

            // 术语预替换：在翻译前对文本进行术语预处理（直接替换为目标术语）
            string toTranslate = text;
            if (_useTerminology && _preprocessTerms && termsToUse != null && termsToUse.Any())
            {
                var before = text;
                _logger.LogInformation($"🔄 开始术语预替换，术语库大小: {termsToUse.Count}");
                _logger.LogDebug($"术语库内容样本: {string.Join(", ", termsToUse.Take(3).Select(kv => $"{kv.Key}→{kv.Value}"))}");

                toTranslate = _termExtractor.ReplaceTermsInText(text, termsToUse, _isCnToForeign);

                var beforeSnippet = before.Substring(0, Math.Min(80, before.Length));
                var afterSnippet = toTranslate.Substring(0, Math.Min(80, toTranslate.Length));

                if (before != toTranslate)
                {
                    _logger.LogInformation($"✅ PPT术语预替换成功（{(_isCnToForeign ? "中文→外语" : "外语→中文")}）：'{beforeSnippet}' -> '{afterSnippet}'");
                }
                else
                {
                    _logger.LogInformation($"⚠️ PPT术语预替换未发现匹配项：'{beforeSnippet}'");
                }
            }
            else if (_useTerminology && !_preprocessTerms)
            {
                _logger.LogInformation("⚡ 术语预处理已关闭，将使用占位符策略");
            }

            // 若已预替换，则不再向翻译器传入术语表，避免二次干预；否则保留术语表供翻译器做占位符策略
            var terminologyForTranslator = (_useTerminology && !_preprocessTerms) ? termsToUse : null;

            for (int attempt = 1; attempt <= _retryCount; attempt++)
            {
                try
                {
                    var translatedText = await _translationService.TranslateTextAsync(toTranslate, terminologyForTranslator, _sourceLang, _targetLang, prompt, originalText);
                    if (!string.IsNullOrEmpty(translatedText))
                    {
                        // 过滤翻译结果中的提示词
                        translatedText = FilterTranslationResult(translatedText);
                        return translatedText;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"第 {attempt} 次翻译尝试失败: {ex.Message}");
                    if (attempt < _retryCount)
                    {
                        await Task.Delay(_retryDelay * attempt);
                    }
                }
            }

            _logger.LogError($"文本翻译失败，已重试 {_retryCount} 次: {text}");
            
            // 如果提供了原文，在翻译失败时返回原文，避免PPT显示默认占位符
            if (!string.IsNullOrEmpty(originalText))
            {
                return originalText;
            }

            return string.Empty;
        }

        /// <summary>
        /// 导出翻译结果到Excel
        /// </summary>
        private Task ExportTranslationResultsAsync(List<TranslationResult> translationResults, string filePath, string targetLanguage)
        {
            try
            {
                var outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输出");
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                var excelPath = Path.Combine(outputDir, $"{fileName}_翻译结果_{timeStamp}.xlsx");

                // 使用ClosedXML创建Excel文件
                using (var workbook = new ClosedXML.Excel.XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("翻译结果");

                    // 添加表头
                    worksheet.Cell("A1").Value = "原文";
                    worksheet.Cell("B1").Value = "译文";
                    worksheet.Cell("C1").Value = "位置";
                    worksheet.Cell("D1").Value = "类型";

                    // 添加数据
                    for (int i = 0; i < translationResults.Count; i++)
                    {
                        var result = translationResults[i];
                        var row = i + 2;
                        worksheet.Cell($"A{row}").Value = result.OriginalText;
                        worksheet.Cell($"B{row}").Value = result.TranslatedText;
                        worksheet.Cell($"C{row}").Value = result.Position;
                        worksheet.Cell($"D{row}").Value = result.ElementType;
                    }

                    // 自动调整列宽
                    worksheet.Columns().AdjustToContents();

                    // 保存文件
                    workbook.SaveAs(excelPath);
                }

                _logger.LogInformation($"翻译结果已导出到: {excelPath}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"导出翻译结果到Excel失败: {ex.Message}");
            }

            // 由于此方法目前是同步的，直接返回已完成的任务
            return Task.CompletedTask;
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
                    _logger.LogInformation($"进度: {progress:P0} - {message}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"进度回调执行失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证PPT文件
        /// </summary>
        public Dictionary<string, object> ValidatePPTFile(string filePath)
        {
            var validationResult = new Dictionary<string, object>
            {
                ["valid"] = false,
                ["error"] = string.Empty,
                ["file_size"] = 0L,
                ["slide_count"] = 0,
                ["supported_elements"] = new List<string>(),
                ["warnings"] = new List<string>()
            };

            try
            {
                // 检查文件大小
                var fileInfo = new FileInfo(filePath);
                validationResult["file_size"] = fileInfo.Length;

                // 检查文件大小限制（100MB）
                if (fileInfo.Length > 100 * 1024 * 1024)
                {
                    ((List<string>)validationResult["warnings"]).Add("文件大小超过100MB，可能影响处理性能");
                }

                // 尝试打开PPT文件
                using (var presentation = PresentationDocument.Open(filePath, false))
                {
                    var presentationPart = presentation.PresentationPart;
                    if (presentationPart != null)
                    {
                        var slides = presentationPart.SlideParts.ToList();
                        validationResult["slide_count"] = slides.Count;

                        // 检查幻灯片数量
                        if (slides.Count == 0)
                        {
                            ((List<string>)validationResult["warnings"]).Add("PPT文件不包含任何幻灯片");
                        }

                        // 检查支持的PPT元素
                        var supportedElements = new List<string>();
                        foreach (var slide in slides)
                        {
                            var shapes = slide.Slide.Descendants<Shape>();
                            foreach (var shape in shapes)
                            {
                                if (shape.TextBody != null)
                                    supportedElements.Add("文本");
                                if (shape.Descendants().Any(e => e.LocalName == "tbl"))
                                    supportedElements.Add("表格");
                                if (shape.Descendants<GraphicFrame>().Any())
                                    supportedElements.Add("图表");
                                if (shape.Descendants<GroupShape>().Any())
                                    supportedElements.Add("SmartArt");
                            }
                        }

                        validationResult["supported_elements"] = supportedElements.Distinct().ToList();
                        validationResult["valid"] = true;
                    }
                }
            }
            catch (Exception ex)
            {
                validationResult["error"] = ex.Message;
            }

            return validationResult;
        }

        /// <summary>
        /// 获取幻灯片数量
        /// </summary>
        public int GetSlideCount(string filePath)
        {
            try
            {
                using (var presentation = PresentationDocument.Open(filePath, false))
                {
                    var presentationPart = presentation.PresentationPart;
                    return presentationPart?.SlideParts?.Count() ?? 0;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"获取幻灯片数量失败: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// 获取PPT详细信息
        /// </summary>
        public Dictionary<string, object> GetSlideInfo(string filePath)
        {
            try
            {
                using (var presentation = PresentationDocument.Open(filePath, false))
                {
                    var presentationPart = presentation.PresentationPart;
                    if (presentationPart == null) return new Dictionary<string, object>();

                    var slides = presentationPart.SlideParts.ToList();
                    var info = new Dictionary<string, object>
                    {
                        ["slide_count"] = slides.Count,
                        ["slide_size"] = "标准尺寸", // 可以进一步获取实际尺寸
                        ["has_notes"] = presentationPart.NotesMasterPart != null,
                        ["has_charts"] = slides.Any(s => s.Slide.Descendants<GraphicFrame>().Any()),
                        ["has_tables"] = slides.Any(s => s.Slide.Descendants().Any(e => e.LocalName == "tbl")),
                        ["has_smartart"] = slides.Any(s => s.Slide.Descendants<GroupShape>().Any())
                    };

                    return info;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"获取PPT信息失败: {ex.Message}");
                return new Dictionary<string, object>();
            }
        }

        /// <summary>
        /// 按用户指定的批处理翻译流程处理PPT文档
        /// 1. 获取文档中所有文本及其位置
        /// 2. 写入CSV
        /// 3. 按术语长度替换术语
        /// 4. 并行翻译处理后的文本
        /// 5. 将翻译结果写回文档
        /// </summary>
        public async Task<string> BatchProcessPPTDocumentAsync(string filePath, string targetLanguage,
            Dictionary<string, string> terminology, bool createCopy = true)
        {
            // 更新进度：开始处理
            UpdateProgress(0.01, "开始批量处理PPT文档...");

            // 检查文件是否存在
            if (!File.Exists(filePath))
            {
                _logger.LogError($"文件不存在: {filePath}");
                throw new FileNotFoundException("文件不存在");
            }

            // 检查文件扩展名
            var extension = Path.GetExtension(filePath).ToLower();
            if (extension != ".pptx" && extension != ".ppt")
            {
                throw new ArgumentException($"不支持的PPT文件格式: {extension}，仅支持 .pptx 和 .ppt 格式");
            }

            // 在程序根目录下创建输出目录和日志目录
            var outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "输出");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var logDir = Path.Combine(outputDir, "日志");
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            // 获取输出路径，添加时间戳避免文件名冲突
            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var outputPath = Path.Combine(outputDir, $"{fileName}_带翻译_{timeStamp}{extension}");
            var csvPath = Path.Combine(logDir, $"{fileName}_翻译数据_{timeStamp}.xlsx");
            var preReplaceCsvPath = Path.Combine(logDir, $"{fileName}_术语预替换记录_{timeStamp}.csv");

            // 设置目标语言名称
            _targetLanguageName = targetLanguage;

            // 获取目标语言的术语表
            var targetTerminology = terminology ?? new Dictionary<string, string>();
            _logger.LogInformation($"使用{targetLanguage}术语表，包含 {targetTerminology.Count} 个术语");

            try
            {
                // 如果需要创建副本，则复制原文件，否则直接处理原文件
                string fileToProcess = filePath;
                if (createCopy)
                {
                    File.Copy(filePath, outputPath, true);
                    fileToProcess = outputPath;
                    UpdateProgress(0.05, "文件复制完成，开始提取文本...");
                }
                else
                {
                    UpdateProgress(0.05, "开始提取文本...");
                }

                _logger.LogInformation($"开始处理文件: {fileToProcess}");

                // 步骤1：提取所有文本及其位置信息
                var textEntries = new List<PPTTextEntry>();

                try
                {
                    using (var presentation = PresentationDocument.Open(fileToProcess, false))
                    {
                        var presentationPart = presentation.PresentationPart;
                        if (presentationPart == null)
                        {
                            throw new InvalidOperationException("无法打开PPT文档，可能文件已损坏");
                        }

                        var slides = presentationPart.SlideParts.ToList();
                        _logger.LogInformation($"PPT文档包含 {slides.Count} 张幻灯片");

                        // 提取每一张幻灯片的文本
                        for (int i = 0; i < slides.Count; i++)
                        {
                            var slide = slides[i];
                            var slideNumber = i + 1;

                            // 提取普通文本形状
                            int textShapeCount = 0;
                            foreach (var shape in slide.Slide.Descendants<Shape>())
                            {
                                if (shape.TextBody != null)
                                {
                                    ExtractTextFromShape(shape, textEntries, slideNumber, "文本");
                                    textShapeCount++;
                                }
                            }
                            _logger.LogInformation($"幻灯片 {slideNumber} 中提取了 {textShapeCount} 个文本形状");

                            // 提取表格中的文本
                            int tableCount = 0;
                            foreach (var table in slide.Slide.Descendants().Where(e => e.LocalName == "tbl"))
                            {
                                ExtractTextFromTable(table, textEntries, slideNumber);
                                tableCount++;
                            }
                            _logger.LogInformation($"幻灯片 {slideNumber} 中提取了 {tableCount} 个表格");

                            int beforeSmartArtCount = textEntries.Count;
                            if (_translateSmartArt)
                            {
                                ExtractTextFromSmartArt(slide, textEntries, slideNumber);
                            }
                            int smartArtCount = textEntries.Count - beforeSmartArtCount;
                            _logger.LogInformation($"幻灯片 {slideNumber} 中提取了 {smartArtCount} 条SmartArt文本");

                            var progress = 0.05 + (0.25 * (i + 1) / slides.Count);
                            UpdateProgress(progress, $"提取第 {slideNumber} 张幻灯片中的文本，共 {slides.Count} 张");
                        }

                        // 提取备注中的文本
                        if (_translateNotes)
                        {
                            var notesParts = slides.Select(s => s.NotesSlidePart).Where(n => n != null);
                            int notesCount = 0;
                            foreach (var notesPart in notesParts)
                            {
                                if (notesPart?.NotesSlide == null) continue;
                                foreach (var shape in notesPart.NotesSlide.Descendants<Shape>())
                                {
                                    if (shape.TextBody != null)
                                    {
                                        ExtractTextFromShape(shape, textEntries, 0, "备注");
                                        notesCount++;
                                    }
                                }
                            }
                            _logger.LogInformation($"提取了 {notesCount} 条备注");
                        }

                        // 提取母版中的文本
                        var masterParts = presentationPart.SlideMasterParts.ToList();
                        _logger.LogInformation($"PPT文档包含 {masterParts.Count} 个母版");
                        for (int i = 0; i < masterParts.Count; i++)
                        {
                            var masterPart = masterParts[i];
                            var masterIndex = -(i + 1); // 使用负数区分母版

                            // 提取母版形状文本
                            int masterShapeCount = 0;
                            foreach (var shape in masterPart.SlideMaster.Descendants<Shape>())
                            {
                                if (shape.TextBody != null)
                                {
                                    ExtractTextFromShape(shape, textEntries, masterIndex, "母版文本");
                                    masterShapeCount++;
                                }
                            }

                            // 提取母版表格文本
                            int masterTableCount = 0;
                            foreach (var table in masterPart.SlideMaster.Descendants().Where(e => e.LocalName == "tbl"))
                            {
                                ExtractTextFromTable(table, textEntries, masterIndex);
                                masterTableCount++;
                            }

                            _logger.LogInformation($"母版 {i + 1} 中提取了 {masterShapeCount} 个文本形状, {masterTableCount} 个表格");

                            // 提取版式(Layout)中的文本
                            var layoutParts = masterPart.SlideLayoutParts.ToList();
                            for (int j = 0; j < layoutParts.Count; j++)
                            {
                                var layoutPart = layoutParts[j];
                                // 使用特殊索引范围区分版式：-1000 - (i * 100) - j
                                // 例如：第1个母版的第1个版式 -> -1000 - 0 - 0 = -1000
                                int layoutIndex = -1000 - (i * 100) - j;

                                int layoutShapeCount = 0;
                                foreach (var shape in layoutPart.SlideLayout.Descendants<Shape>())
                                {
                                    if (shape.TextBody != null)
                                    {
                                        ExtractTextFromShape(shape, textEntries, layoutIndex, "版式文本");
                                        layoutShapeCount++;
                                    }
                                }

                                int layoutTableCount = 0;
                                foreach (var table in layoutPart.SlideLayout.Descendants().Where(e => e.LocalName == "tbl"))
                                {
                                    ExtractTextFromTable(table, textEntries, layoutIndex);
                                    layoutTableCount++;
                                }
                                _logger.LogDebug($"母版 {i + 1} 的版式 {j + 1} 中提取了 {layoutShapeCount} 个形状");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"提取文本阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                UpdateProgress(0.3, $"文本提取完成，共 {textEntries.Count} 条文本");

                if (textEntries.Count == 0)
                {
                    _logger.LogWarning("未提取到任何可翻译的文本，请检查PPT文档是否包含文本内容");
                    UpdateProgress(1.0, "处理完成，未找到可翻译的文本");
                    return outputPath;
                }

                // 步骤2：将文本信息写入Excel文件
                try
                {
                    await ExportToExcelAsync(textEntries, csvPath);
                    UpdateProgress(0.35, $"已将文本信息导出到: {csvPath}");
                }
                catch (Exception ex)
                {
                    _logger.LogError($"导出到Excel阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                // 步骤3：应用术语处理 - 按术语长度从长到短进行替换
                try
                {
                // 并发翻译前，按唯一键去重（避免重复翻译/写回）
                var distinctMap = new ConcurrentDictionary<string, PPTTextEntry>();
                foreach (var e in textEntries)
                {
                    // 唯一键：SlideNumber|ElementId|ParagraphIndex|Row|Col|OriginalText
                    var key = $"{e.SlideNumber}|{e.ElementId}|{e.ParagraphIndex}|{e.RowIndex}|{e.ColumnIndex}|{e.OriginalText}";
                    distinctMap.TryAdd(key, e);
                }
                var textToTranslate = distinctMap.Values.ToList();

                    if (_useTerminology && targetTerminology.Any())
                    {
                        // 与Word流程一致：MainWindow已经直接加载了外译中术语库，无需反转
                        Dictionary<string, string> termsForReplace = targetTerminology;

                        // 预处理替换
                        try
                        {
                            var allOriginalTexts = textEntries.Select(e => e.OriginalText).ToList();

                            if (!_isCnToForeign)
                            {
                                // 外译中：使用外译中术语库从长到短全局预替换，并记录CSV
                                var processedTexts = _termExtractor.PreprocessTextsWithMirroredTerms(allOriginalTexts, termsForReplace, preReplaceCsvPath, fileName);
                                for (int i = 0; i < textEntries.Count; i++)
                                {
                                    textEntries[i].ProcessedText = processedTexts[i];
                                }
                                _logger.LogInformation($"术语预处理完成（外译中：从长到短 + CSV记录）：{preReplaceCsvPath}");
                            }
                            else
                            {
                                // 中译外：逐条按长度优先替换
                                for (int i = 0; i < textEntries.Count; i++)
                                {
                                    var entry = textEntries[i];
                                    entry.ProcessedText = _termExtractor.ReplaceTermsInText(entry.OriginalText, termsForReplace, _isCnToForeign);
                                }
                                _logger.LogInformation("术语预处理完成（中译外：逐条替换）");
                            }
                        }
                        catch (Exception ex2)
                        {
                            _logger.LogError($"术语预处理失败，回退逐条替换: {ex2.Message}");
                            // 回退：逐条替换（按当前方向）
                            for (int i = 0; i < textEntries.Count; i++)
                            {
                                var entry = textEntries[i];
                                entry.ProcessedText = _termExtractor.ReplaceTermsInText(entry.OriginalText, termsForReplace, _isCnToForeign);
                            }
                        }

                        UpdateProgress(0.4, "术语预处理完成");
                    }
                    else
                    {
                        // 若不使用术语，则处理后的文本等于原始文本
                        foreach (var entry in textEntries)
                        {
                            entry.ProcessedText = entry.OriginalText;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"术语处理阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                // 步骤4：并行翻译处理后的文本
                try
                {
                    // 开始新的翻译批次，为当前文档创建独立的日志文件
                    _translationService.CurrentTranslator?.StartNewBatch();

                    var textToTranslate = textEntries.Where(e => !string.IsNullOrWhiteSpace(e.ProcessedText))
                                                     .GroupBy(e => $"{e.SlideNumber}|{e.ElementId}|{e.ParagraphIndex}|{e.RowIndex}|{e.ColumnIndex}|{e.OriginalText}")
                                                     .Select(g => g.First())
                                                     .ToList();
                    var totalTexts = textToTranslate.Count;

                    _logger.LogInformation($"开始并行翻译 {totalTexts} 条文本");

                    var translatedCount = 0;

                    _logger.LogInformation($"当前翻译器: {_translationService.CurrentTranslatorType}, 并发限制: {Translators.RWKVTranslator.MaxConcurrency}");

                    // 使用最稳妥的可等待并发：Task.WhenAll
                    var tasks = new List<Task>();
                    foreach (var entry in textToTranslate)
                    {
                        tasks.Add(Task.Run(async () =>
                        {
                            try
                            {
                                // 为每种文本类型选择合适的提示词
                                string promptType = entry.ElementType switch
                                {
                                    "title" => "title",
                                    "bullet" => "bullet",
                                    "table" => "content",
                                    "SmartArt" => "smartart",
                                    "备注" => "note",
                                    _ => "content"
                                };
                                var translationPrompt = _useImmersiveStyle
                                    ? _immersivePrompts.GetValueOrDefault(promptType, _immersivePrompts["content"])
                                    : _pptPrompts.GetValueOrDefault(promptType, _pptPrompts["content"]);

                                // 先按“空白分行”拆分段内子段：多个空白行视为分割边界
                                var textForTranslate = entry.ProcessedText ?? string.Empty;
                                var subSegments = SplitSegmentsByBlankLines(textForTranslate);

                                if (subSegments.Count <= 1)
                                {
                                    // 单段逻辑：前置过滤 → 翻译 → 过滤提示 → 回退
                                    if (ShouldSkipTranslation(textForTranslate))
                                    {
                                        entry.TranslatedText = textForTranslate;
                                    }
                                    else
                                    {
                                        var translatedText = await _translationService.TranslateTextAsync(
                                            textForTranslate,
                                            null,
                                            _sourceLang,
                                            _targetLang,
                                            translationPrompt,
                                            entry.OriginalText);

                                        translatedText = FilterTranslationResult(translatedText);
                                        if (string.IsNullOrWhiteSpace(translatedText))
                                        {
                                            translatedText = textForTranslate;
                                        }
                                        entry.TranslatedText = translatedText;
                                    }
                                }
                                else
                                {
                                    // 多子段：每个子段独立翻译，再用单个空白行拼接，避免多余空白
                                    var translatedSegments = new List<string>(subSegments.Count);
                                    foreach (var seg in subSegments)
                                    {
                                        // 逐行对齐：将子段按单换行拆成多行，逐行翻译，保持“一行英文→一行中文”
                                        var lines = (seg ?? string.Empty).Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                                        var translatedLines = new List<string>(lines.Length);

                                        foreach (var line in lines)
                                        {
                                            var trimmed = line?.TrimEnd() ?? string.Empty;
                                            if (string.IsNullOrWhiteSpace(trimmed))
                                            {
                                                translatedLines.Add(""); // 保留空行
                                                continue;
                                            }

                                            if (ShouldSkipTranslation(trimmed))
                                            {
                                                translatedLines.Add(trimmed);
                                                continue;
                                            }

                                            var lineTranslated = await _translationService.TranslateTextAsync(
                                                trimmed,
                                                null,
                                                _sourceLang,
                                                _targetLang,
                                                translationPrompt,
                                                entry.OriginalText);

                                            lineTranslated = FilterTranslationResult(lineTranslated);
                                            if (string.IsNullOrWhiteSpace(lineTranslated))
                                            {
                                                lineTranslated = trimmed; // 回退
                                            }
                                            translatedLines.Add(lineTranslated);
                                        }

                                        // 行内用单换行连接
                                        var segJoined = string.Join("\n", translatedLines);
                                        translatedSegments.Add(segJoined);
                                    }

                                    // 子段之间使用单个空白行分隔
                                    entry.TranslatedText = string.Join("\n\n", translatedSegments);
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError($"翻译文本失败: {ex.Message}");
                            }
                            finally
                            {
                                var count = Interlocked.Increment(ref translatedCount);
                                var progress = 0.4 + (0.5 * count / totalTexts);
                                UpdateProgress(progress, $"已翻译 {count}/{totalTexts} 条文本");
                            }
                        }));
                    }
                    await Task.WhenAll(tasks);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"文本翻译阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                UpdateProgress(0.9, "翻译完成，开始更新文档...");

                // 步骤5：将翻译结果写回文档（只写入原页面，不创建副本）
                try
                {
                    using (var presentation = PresentationDocument.Open(fileToProcess, true))
                    {
                        var presentationPart = presentation.PresentationPart;
                        if (presentationPart == null)
                        {
                            throw new InvalidOperationException("无法打开PPT文档进行更新");
                        }

                        var slides = presentationPart.SlideParts.ToList();

                        // 按幻灯片顺序处理，找到对应的文本并替换
                        for (int i = 0; i < slides.Count; i++)
                        {
                            var slide = slides[i];
                            var slideNumber = i + 1;

                            // 处理幻灯片中的所有形状
                            ReplaceTextInSlide(slide, textEntries, slideNumber);
                        }

                        // 处理备注
                        if (_translateNotes)
                        {
                            var notesParts = slides.Select(s => s.NotesSlidePart).Where(n => n != null);
                            foreach (var notesPart in notesParts)
                            {
                                if (notesPart != null) ReplaceTextInNotes(notesPart, textEntries);
                            }
                        }

                        // 处理母版
                        var masterPartsToUpdate = presentationPart.SlideMasterParts.ToList();
                        for (int i = 0; i < masterPartsToUpdate.Count; i++)
                        {
                            var masterPart = masterPartsToUpdate[i];
                            var masterIndex = -(i + 1);

                            // 替换母版形状文本
                            foreach (var shape in masterPart.SlideMaster.Descendants<Shape>())
                            {
                                if (shape.TextBody != null)
                                {
                                    ReplaceTextInShape(shape, textEntries, masterIndex, "母版文本");
                                }
                            }

                            // 替换母版表格文本
                            foreach (var table in masterPart.SlideMaster.Descendants().Where(e => e.LocalName == "tbl"))
                            {
                                ReplaceTextInTable(table, textEntries, masterIndex);
                            }

                            // 替换版式(Layout)文本
                            var layoutParts = masterPart.SlideLayoutParts.ToList();
                            for (int j = 0; j < layoutParts.Count; j++)
                            {
                                var layoutPart = layoutParts[j];
                                int layoutIndex = -1000 - (i * 100) - j;

                                foreach (var shape in layoutPart.SlideLayout.Descendants<Shape>())
                                {
                                    if (shape.TextBody != null)
                                    {
                                        ReplaceTextInShape(shape, textEntries, layoutIndex, "版式文本");
                                    }
                                }

                                foreach (var table in layoutPart.SlideLayout.Descendants().Where(e => e.LocalName == "tbl"))
                                {
                                    ReplaceTextInTable(table, textEntries, layoutIndex);
                                }
                            }
                        }

                        _logger.LogInformation("双语对照文档生成完成");

                        // 第二步：根据输出模式进行二次处理
                        _logger.LogInformation($"第二步：根据输出模式 '{_outputFormat}' 进行二次处理...");

                        if (_outputFormat == "bilingual")
                        {
                            // 双语对照模式：复制每一页，原页删除译文，复制页删除原文
                            ProcessBilingualModeAsync(presentationPart, textEntries);
                        }
                        else if (_outputFormat == "translation_only")
                        {
                            // 仅译文模式：删除原页，只保留译文页
                            ProcessTranslationOnlyModeAsync(presentationPart, textEntries);
                        }

                        // 深度验证并修复结构，减少“需要修复”提示
                        ValidateAndFixPresentationStructure(presentation);
                        // 保存文档
                        presentation.Save();
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"写回文档阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                // 更新Excel文件，加入翻译结果
                try
                {
                    await UpdateExcelWithTranslationsAsync(textEntries, csvPath);
                }
                catch (Exception ex)
                {
                    _logger.LogError($"更新Excel结果阶段出错: {ex.Message}");
                    _logger.LogError($"异常详情: {ex}");
                    throw;
                }

                UpdateProgress(1.0, "PPT批量翻译完成！");
                _logger.LogInformation($"PPT文档翻译完成，输出文件: {fileToProcess}");
                return fileToProcess;
            }
            catch (Exception ex)
            {
                _logger.LogError($"PPT文档批量处理失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 获取形状ID或名称作为标识
        /// </summary>
        private string GetShapeId(Shape shape)
        {
            string shapeId = "unknown";
            if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id != null)
            {
                // 将uint转换为string
                shapeId = shape.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value.ToString();
            }
            else if (shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value != null)
            {
                shapeId = shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value;
            }
            return shapeId;
        }

        private List<OpenXmlLeafTextElement> GetSmartArtTextElements(DiagramDataPart diagramDataPart)
        {
            var root = diagramDataPart.RootElement;
            if (root == null) return new List<OpenXmlLeafTextElement>();
            return root.Descendants<OpenXmlLeafTextElement>()
                .Where(e => e.LocalName == "t")
                .ToList();
        }

        /// <summary>
        /// 从形状中提取文本并记录位置信息
        /// </summary>
        private void ExtractTextFromShape(Shape shape, List<PPTTextEntry> entries, int slideNumber, string elementType)
        {
            try
            {
                if (shape.TextBody == null) return;

                // 获取形状ID或名称作为标识
                string shapeId = GetShapeId(shape);

                _logger.LogDebug($"开始提取形状文本，幻灯片: {slideNumber}, 形状ID: {shapeId}");

                var paragraphs = shape.TextBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>().ToList();
                int paragraphCount = paragraphs.Count;

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var paragraph = paragraphs[i];
                    var runs = paragraph.Elements<DocumentFormat.OpenXml.Drawing.Run>();
                    StringBuilder paragraphText = new StringBuilder();
                    int runCount = 0;

                    // 合并段落中的所有文本运行
                    foreach (var run in runs)
                    {
                        runCount++;
                        var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                        if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
                        {
                            paragraphText.Append(textElement.Text);
                            _logger.LogDebug($"提取到文本: '{textElement.Text}'");
                        }
                    }

                    _logger.LogDebug($"段落 {i} 包含 {runCount} 个文本运行");

                    string text = paragraphText.ToString().Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        // 确定文本类型
                        string textType = elementType;
                        if (elementType == "文本" || elementType == "母版文本" || elementType == "版式文本")
                        {
                            textType = DetermineTextType(shape, text);
                        }

                        // 生成唯一标识符，使用实际的段落索引i
                        string uniqueId = GenerateUniqueId(slideNumber, elementType, shapeId, i);
                        string contentHash = ComputeContentHash(text);

                        // 收集格式属性
                        var formatProps = ExtractFormatProperties(paragraph);

                        _logger.LogDebug($"提取到文本(类型:{textType}): '{text}', 唯一ID: {uniqueId}, 段落索引: {i}");

                        entries.Add(new PPTTextEntry
                        {
                            OriginalText = text,
                            SlideNumber = slideNumber,
                            Position = $"幻灯片 {slideNumber}, {elementType}, 形状ID: {shapeId}, 段落索引: {i}",
                            ElementType = textType,
                            ElementId = shapeId,
                            ParagraphIndex = i,
                            UniqueId = uniqueId,
                            ContentHash = contentHash,
                            TextType = textType,
                            FormatProperties = formatProps,
                            RunProperties = ExtractRunProperties(paragraph),
                            ContextInfo = $"幻灯片{slideNumber}_{elementType}_{shapeId}",
                            TranslationStatus = "pending"
                        });
                    }
                    else
                    {
                        _logger.LogDebug($"段落 {i} 不包含文本");
                    }
                }

                _logger.LogDebug($"形状包含 {paragraphCount} 个段落");

                if (paragraphCount == 0)
                {
                    _logger.LogWarning($"形状不包含任何段落文本，幻灯片: {slideNumber}, 形状ID: {shapeId}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"从形状提取文本时出错: {ex.Message}");
                _logger.LogError($"异常详情: {ex}");
            }
        }

        /// <summary>
        /// 从表格中提取文本并记录位置信息
        /// </summary>
        private void ExtractTextFromTable(OpenXmlElement table, List<PPTTextEntry> entries, int slideNumber)
        {
            try
            {
                var rows = table.Elements().Where(e => e.LocalName == "tr");
                int rowIndex = 0;

                foreach (var row in rows)
                {
                    var cells = row.Elements().Where(e => e.LocalName == "tc");
                    int cellIndex = 0;

                    foreach (var cell in cells)
                    {
                        var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                        if (textBody == null)
                        {
                            cellIndex++;
                            continue;
                        }

                        var paragraphs = textBody.Elements().Where(e => e.LocalName == "p");
                        StringBuilder cellText = new StringBuilder();

                        foreach (var paragraph in paragraphs)
                        {
                            var runs = paragraph.Elements().Where(e => e.LocalName == "r");
                            foreach (var run in runs)
                            {
                                var textElement = run.Elements().FirstOrDefault(e => e.LocalName == "t");
                                if (textElement != null)
                                {
                                    cellText.Append(textElement.InnerText);
                                }
                            }

                            // 如果有多个段落，添加换行符
                            cellText.Append(Environment.NewLine);
                        }

                        string text = cellText.ToString().Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            // 生成唯一标识符
                            string uniqueId = GenerateUniqueId(slideNumber, "table", $"row{rowIndex}_col{cellIndex}", 0);
                            string contentHash = ComputeContentHash(text);

                            // 收集格式属性
                            var formatProps = ExtractTableFormatProperties(cell);

                            entries.Add(new PPTTextEntry
                            {
                                OriginalText = text,
                                SlideNumber = slideNumber,
                                Position = $"幻灯片 {slideNumber}, 表格, 行: {rowIndex + 1}, 列: {cellIndex + 1}",
                                ElementType = "table",
                                ElementId = $"table_row{rowIndex}_col{cellIndex}",
                                RowIndex = rowIndex,
                                ColumnIndex = cellIndex,
                                UniqueId = uniqueId,
                                ContentHash = contentHash,
                                TextType = "table",
                                FormatProperties = formatProps,
                                RunProperties = ExtractTableRunProperties(cell),
                                ContextInfo = $"幻灯片{slideNumber}_table_row{rowIndex}_col{cellIndex}",
                                TranslationStatus = "pending"
                            });
                        }

                        cellIndex++;
                    }

                    rowIndex++;
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"从表格提取文本时出错: {ex.Message}");
            }
        }

        private void ExtractTextFromSmartArt(SlidePart slidePart, List<PPTTextEntry> entries, int slideNumber)
        {
            try
            {
                var diagramParts = slidePart.DiagramDataParts.ToList();
                if (diagramParts.Count == 0) return;

                foreach (var diagramPart in diagramParts)
                {
                    var diagramRelId = slidePart.GetIdOfPart(diagramPart) ?? "diagram";
                    var textElements = GetSmartArtTextElements(diagramPart);
                    for (int i = 0; i < textElements.Count; i++)
                    {
                        var textElement = textElements[i];
                        var text = textElement.Text?.Trim();
                        if (string.IsNullOrEmpty(text)) continue;

                        string uniqueId = GenerateUniqueId(slideNumber, "smartart", diagramRelId, i);
                        string contentHash = ComputeContentHash(text);

                        entries.Add(new PPTTextEntry
                        {
                            OriginalText = text,
                            SlideNumber = slideNumber,
                            Position = $"幻灯片 {slideNumber}, SmartArt, {diagramRelId}, 索引: {i}",
                            ElementType = "smartart",
                            ElementId = diagramRelId,
                            ParagraphIndex = i,
                            UniqueId = uniqueId,
                            ContentHash = contentHash,
                            TextType = "smartart",
                            ContextInfo = $"幻灯片{slideNumber}_smartart_{diagramRelId}",
                            TranslationStatus = "pending"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"从SmartArt提取文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 将文本信息导出到Excel文件
        /// </summary>
        private Task ExportToExcelAsync(List<PPTTextEntry> entries, string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    // 添加PPT信息工作表
                    var infoWorksheet = workbook.Worksheets.Add("PPT信息");
                    infoWorksheet.Cell("A1").Value = "PPT文件名";
                    infoWorksheet.Cell("B1").Value = Path.GetFileNameWithoutExtension(filePath);
                    
                    infoWorksheet.Cell("A2").Value = "导出时间";
                    infoWorksheet.Cell("B2").Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    
                    infoWorksheet.Cell("A3").Value = "总文本条数";
                    infoWorksheet.Cell("B3").Value = entries.Count;
                    
                    infoWorksheet.Cell("A4").Value = "目标语言";
                    infoWorksheet.Cell("B4").Value = _targetLanguageName;
                    
                    infoWorksheet.Cell("A5").Value = "翻译模式";
                    infoWorksheet.Cell("B5").Value = _outputFormat;
                    
                    // 设置信息工作表列宽
                    infoWorksheet.Column(1).Width = 15;
                    infoWorksheet.Column(2).Width = 30;

                    // 添加文本数据工作表
                    var worksheet = workbook.Worksheets.Add("PPT文本数据");

                    // 添加表头
                    worksheet.Cell("A1").Value = "幻灯片编号";
                    worksheet.Cell("B1").Value = "元素类型";
                    worksheet.Cell("C1").Value = "位置";
                    worksheet.Cell("D1").Value = "原始文本";
                    worksheet.Cell("E1").Value = "术语处理后文本";
                    worksheet.Cell("F1").Value = "翻译文本";
                    worksheet.Cell("G1").Value = "元素ID";
                    worksheet.Cell("H1").Value = "段落索引";
                    worksheet.Cell("I1").Value = "行索引";
                    worksheet.Cell("J1").Value = "列索引";

                    // 添加数据
                    for (int i = 0; i < entries.Count; i++)
                    {
                        var entry = entries[i];
                        var row = i + 2;

                        worksheet.Cell($"A{row}").Value = entry.SlideNumber;
                        worksheet.Cell($"B{row}").Value = entry.ElementType;
                        worksheet.Cell($"C{row}").Value = entry.Position;
                        worksheet.Cell($"D{row}").Value = entry.OriginalText;
                        worksheet.Cell($"E{row}").Value = entry.ProcessedText ?? "";
                        worksheet.Cell($"F{row}").Value = entry.TranslatedText ?? "";
                        worksheet.Cell($"G{row}").Value = entry.ElementId ?? "";
                        worksheet.Cell($"H{row}").Value = entry.ParagraphIndex;
                        worksheet.Cell($"I{row}").Value = entry.RowIndex;
                        worksheet.Cell($"J{row}").Value = entry.ColumnIndex;
                    }

                    // 自动调整列宽
                    worksheet.Columns().AdjustToContents();

                    // 保存文件
                    workbook.SaveAs(filePath);
                }

                _logger.LogInformation($"文本信息已导出到Excel: {filePath}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"导出文本信息到Excel失败: {ex.Message}");
            }

            // 由于此方法目前是同步的，直接返回已完成的任务
            return Task.CompletedTask;
        }

        /// <summary>
        /// 更新Excel文件，添加翻译结果
        private A.Paragraph GetOrCreateParagraph(DocumentFormat.OpenXml.Presentation.TextBody tBody, int index)
        {
            if (tBody == null) throw new ArgumentNullException(nameof(tBody));
            var paras = tBody.Elements<A.Paragraph>().ToList();
            while (paras.Count <= index)
            {
                var p = new A.Paragraph();
                var r = new A.Run();
                r.AppendChild(new A.Text { Text = "" });
                p.AppendChild(r);
                tBody.AppendChild(p);
                paras.Add(p);
            }
            return paras[index];
        }

        /// </summary>
        private Task UpdateExcelWithTranslationsAsync(List<PPTTextEntry> entries, string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1);

                    // 同步回写术语处理后文本(E)与翻译结果(F)
                    for (int i = 0; i < entries.Count; i++)
                    {
                        var entry = entries[i];
                        var row = i + 2;
                        worksheet.Cell($"E{row}").Value = entry.ProcessedText ?? "";
                        worksheet.Cell($"F{row}").Value = entry.TranslatedText ?? "";
                    }

                    // 保存文件
                    workbook.Save();
                }

                _logger.LogInformation($"Excel文件已更新翻译结果: {filePath}");
            }
            catch (Exception ex)
            {
                _logger.LogError($"更新Excel文件翻译结果失败: {ex.Message}");
            }

            // 由于此方法目前是同步的，直接返回已完成的任务
            return Task.CompletedTask;
        }

        /// <summary>
        /// 替换幻灯片中的文本
        /// </summary>
        private void ReplaceTextInSlide(SlidePart slidePart, List<PPTTextEntry> entries, int slideNumber)
        {
            try
            {
                // 处理普通文本形状
                foreach (var shape in slidePart.Slide.Descendants<Shape>())
                {
                    if (shape.TextBody != null)
                    {
                        ReplaceTextInShape(shape, entries, slideNumber);
                    }
                }

                // 处理表格
                foreach (var table in slidePart.Slide.Descendants().Where(e => e.LocalName == "tbl"))
                {
                    ReplaceTextInTable(table, entries, slideNumber);
                }
                if (_translateSmartArt)
                {
                    ReplaceTextInSmartArt(slidePart, entries, slideNumber);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"替换幻灯片文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换形状中的文本
        /// </summary>
        private void ReplaceTextInShape(Shape shape, List<PPTTextEntry> entries, int slideNumber, string elementType = "文本")
        {
            try
            {
                if (shape.TextBody == null) return;

                // 获取形状ID
                string shapeId = GetShapeId(shape);

                var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var paragraph = i < paragraphs.Count ? paragraphs[i] : GetOrCreateParagraph(shape.TextBody, i);
                    var runs = paragraph.Elements<A.Run>().ToList();

                    // 使用唯一标识符查找匹配的翻译条目
                    var matchingEntries = entries.Where(e =>
                        e.SlideNumber == slideNumber &&
                        e.ElementId == shapeId &&
                        e.ParagraphIndex == i &&
                        !string.IsNullOrEmpty(e.TranslatedText)).ToList();

                    if (!matchingEntries.Any())
                    {
                        // 如果没有匹配到翻译条目，在双语模式下也要保留原文
                        if (_outputFormat == "bilingual")
                        {
                            var originalText = string.Join("", runs.SelectMany(r => r.Elements<A.Text>().Select(t => t.Text)));
                            if (!string.IsNullOrWhiteSpace(originalText))
                            {
                                // 在清空之前保存格式属性
                                var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties ?? new A.RunProperties();

                                // 清空所有run
                                foreach (var r in runs) r.Remove();

                                // 创建原文run（使用保存的格式属性）
                                var originalRun = new A.Run { RunProperties = baseRunProps.CloneNode(true) as A.RunProperties };
                                originalRun.AppendChild(new A.Text { Text = originalText });
                                paragraph.AppendChild(originalRun);

                                // 添加分隔符##-----##
                                var separatorRun = new A.Run { RunProperties = baseRunProps.CloneNode(true) as A.RunProperties };
                                separatorRun.AppendChild(new A.Text { Text = "##-----##" });
                                paragraph.AppendChild(separatorRun);

                                // 添加空译文
                                var translatedRun = new A.Run { RunProperties = baseRunProps.CloneNode(true) as A.RunProperties };
                                translatedRun.AppendChild(new A.Text { Text = "" });
                                paragraph.AppendChild(translatedRun);

                                _logger.LogDebug($"双语模式：无翻译条目，保留原文，添加空译文");
                            }
                        }
                        continue;
                    }

                    var entry = matchingEntries.First(); // 如果有多个匹配，取第一个

                    if (_outputFormat == "bilingual")
                    {
                        // 双语：将原文和译文合并到同一个段落中，用##-----##分隔
                        var originalText = string.Join("", runs.SelectMany(r => r.Elements<A.Text>().Select(t => t.Text)));

                        // 在清空之前保存格式属性
                        // 注意：这里需要深度克隆，否则修改RunProperties会影响后续
                        var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties ?? new A.RunProperties();

                        // 清空所有run
                        foreach (var r in runs) r.Remove();

                        // 创建原文run（使用保存的格式属性的副本）
                        var originalRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                        originalRun.AppendChild(new A.Text { Text = originalText });
                        paragraph.AppendChild(originalRun);

                        // 添加分隔符##-----##（使用保存的格式属性的副本）
                        var separatorRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                        separatorRun.AppendChild(new A.Text { Text = "##-----##" });
                        paragraph.AppendChild(separatorRun);

                        // 添加译文（使用保存的格式属性的副本）
                        var translatedRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                        translatedRun.AppendChild(new A.Text { Text = entry.TranslatedText });
                        paragraph.AppendChild(translatedRun);

                        _logger.LogDebug($"双语模式：原文和译文已合并到同一段落，使用##-----##分隔");
                    }
                    else
                    {
                        // 仅译文模式：用译文替换原文
                        var baseRunProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties;

                        // 应用保存的格式属性
                        if (entry.FormatProperties != null && entry.FormatProperties.Any())
                        {
                            ApplyFormatProperties(paragraph, entry.FormatProperties);
                        }

                        // 保持原字体大小（100%）
                        if (baseRunProps?.FontSize != null && !_isCnToForeign)
                        {
                            int originalSize = baseRunProps.FontSize;
                            int newSize = (int)Math.Max(100, Math.Round(originalSize * 1.0)); // 100%
                            baseRunProps.FontSize = newSize;
                            _logger.LogDebug($"外译中字号调整 ({elementType}): {originalSize} -> {newSize} (100%)");
                        }

                        // 清空现有文本（包含 Field/Break），只保留段落属性
                        foreach (var fld in paragraph.Elements<A.Field>().ToList()) fld.Remove();
                        foreach (var br in paragraph.Elements<A.Break>().ToList()) br.Remove();
                        foreach (var run in runs) run.Remove();

                        // 多行文本写入（支持换行）
                        AppendTextWithLineBreaks(paragraph, entry.TranslatedText, baseRunProps);
                        EnsureEndParaRPr(paragraph);
                        ValidateAndFixTextShape(shape);
                    }

                    // 更新翻译状态
                    entry.TranslationStatus = "completed";
                    entry.IsPreserved = true;

                    _logger.LogDebug($"已替换/插入形状文本: 幻灯片 {slideNumber}, {elementType}, 形状ID: {shapeId}, 段落索引: {i}, 唯一ID: {entry.UniqueId}");
                }

                // 自动调整布局
                if (_autoAdjustLayout)
                {
                    AdjustTextShapeLayout(shape);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"替换形状文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换表格中的文本
        /// </summary>
        private void ReplaceTextInTable(OpenXmlElement table, List<PPTTextEntry> entries, int slideNumber)
        {
            try
            {
                var rows = table.Elements().Where(e => e.LocalName == "tr").ToList();

                for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
                {
                    var cells = rows[rowIndex].Elements().Where(e => e.LocalName == "tc").ToList();

                    for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                    {
                        // 使用唯一标识符查找匹配的翻译条目
                        var matchingEntries = entries.Where(e =>
                            e.SlideNumber == slideNumber &&
                            e.ElementType == "table" &&
                            e.RowIndex == rowIndex &&
                            e.ColumnIndex == colIndex &&
                            !string.IsNullOrEmpty(e.TranslatedText)).ToList();

                        if (matchingEntries.Any())
                        {
                            var entry = matchingEntries.First(); // 如果有多个匹配，取第一个
                            var cell = cells[colIndex];
                            var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");

                            if (textBody != null)
                            {
                                var paragraphs = textBody.Elements<A.Paragraph>().ToList();

                                if (_outputFormat == "bilingual")
                                {
                                    // 双语：将原文和译文合并到同一个段落中，用##-----##分隔
                                    var firstParagraph = paragraphs.FirstOrDefault();
                                    if (firstParagraph != null)
                                    {
                                        var runs = firstParagraph.Elements<A.Run>().ToList();
                                        var baseRunProps = (runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as A.RunProperties) ?? new A.RunProperties();
                                        
                                        // 获取原文文本
                                        var originalText = string.Join("", runs.SelectMany(r => r.Elements<A.Text>().Select(t => t.Text)));
                                        
                                        // 清空所有run
                                        foreach (var r in runs) r.Remove();
                                        
                                        // 添加原文
                                        var originalRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                                        originalRun.AppendChild(new A.Text { Text = originalText });
                                        firstParagraph.AppendChild(originalRun);
                                        
                                        // 添加分隔符##-----##
                                        var separatorRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                                        separatorRun.AppendChild(new A.Text { Text = "##-----##" });
                                        firstParagraph.AppendChild(separatorRun);
                                        
                                        // 添加译文
                                        var translatedRun = new A.Run { RunProperties = (A.RunProperties)baseRunProps.CloneNode(true) };
                                        translatedRun.AppendChild(new A.Text { Text = entry.TranslatedText });
                                        firstParagraph.AppendChild(translatedRun);
                                        
                                        _logger.LogDebug($"双语模式表格：原文和译文已合并到同一段落，使用##-----##分隔");
                                    }
                                }
                                else
                                {
                                    // 仅译文模式：用译文替换原文
                                    // 先基于原有样式获取基准Run属性（在删除段落前）
                                    A.RunProperties? baseRunProps = null;
                                    var oldParas = textBody.Elements<A.Paragraph>().ToList();
                                    var baseRun = oldParas.SelectMany(p => p.Elements<A.Run>()).FirstOrDefault();
                                    if (baseRun?.RunProperties != null)
                                    {
                                        baseRunProps = baseRun.RunProperties.CloneNode(true) as A.RunProperties;
                                    }

                                    // 删除所有现有段落，只保留 a:bodyPr 与 a:lstStyle
                                    foreach (var p in oldParas)
                                    {
                                        p.Remove();
                                    }

                                    // 确保 txBody 具备必要的结构
                                    EnsureTableTextBodyScaffold(textBody);

                                    // 仅译文模式：用译文替换原文
                                    var newParagraph = new A.Paragraph { ParagraphProperties = new A.ParagraphProperties() };
                                    var newRunProps = baseRunProps?.CloneNode(true) as A.RunProperties;

                                    // 应用保存的格式属性
                                    if (entry.FormatProperties != null && entry.FormatProperties.Any())
                                    {
                                        ApplyFormatProperties(newParagraph, entry.FormatProperties);
                                    }

                                    // 保持原字体大小（100%）
                                    if (newRunProps?.FontSize != null && !_isCnToForeign)
                                    {
                                        int originalSize = newRunProps.FontSize;
                                        int newSize = (int)Math.Max(100, Math.Round(originalSize * 1.0));
                                        newRunProps.FontSize = newSize;
                                        _logger.LogDebug($"外译中表格字号调整: {originalSize} -> {newSize} (100%)");
                                    }

                                    AppendTextWithLineBreaks(newParagraph, entry.TranslatedText, newRunProps);
                                    textBody.AppendChild(newParagraph);

                                    // 再次验证并修复单元格文本体结构
                                    EnsureTableTextBodyScaffold(textBody, ensureParagraphsNonEmpty: true);
                                }

                                // 更新翻译状态
                                entry.TranslationStatus = "completed";
                                entry.IsPreserved = true;

                                _logger.LogDebug($"已替换表格文本: 幻灯片 {slideNumber}, 表格, 行: {rowIndex + 1}, 列: {colIndex + 1}, 唯一ID: {entry.UniqueId}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"替换表格文本时出错: {ex.Message}");
            }
        }

        private void ReplaceTextInSmartArt(SlidePart slidePart, List<PPTTextEntry> entries, int slideNumber)
        {
            try
            {
                var diagramParts = slidePart.DiagramDataParts.ToList();
                if (diagramParts.Count == 0) return;

                foreach (var diagramPart in diagramParts)
                {
                    var diagramRelId = slidePart.GetIdOfPart(diagramPart) ?? "diagram";
                    var textElements = GetSmartArtTextElements(diagramPart);
                    for (int i = 0; i < textElements.Count; i++)
                    {
                        var textElement = textElements[i];
                        var matchingEntries = entries.Where(e =>
                            e.SlideNumber == slideNumber &&
                            e.ElementType == "smartart" &&
                            e.ElementId == diagramRelId &&
                            e.ParagraphIndex == i &&
                            !string.IsNullOrEmpty(e.TranslatedText)).ToList();

                        if (!matchingEntries.Any()) continue;

                        var entry = matchingEntries.First();
                        if (_outputFormat == "bilingual")
                        {
                            textElement.Text = $"{entry.OriginalText}##-----##{entry.TranslatedText}";
                        }
                        else
                        {
                            textElement.Text = entry.TranslatedText;
                        }

                        entry.TranslationStatus = "completed";
                        entry.IsPreserved = true;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"替换SmartArt文本时出错: {ex.Message}");
            }
        }

        private void RemoveOriginalTextFromSmartArt(SlidePart slidePart)
        {
            try
            {
                var diagramParts = slidePart.DiagramDataParts.ToList();
                if (diagramParts.Count == 0) return;

                foreach (var diagramPart in diagramParts)
                {
                    var textElements = GetSmartArtTextElements(diagramPart);
                    foreach (var textElement in textElements)
                    {
                        var text = textElement.Text;
                        if (string.IsNullOrEmpty(text)) continue;
                        var separatorIndex = text.IndexOf("##-----##", StringComparison.Ordinal);
                        if (separatorIndex < 0) continue;
                        var translated = text.Substring(separatorIndex + "##-----##".Length);
                        textElement.Text = translated;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"删除SmartArt原文失败: {ex.Message}");
            }
        }

        private void RemoveTranslatedTextFromSmartArt(SlidePart slidePart)
        {
            try
            {
                var diagramParts = slidePart.DiagramDataParts.ToList();
                if (diagramParts.Count == 0) return;

                foreach (var diagramPart in diagramParts)
                {
                    var textElements = GetSmartArtTextElements(diagramPart);
                    foreach (var textElement in textElements)
                    {
                        var text = textElement.Text;
                        if (string.IsNullOrEmpty(text)) continue;
                        var separatorIndex = text.IndexOf("##-----##", StringComparison.Ordinal);
                        if (separatorIndex < 0) continue;
                        var original = text.Substring(0, separatorIndex);
                        textElement.Text = original;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"删除SmartArt译文失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 替换备注中的文本
        /// </summary>
        private void ReplaceTextInNotes(NotesSlidePart notesPart, List<PPTTextEntry> entries)
        {
            try
            {
                foreach (var shape in notesPart.NotesSlide.Descendants<Shape>())
                {
                    if (shape.TextBody != null)
                    {
                        // 备注页的幻灯片编号为0
                        ReplaceTextInShape(shape, entries, 0, "备注");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"替换备注文本时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从幻灯片中删除中文内容，只保留译文
        /// </summary>
        private void RemoveChineseTextFromSlide(SlidePart slidePart)
        {
            try
            {
                // 处理普通文本形状
                foreach (var shape in slidePart.Slide.Descendants<Shape>())
                {
                    if (shape.TextBody != null)
                    {
                        RemoveChineseTextFromShape(shape);
                    }
                }

                // 处理表格
                foreach (var table in slidePart.Slide.Descendants().Where(e => e.LocalName == "tbl"))
                {
                    RemoveChineseTextFromTable(table);
                }

                // 处理SmartArt
                foreach (var groupShape in slidePart.Slide.Descendants<GroupShape>())
                {
                    foreach (var subShape in groupShape.Descendants<Shape>())
                    {
                        if (subShape.TextBody != null)
                        {
                            RemoveChineseTextFromShape(subShape);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"删除幻灯片中文内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从形状中删除中文内容，只保留译文
        /// </summary>
        private void RemoveChineseTextFromShape(Shape shape)
        {
            try
            {
                if (shape.TextBody == null) return;

                var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
                var paragraphsToRemove = new List<A.Paragraph>();

                for (int i = 0; i < paragraphs.Count; i++)
                {
                    var paragraph = paragraphs[i];
                    var text = GetParagraphText(paragraph);

                    // 如果段落包含中文字符，标记为删除
                    if (ContainsChinese(text))
                    {
                        paragraphsToRemove.Add(paragraph);
                    }
                }

                // 删除包含中文的段落
                foreach (var paragraph in paragraphsToRemove)
                {
                    paragraph.Remove();
                }

                _logger.LogDebug($"已从形状中删除 {paragraphsToRemove.Count} 个中文段落");
            }
            catch (Exception ex)
            {
                _logger.LogError($"删除形状中文内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从表格中删除中文内容
        /// </summary>
        private void RemoveChineseTextFromTable(OpenXmlElement table)
        {
            try
            {
                var rows = table.Elements().Where(e => e.LocalName == "tr").ToList();

                foreach (var row in rows)
                {
                    var cells = row.Elements().Where(e => e.LocalName == "tc").ToList();

                    foreach (var cell in cells)
                    {
                        var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");

                        if (textBody != null)
                        {
                            var paragraphs = textBody.Elements().Where(e => e.LocalName == "p").ToList();
                            var paragraphsToRemove = new List<OpenXmlElement>();

                            foreach (var paragraph in paragraphs)
                            {
                                var text = GetElementText(paragraph);

                                // 如果包含中文，标记删除
                                if (ContainsChinese(text))
                                {
                                    paragraphsToRemove.Add(paragraph);
                                }
                            }

                            // 删除包含中文的段落
                            foreach (var paragraph in paragraphsToRemove)
                            {
                                paragraph.Remove();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"删除表格中文内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 从备注中删除中文内容
        /// </summary>
        private void RemoveChineseTextFromNotes(NotesSlidePart notesPart)
        {
            try
            {
                var shapes = notesPart.NotesSlide.Descendants<Shape>().ToList();

                foreach (var shape in shapes)
                {
                    if (shape.TextBody != null)
                    {
                        RemoveChineseTextFromShape(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"删除备注中文内容时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 将多行文本加入段落，使用换行标记，尽量继承样式
        /// </summary>
        private void AppendTextWithLineBreaks(A.Paragraph paragraph, string text, A.RunProperties? baseRunProps)
        {
            if (paragraph == null) return;
            if (text == null) text = string.Empty;
            var lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

            for (int i = 0; i < lines.Length; i++)
            {
                var run = new A.Run();
                if (baseRunProps != null)
                {
                    run.RunProperties = baseRunProps.CloneNode(true) as A.RunProperties;
                }
                var t = new A.Text { Text = lines[i] };
                run.AppendChild(t);
                paragraph.AppendChild(run);

                if (i < lines.Length - 1)
                {
                    paragraph.AppendChild(new A.Break());
                }
            }
        }
            /// <summary>
            /// 确保段落末尾有 endParaRPr（某些 PPT 主题需要），避免“需要修复”提示
            /// </summary>
            private void EnsureEndParaRPr(A.Paragraph paragraph)
            {
                try
                {
                    if (paragraph == null) return;
                    var end = paragraph.GetFirstChild<A.EndParagraphRunProperties>();
                    if (end == null)
                    {
                        end = new A.EndParagraphRunProperties();
                        paragraph.Append(end);
                    }
                    if (end.Language == null)
                    {
                        end.Language = _isCnToForeign ? "en-US" : "zh-CN"; // 随方向设置语言
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"EnsureEndParaRPr 失败: {ex.Message}");
                }
            }



        /// <summary>
        /// 检查文本是否包含中文字符
        /// </summary>
        private bool ContainsChinese(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;

            // 检查是否包含中文字符（Unicode范围：\u4e00-\u9fff）
            return System.Text.RegularExpressions.Regex.IsMatch(text, @"[\u4e00-\u9fff]");
        }

        /// <summary>
        /// 根据元素类型获取字号调整比例
        /// </summary>
        private double GetFontSizeRatio(string elementType)
        {
            return elementType switch
            {
                // 所有文本保持原字体大小（100%）
                "SmartArt" => 1.0,
                "GroupShape" => 1.0,
                "几何图形" => 1.0,

                // 所有翻译结果文本保持原字体大小（100%）
                "文本" => 1.0,
                "title" => 1.0,
                "content" => 1.0,
                "bullet" => 1.0,

                // 表格和其他元素保持原字体大小（100%）
                "table" => 1.0,
                "备注" => 1.0,
                "note" => 1.0,

                // 默认保持原字体大小（100%）
                _ => 1.0
            };
        }

            /// <summary>
            /// 确保表格单元格 txBody 具备必要的结构：a:bodyPr、a:lstStyle、至少一个段落与Run
            /// </summary>
            private void EnsureTableTextBodyScaffold(OpenXmlElement textBodyElement, bool ensureParagraphsNonEmpty = false)
            {
                try
                {
                    if (textBodyElement == null) return;

                    // a:bodyPr
                    var bodyPr = textBodyElement.Elements<A.BodyProperties>().FirstOrDefault();
                    if (bodyPr == null)
                    {
                        bodyPr = new A.BodyProperties();
                        textBodyElement.PrependChild(bodyPr);
                    }

                    // a:lstStyle
                    var lstStyle = textBodyElement.Elements<A.ListStyle>().FirstOrDefault();
                    if (lstStyle == null)
                    {
                        lstStyle = new A.ListStyle();
                        textBodyElement.InsertAfter(lstStyle, bodyPr);
                    }

                    // 段落
                    var paras = textBodyElement.Elements<A.Paragraph>().ToList();
                    if (!paras.Any())
                    {
                        var p = new A.Paragraph();
                        var r = new A.Run();
                        r.AppendChild(new A.Text { Text = "" });
                        p.AppendChild(r);
                        textBodyElement.AppendChild(p);
                        paras.Add(p);
                    }
                    EnsureTableTextBodyScaffold_Tail(textBodyElement, paras, ensureParagraphsNonEmpty);

                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"EnsureTableTextBodyScaffold 失败: {ex.Message}");
                }
            }


            /// <summary>
            /// 清理表格单元格的文本体：移除空段落、空 run、裁剪首尾空行
            /// </summary>
            private void CleanupTableTextBody(OpenXmlElement textBodyElement)
            {
                try
                {
                    if (textBodyElement == null) return;

                    // 移除空 run（无 A.Text 或全空白）
                    foreach (var run in textBodyElement.Descendants<A.Run>().ToList())
                    {
                        var t = run.GetFirstChild<A.Text>();
                        if (t == null)
                        {
                            run.Remove();
                            continue;
                        }
                        if (string.IsNullOrWhiteSpace(t.Text))
                        {
                            run.Remove();
                        }
                    }

                    // 清理段落：去掉没有 run 的段落
                    var paras = textBodyElement.Elements<A.Paragraph>().ToList();
                    foreach (var p in paras)
                    {
                        if (!p.Elements<A.Run>().Any())
                        {
                            p.Remove();
                        }
                    }

                    // 首尾多余空行：如果第一个或最后一个段落只有空白，则移除
                    paras = textBodyElement.Elements<A.Paragraph>().ToList();
                    if (paras.Count >= 1)
                    {
                        var first = paras.First();
                        var last = paras.Last();
                        bool IsParaEmpty(A.Paragraph para)
                        {
                            var text = string.Join("", para.Elements<A.Run>().Select(r => r.GetFirstChild<A.Text>()?.Text ?? ""));
                            return string.IsNullOrWhiteSpace(text);
                        }
                        if (IsParaEmpty(first) && paras.Count > 1) first.Remove();
                        if (paras.Count > 1)
                        {
                            // 重新获取最后一段（可能因为删除首段发生变化）
                            last = textBodyElement.Elements<A.Paragraph>().LastOrDefault() ?? last;
                            if (last != null && IsParaEmpty(last)) last.Remove();
                        }
                    }

                    // 兜底：确保至少一个非空段落
                    if (!textBodyElement.Elements<A.Paragraph>().Any())
                    {
                        var p = new A.Paragraph();
                        var r = new A.Run();
                        r.AppendChild(new A.Text { Text = "" });
                        p.AppendChild(r);
                        textBodyElement.AppendChild(p);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"CleanupTableTextBody 失败: {ex.Message}");
                }
            }

            // 回到 EnsureTableTextBodyScaffold 方法体的剩余部分
            private void EnsureTableTextBodyScaffold_Tail(OpenXmlElement textBodyElement, List<A.Paragraph> paras, bool ensureParagraphsNonEmpty)
            {
                try
                {
                    if (ensureParagraphsNonEmpty)
                    {
                        foreach (var p in paras)
                        {
                            if (p.ParagraphProperties == null)
                                p.ParagraphProperties = new A.ParagraphProperties();
                            if (!p.Elements<A.Run>().Any())
                            {
                                var r = new A.Run();
                                r.AppendChild(new A.Text { Text = "" });
                                p.AppendChild(r);
                            }
                            else
                            {
                                foreach (var run in p.Elements<A.Run>())
                                {
                                    if (!run.Elements<A.Text>().Any())
                                    {
                                        run.AppendChild(new A.Text { Text = "" });
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"EnsureTableTextBodyScaffold_Tail 失败: {ex.Message}");
                }
            }



        /// <summary>
        /// 获取OpenXml元素的文本内容
        /// </summary>
        private string GetElementText(OpenXmlElement element)
        {
            if (element == null) return string.Empty;

            var textElements = element.Descendants().Where(e => e.LocalName == "t");
            return string.Join("", textElements.Select(t => t.InnerText));
        }

        /// <summary>
        /// 生成唯一标识符
        /// </summary>
        private string GenerateUniqueId(int slideNumber, string elementType, string elementId, int paragraphIndex)
        {
            return $"{slideNumber}_{elementType}_{elementId}_p{paragraphIndex}_{Guid.NewGuid().ToString("N").Substring(0, 8)}";
        }

        /// <summary>
        /// 计算内容哈希值
        /// </summary>
        private string ComputeContentHash(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes(text);
                var hash = sha256.ComputeHash(bytes);
                return Convert.ToBase64String(hash).Substring(0, 16);
            }
        }

        /// <summary>
        /// 提取段落的格式属性
        /// </summary>
        private Dictionary<string, string> ExtractFormatProperties(A.Paragraph paragraph)
        {
            var props = new Dictionary<string, string>();
            
            try
            {
                if (paragraph?.ParagraphProperties != null)
                {
                    var paraProps = paragraph.ParagraphProperties;
                    
                    if (paraProps.Alignment != null)
                        props["Alignment"] = paraProps.Alignment.Value.ToString();
                    
                    if (paraProps.LeftMargin != null)
                        props["LeftMargin"] = paraProps.LeftMargin.Value.ToString();
                    
                    if (paraProps.RightMargin != null)
                        props["RightMargin"] = paraProps.RightMargin.Value.ToString();
                    
                    if (paraProps.SpaceBefore != null)
                        props["SpaceBefore"] = paraProps.SpaceBefore.ToString() ?? "";
                    
                    if (paraProps.SpaceAfter != null)
                        props["SpaceAfter"] = paraProps.SpaceAfter.ToString() ?? "";
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"提取段落格式属性失败: {ex.Message}");
            }
            
            return props;
        }

        /// <summary>
        /// 提取Run的格式属性
        /// </summary>
        private List<string> ExtractRunProperties(A.Paragraph paragraph)
        {
            var props = new List<string>();
            
            try
            {
                var runs = paragraph.Elements<A.Run>();
                foreach (var run in runs)
                {
                    if (run?.RunProperties != null)
                    {
                        var runProps = run.RunProperties;
                        var propList = new List<string>();
                        
                        if (runProps.FontSize != null)
                            propList.Add($"FontSize:{runProps.FontSize}");
                        
                        if (runProps.Bold != null)
                            propList.Add($"Bold:{runProps.Bold}");
                        
                        if (runProps.Italic != null)
                            propList.Add($"Italic:{runProps.Italic}");
                        
                        if (runProps.Underline != null)
                            propList.Add($"Underline:{runProps.Underline}");
                        
                        var solidFill = runProps.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
                        if (solidFill != null)
                        {
                            var srgbClr = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
                            if (srgbClr != null && srgbClr.GetAttributes().Any())
                            {
                                var val = srgbClr.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                                if (!string.IsNullOrEmpty(val.Value))
                                    propList.Add($"Color:{val.Value}");
                            }
                        }
                        
                        var latin = runProps.Elements().FirstOrDefault(e => e.LocalName == "latin");
                        if (latin != null && latin.GetAttributes().Any())
                        {
                            var typeface = latin.GetAttributes().FirstOrDefault(a => a.LocalName == "typeface");
                            propList.Add($"Font:{typeface.Value}");
                        }
                        
                        props.Add(string.Join(";", propList));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"提取Run格式属性失败: {ex.Message}");
            }
            
            return props;
        }

        /// <summary>
        /// 提取表格单元格的格式属性
        /// </summary>
        private Dictionary<string, string> ExtractTableFormatProperties(OpenXmlElement cell)
        {
            var props = new Dictionary<string, string>();
            
            try
            {
                var cellProps = cell.Elements().FirstOrDefault(e => e.LocalName == "tcPr");
                if (cellProps != null)
                {
                    var marL = cellProps.Elements().FirstOrDefault(e => e.LocalName == "marL");
                    if (marL != null)
                        props["MarginLeft"] = marL.InnerText;
                    
                    var marR = cellProps.Elements().FirstOrDefault(e => e.LocalName == "marR");
                    if (marR != null)
                        props["MarginRight"] = marR.InnerText;
                    
                    var marT = cellProps.Elements().FirstOrDefault(e => e.LocalName == "marT");
                    if (marT != null)
                        props["MarginTop"] = marT.InnerText;
                    
                    var marB = cellProps.Elements().FirstOrDefault(e => e.LocalName == "marB");
                    if (marB != null)
                        props["MarginBottom"] = marB.InnerText;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"提取表格格式属性失败: {ex.Message}");
            }
            
            return props;
        }

        /// <summary>
        /// 提取表格单元格的Run格式属性
        /// </summary>
        private List<string> ExtractTableRunProperties(OpenXmlElement cell)
        {
            var props = new List<string>();
            
            try
            {
                var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                if (textBody != null)
                {
                    var paragraphs = textBody.Elements().Where(e => e.LocalName == "p");
                    foreach (var paragraph in paragraphs)
                    {
                        var runs = paragraph.Elements().Where(e => e.LocalName == "r");
                        foreach (var run in runs)
                        {
                            var runProps = run.Elements().FirstOrDefault(e => e.LocalName == "rPr");
                            if (runProps != null)
                            {
                                var propList = new List<string>();
                                
                                var sz = runProps.Elements().FirstOrDefault(e => e.LocalName == "sz");
                                if (sz != null)
                                    propList.Add($"FontSize:{sz.InnerText}");
                                
                                var b = runProps.Elements().FirstOrDefault(e => e.LocalName == "b");
                                if (b != null)
                                    propList.Add($"Bold:{b.InnerText}");
                                
                                var i = runProps.Elements().FirstOrDefault(e => e.LocalName == "i");
                                if (i != null)
                                    propList.Add($"Italic:{i.InnerText}");
                                
                                var u = runProps.Elements().FirstOrDefault(e => e.LocalName == "u");
                                if (u != null)
                                    propList.Add($"Underline:{u.InnerText}");
                                
                                var solidFill = runProps.Elements().FirstOrDefault(e => e.LocalName == "solidFill");
                                if (solidFill != null)
                                {
                                    var srgbClr = solidFill.Elements().FirstOrDefault(e => e.LocalName == "srgbClr");
                                    if (srgbClr != null && srgbClr.GetAttributes().Any())
                                    {
                                        var val = srgbClr.GetAttributes().FirstOrDefault(a => a.LocalName == "val");
                                        if (!string.IsNullOrEmpty(val.Value))
                                            propList.Add($"Color:{val.Value}");
                                    }
                                }
                                
                                var latin = runProps.Elements().FirstOrDefault(e => e.LocalName == "latin");
                                if (latin != null && latin.GetAttributes().Any())
                                {
                                    var typeface = latin.GetAttributes().FirstOrDefault(a => a.LocalName == "typeface");
                                    if (!string.IsNullOrEmpty(typeface.Value))
                                        propList.Add($"Font:{typeface.Value}");
                                }
                                
                                props.Add(string.Join(";", propList));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"提取表格Run格式属性失败: {ex.Message}");
            }
            
            return props;
        }

        /// <summary>
        /// 应用格式属性到段落
        /// </summary>
        private void ApplyFormatProperties(A.Paragraph paragraph, Dictionary<string, string> formatProperties)
        {
            try
            {
                if (paragraph == null || formatProperties == null || !formatProperties.Any())
                    return;

                if (paragraph.ParagraphProperties == null)
                {
                    paragraph.ParagraphProperties = new A.ParagraphProperties();
                }

                var paraProps = paragraph.ParagraphProperties;

                if (formatProperties.ContainsKey("Alignment"))
                {
                    if (Enum.TryParse(formatProperties["Alignment"], out A.TextAlignmentTypeValues alignment))
                    {
                        paraProps.Alignment = alignment;
                    }
                }

                if (formatProperties.ContainsKey("LeftMargin"))
                {
                    if (int.TryParse(formatProperties["LeftMargin"], out int leftMargin))
                    {
                        paraProps.LeftMargin = leftMargin;
                    }
                }

                if (formatProperties.ContainsKey("RightMargin"))
                {
                    if (int.TryParse(formatProperties["RightMargin"], out int rightMargin))
                    {
                        paraProps.RightMargin = rightMargin;
                    }
                }

                if (formatProperties.ContainsKey("SpaceBefore"))
                {
                    if (int.TryParse(formatProperties["SpaceBefore"], out int spaceBefore))
                    {
                        paraProps.SpaceBefore = new A.SpaceBefore(new Int32Value(spaceBefore)!);
                    }
                }

                if (formatProperties.ContainsKey("SpaceAfter"))
                {
                    if (int.TryParse(formatProperties["SpaceAfter"], out int spaceAfter))
                    {
                        paraProps.SpaceAfter = new A.SpaceAfter(new Int32Value(spaceAfter)!);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"应用段落格式属性失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从图表中提取文本
        /// </summary>
        private List<PPTTextEntry> ExtractChartText(ChartPart chartPart, int slideNumber, string chartId, int chartIndex)
        {
            var entries = new List<PPTTextEntry>();
            
            try
            {
                var chart = chartPart?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                if (chart == null) return entries;

                // 提取图表标题
                var title = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                if (title != null)
                {
                    var chartText = title.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                    if (chartText != null)
                    {
                        var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                        if (richText != null)
                        {
                            var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                            if (textPara != null)
                            {
                                var text = GetElementText(textPara);
                                if (!string.IsNullOrEmpty(text))
                                {
                                    var uniqueId = GenerateUniqueId(slideNumber, "chart", chartId, chartIndex);
                                    var contentHash = ComputeContentHash(text);
                                    
                                    entries.Add(new PPTTextEntry
                                    {
                                        OriginalText = text,
                                        SlideNumber = slideNumber,
                                        Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}",
                                        ElementType = "chart",
                                        ElementId = chartId,
                                        UniqueId = uniqueId,
                                        ContentHash = contentHash,
                                        TextType = "chart_title",
                                        ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_title",
                                        TranslationStatus = "pending"
                                    });
                                }
                            }
                        }
                    }
                }

                // 提取图表轴标签
                var plotArea = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
                if (plotArea != null)
                {
                    // 提取数值轴标签
                    var valueAxes = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>();
                    int axisIndex = 0;
                    
                    foreach (var axis in valueAxes)
                    {
                        var axisTitle = axis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                        if (axisTitle != null)
                        {
                            var chartText = axisTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                            if (chartText != null)
                            {
                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                if (richText != null)
                                {
                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                    if (textPara != null)
                                    {
                                        var text = GetElementText(textPara);
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            var uniqueId = GenerateUniqueId(slideNumber, "chart", $"{chartId}_valueAxis{axisIndex}", chartIndex);
                                            var contentHash = ComputeContentHash(text);
                                            
                                            entries.Add(new PPTTextEntry
                                            {
                                                OriginalText = text,
                                                SlideNumber = slideNumber,
                                                Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}, 数值轴: {axisIndex}",
                                                ElementType = "chart",
                                                ElementId = $"{chartId}_valueAxis{axisIndex}",
                                                UniqueId = uniqueId,
                                                ContentHash = contentHash,
                                                TextType = "chart_axis",
                                                ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_valueAxis{axisIndex}",
                                                TranslationStatus = "pending"
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        axisIndex++;
                    }

                    // 提取类别轴标签
                    var categoryAxes = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>();
                    int catAxisIndex = 0;
                    
                    foreach (var catAxis in categoryAxes)
                    {
                        var axisTitle = catAxis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                        if (axisTitle != null)
                        {
                            var chartText = axisTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                            if (chartText != null)
                            {
                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                if (richText != null)
                                {
                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                    if (textPara != null)
                                    {
                                        var text = GetElementText(textPara);
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            var uniqueId = GenerateUniqueId(slideNumber, "chart", $"{chartId}_catAxis{catAxisIndex}", chartIndex);
                                            var contentHash = ComputeContentHash(text);
                                            
                                            entries.Add(new PPTTextEntry
                                            {
                                                OriginalText = text,
                                                SlideNumber = slideNumber,
                                                Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}, 类别轴: {catAxisIndex}",
                                                ElementType = "chart",
                                                ElementId = $"{chartId}_catAxis{catAxisIndex}",
                                                UniqueId = uniqueId,
                                                ContentHash = contentHash,
                                                TextType = "chart_axis",
                                                ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_catAxis{catAxisIndex}",
                                                TranslationStatus = "pending"
                                            });
                                        }
                                    }
                                }
                            }
                        }
                        catAxisIndex++;
                    }

                    // 提取系列名称和标题
                    var barChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.BarChart>();
                    var lineChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.LineChart>();
                    var pieChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PieChart>();
                    var areaChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.AreaChart>();
                    var scatterChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>();
                    
                    var allSeries = new List<OpenXmlElement>();
                    
                    if (barChart != null)
                    {
                        allSeries.AddRange(barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>());
                    }
                    if (lineChart != null)
                    {
                        allSeries.AddRange(lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>());
                    }
                    if (pieChart != null)
                    {
                        allSeries.AddRange(pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>());
                    }
                    if (areaChart != null)
                    {
                        allSeries.AddRange(areaChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.AreaChartSeries>());
                    }
                    if (scatterChart != null)
                    {
                        allSeries.AddRange(scatterChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.ScatterChartSeries>());
                    }
                    
                    int seriesIndex = 0;
                    
                    foreach (var series in allSeries)
                    {
                        var seriesTitle = series.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                        if (seriesTitle != null)
                        {
                            var chartText = seriesTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                            if (chartText != null)
                            {
                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                if (richText != null)
                                {
                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                    if (textPara != null)
                                    {
                                        var text = GetElementText(textPara);
                                        if (!string.IsNullOrEmpty(text))
                                        {
                                            var uniqueId = GenerateUniqueId(slideNumber, "chart", $"{chartId}_series{seriesIndex}", chartIndex);
                                            var contentHash = ComputeContentHash(text);
                                            
                                            entries.Add(new PPTTextEntry
                                            {
                                                OriginalText = text,
                                                SlideNumber = slideNumber,
                                                Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}, 系列: {seriesIndex}",
                                                ElementType = "chart",
                                                ElementId = $"{chartId}_series{seriesIndex}",
                                                UniqueId = uniqueId,
                                                ContentHash = contentHash,
                                                TextType = "chart_series",
                                                ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_series{seriesIndex}",
                                                TranslationStatus = "pending"
                                            });
                                        }
                                    }
                                }
                            }
                        }

                        // 提取数据标签
                        var dataLabels = series.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
                        if (dataLabels != null)
                        {
                            var labelValues = dataLabels.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataLabel>();
                            int labelIndex = 0;
                            
                            foreach (var label in labelValues)
                            {
                                var chartText = label.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                if (chartText != null)
                                {
                                    var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                    if (richText != null)
                                    {
                                        var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                        if (textPara != null)
                                        {
                                            var text = GetElementText(textPara);
                                            if (!string.IsNullOrEmpty(text))
                                            {
                                                var uniqueId = GenerateUniqueId(slideNumber, "chart", $"{chartId}_series{seriesIndex}_label{labelIndex}", chartIndex);
                                                var contentHash = ComputeContentHash(text);
                                                
                                                entries.Add(new PPTTextEntry
                                                {
                                                    OriginalText = text,
                                                    SlideNumber = slideNumber,
                                                    Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}, 系列: {seriesIndex}, 标签: {labelIndex}",
                                                    ElementType = "chart",
                                                    ElementId = $"{chartId}_series{seriesIndex}_label{labelIndex}",
                                                    UniqueId = uniqueId,
                                                    ContentHash = contentHash,
                                                    TextType = "chart_datalabel",
                                                    ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_series{seriesIndex}_label{labelIndex}",
                                                    TranslationStatus = "pending"
                                                });
                                            }
                                        }
                                    }
                                }
                                labelIndex++;
                            }
                        }
                        
                        seriesIndex++;
                    }
                }

                // 提取图例文本
                var legend = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Legend>();
                if (legend != null)
                {
                    var legendEntries = legend.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LegendEntry>();
                    int legendIndex = 0;
                    
                    foreach (var legendEntry in legendEntries)
                    {
                        var chartText = legendEntry.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                        if (chartText != null)
                        {
                            var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                            if (richText != null)
                            {
                                var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                if (textPara != null)
                                {
                                    var text = GetElementText(textPara);
                                    if (!string.IsNullOrEmpty(text))
                                    {
                                        var uniqueId = GenerateUniqueId(slideNumber, "chart", $"{chartId}_legend{legendIndex}", chartIndex);
                                        var contentHash = ComputeContentHash(text);
                                        
                                        entries.Add(new PPTTextEntry
                                        {
                                            OriginalText = text,
                                            SlideNumber = slideNumber,
                                            Position = $"幻灯片 {slideNumber}, 图表, 图表ID: {chartId}, 图例: {legendIndex}",
                                            ElementType = "chart",
                                            ElementId = $"{chartId}_legend{legendIndex}",
                                            UniqueId = uniqueId,
                                            ContentHash = contentHash,
                                            TextType = "chart_legend",
                                            ContextInfo = $"幻灯片{slideNumber}_chart_{chartId}_legend{legendIndex}",
                                            TranslationStatus = "pending"
                                        });
                                    }
                                }
                            }
                        }
                        legendIndex++;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"提取图表文本失败: {ex.Message}");
            }
            
            return entries;
        }

        /// <summary>
        /// 翻译文本条目
        /// </summary>
        private List<PPTTextEntry> TranslateTextEntries(List<PPTTextEntry> entries, Dictionary<string, string> terminology, int slideNumber, string elementType)
        {
            var translatedEntries = new List<PPTTextEntry>();
            
            try
            {
                foreach (var entry in entries)
                {
                    var text = entry.OriginalText;
                    
                    // 应用术语预处理
                    if (_useTerminology && _preprocessTerms && terminology != null && terminology.Any())
                    {
                        text = _termExtractor.ReplaceTermsInText(text, terminology, _isCnToForeign);
                    }
                    
                    // 翻译文本
                    var promptText = _useImmersiveStyle && _immersivePrompts.ContainsKey(entry.TextType) 
                        ? _immersivePrompts[entry.TextType] 
                        : (_pptPrompts.ContainsKey(entry.TextType) ? _pptPrompts[entry.TextType] : "");
                    var translatedText = _translationService.TranslateTextAsync(text, terminology, _sourceLang, _targetLang, promptText, entry.OriginalText).GetAwaiter().GetResult();
                    
                    var translatedEntry = new PPTTextEntry
                    {
                        OriginalText = entry.OriginalText,
                        ProcessedText = text,
                        TranslatedText = translatedText,
                        SlideNumber = entry.SlideNumber,
                        Position = entry.Position,
                        ElementType = entry.ElementType,
                        ElementId = entry.ElementId,
                        ParagraphIndex = entry.ParagraphIndex,
                        RowIndex = entry.RowIndex,
                        ColumnIndex = entry.ColumnIndex,
                        UniqueId = entry.UniqueId,
                        ContentHash = entry.ContentHash,
                        ParentElementId = entry.ParentElementId,
                        TextType = entry.TextType,
                        FormatProperties = entry.FormatProperties,
                        RunProperties = entry.RunProperties,
                        IsPreserved = entry.IsPreserved,
                        TranslationStatus = "completed",
                        ExtractionTime = entry.ExtractionTime,
                        ContextInfo = entry.ContextInfo
                    };
                    
                    translatedEntries.Add(translatedEntry);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"翻译文本条目失败: {ex.Message}");
            }
            
            return translatedEntries;
        }

        /// <summary>
        /// 将翻译文本写入图表
        /// </summary>
        private void WriteChartText(ChartPart chartPart, List<PPTTextEntry> translatedEntries, int slideNumber)
        {
            try
            {
                var chart = chartPart?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
                if (chart == null) return;

                foreach (var entry in translatedEntries)
                {
                    try
                    {
                        if (entry.TextType == "chart_title")
                        {
                            var title = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                            if (title != null)
                            {
                                var chartText = title.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                if (chartText != null)
                                {
                                    var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                    if (richText != null)
                                    {
                                        var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                        if (textPara != null)
                                        {
                                            var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                            foreach (var run in runs)
                                            {
                                                var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                if (textElement != null)
                                                {
                                                    textElement.Text = entry.TranslatedText;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (entry.TextType == "chart_axis")
                        {
                            var plotArea = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
                            if (plotArea != null)
                            {
                                // 处理数值轴
                                if (entry.ElementId.Contains("valueAxis"))
                                {
                                    var valueAxes = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>();
                                    foreach (var axis in valueAxes)
                                    {
                                        var axisTitle = axis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                                        if (axisTitle != null)
                                        {
                                            var chartText = axisTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                            if (chartText != null)
                                            {
                                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                                if (richText != null)
                                                {
                                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                                    if (textPara != null)
                                                    {
                                                        var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                                        foreach (var run in runs)
                                                        {
                                                            var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                            if (textElement != null)
                                                            {
                                                                textElement.Text = entry.TranslatedText;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                // 处理类别轴
                                else if (entry.ElementId.Contains("catAxis"))
                                {
                                    var categoryAxes = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>();
                                    foreach (var catAxis in categoryAxes)
                                    {
                                        var axisTitle = catAxis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                                        if (axisTitle != null)
                                        {
                                            var chartText = axisTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                            if (chartText != null)
                                            {
                                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                                if (richText != null)
                                                {
                                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                                    if (textPara != null)
                                                    {
                                                        var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                                        foreach (var run in runs)
                                                        {
                                                            var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                            if (textElement != null)
                                                            {
                                                                textElement.Text = entry.TranslatedText;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (entry.TextType == "chart_series")
                        {
                            var plotArea = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
                            if (plotArea != null)
                            {
                                var barChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.BarChart>();
                                var lineChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.LineChart>();
                                var pieChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PieChart>();
                                var areaChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.AreaChart>();
                                var scatterChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>();
                                
                                var allSeries = new List<OpenXmlElement>();
                                
                                if (barChart != null) allSeries.AddRange(barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>());
                                if (lineChart != null) allSeries.AddRange(lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>());
                                if (pieChart != null) allSeries.AddRange(pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>());
                                if (areaChart != null) allSeries.AddRange(areaChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.AreaChartSeries>());
                                if (scatterChart != null) allSeries.AddRange(scatterChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.ScatterChartSeries>());
                                
                                foreach (var series in allSeries)
                                {
                                    var seriesTitle = series.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Title>();
                                    if (seriesTitle != null)
                                    {
                                        var chartText = seriesTitle.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                        if (chartText != null)
                                        {
                                            var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                            if (richText != null)
                                            {
                                                var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                                if (textPara != null)
                                                {
                                                    var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                                    foreach (var run in runs)
                                                    {
                                                        var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                        if (textElement != null)
                                                        {
                                                            textElement.Text = entry.TranslatedText;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (entry.TextType == "chart_datalabel")
                        {
                            var plotArea = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>();
                            if (plotArea != null)
                            {
                                var barChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.BarChart>();
                                var lineChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.LineChart>();
                                var pieChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.PieChart>();
                                var areaChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.AreaChart>();
                                var scatterChart = plotArea.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>();
                                
                                var allSeries = new List<OpenXmlElement>();
                                
                                if (barChart != null) allSeries.AddRange(barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>());
                                if (lineChart != null) allSeries.AddRange(lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>());
                                if (pieChart != null) allSeries.AddRange(pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>());
                                if (areaChart != null) allSeries.AddRange(areaChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.AreaChartSeries>());
                                if (scatterChart != null) allSeries.AddRange(scatterChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.ScatterChartSeries>());
                                
                                foreach (var series in allSeries)
                                {
                                    var dataLabels = series.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
                                    if (dataLabels != null)
                                    {
                                        var labelValues = dataLabels.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataLabel>();
                                        foreach (var label in labelValues)
                                        {
                                            var chartText = label.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                            if (chartText != null)
                                            {
                                                var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                                if (richText != null)
                                                {
                                                    var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                                    if (textPara != null)
                                                    {
                                                        var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                                        foreach (var run in runs)
                                                        {
                                                            var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                            if (textElement != null)
                                                            {
                                                                textElement.Text = entry.TranslatedText;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (entry.TextType == "chart_legend")
                        {
                            var legend = chart.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Legend>();
                            if (legend != null)
                            {
                                var legendEntries = legend.Descendants<DocumentFormat.OpenXml.Drawing.Charts.LegendEntry>();
                                foreach (var legendEntry in legendEntries)
                                {
                                    var chartText = legendEntry.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartText>().FirstOrDefault();
                                    if (chartText != null)
                                    {
                                        var richText = chartText.Descendants<DocumentFormat.OpenXml.Drawing.Charts.RichText>().FirstOrDefault();
                                        if (richText != null)
                                        {
                                            var textPara = richText.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault();
                                            if (textPara != null)
                                            {
                                                var runs = textPara.Elements<DocumentFormat.OpenXml.Drawing.Run>().ToList();
                                                foreach (var run in runs)
                                                {
                                                    var textElement = run.GetFirstChild<DocumentFormat.OpenXml.Drawing.Text>();
                                                    if (textElement != null)
                                                    {
                                                        textElement.Text = entry.TranslatedText;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        
                        entry.TranslationStatus = "completed";
                        entry.IsPreserved = true;
                        
                        _logger.LogDebug($"已写入图表文本: {entry.UniqueId}");
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning($"写入单个图表文本失败: {ex.Message}");
                    }
                }
                
                if (chartPart?.ChartSpace != null)
                {
                    chartPart.ChartSpace.Save();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError($"写入图表文本失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证并修复演示文稿结构
        /// </summary>
        private void ValidateAndFixPresentationStructure(PresentationDocument presentation)
        {
            try
            {
                var presentationPart = presentation.PresentationPart;
                if (presentationPart?.Presentation == null) return;

                // 验证每张幻灯片
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    ValidateAndFixSlideStructure(slidePart);
                }

                _logger.LogDebug("PPT文档结构验证和修复完成");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"PPT文档结构验证失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证并修复幻灯片结构
        /// </summary>
        private void ValidateAndFixSlideStructure(SlidePart slidePart)
        {
            try
            {
                if (slidePart?.Slide == null) return;

                // 验证文本形状
                foreach (var shape in slidePart.Slide.Descendants<Shape>())
                {
                    ValidateAndFixTextShape(shape);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"幻灯片结构验证失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 验证并修复文本形状
        /// </summary>
        private void ValidateAndFixTextShape(Shape shape)
        {
            try
            {
                if (shape?.TextBody == null) return;

                // 确保每个段落都有有效的内容
                var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
                foreach (var paragraph in paragraphs)
                {
                    // 如果段落为空，添加一个空的Run
                    if (!paragraph.Elements<A.Run>().Any())
                    {
                        var emptyRun = new A.Run();
                        emptyRun.AppendChild(new A.Text { Text = "" });
                        paragraph.AppendChild(emptyRun);
                    }

                    // 验证每个Run都有Text元素
                    foreach (var run in paragraph.Elements<A.Run>())
                    {
                        if (!run.Elements<A.Text>().Any())
                        {
                            run.AppendChild(new A.Text { Text = "" });
                        }

                        // 确保RunProperties存在且有效
                        if (run.RunProperties == null)
                        {
                            run.RunProperties = new A.RunProperties();
                        }
                    }

                    // 确保段落属性存在
                    if (paragraph.ParagraphProperties == null)
                    {
                        paragraph.ParagraphProperties = new A.ParagraphProperties();
                    }
                }

                // 如果TextBody没有段落，添加一个空段落
                if (!paragraphs.Any())
                {
                    var emptyParagraph = new A.Paragraph();
                    var emptyRun = new A.Run();
                    emptyRun.AppendChild(new A.Text { Text = "" });
                    emptyParagraph.AppendChild(emptyRun);
                    shape.TextBody.AppendChild(emptyParagraph);
                }

                // 确保TextBody属性存在
                if (shape.TextBody.BodyProperties == null)
                {
                    shape.TextBody.BodyProperties = new A.BodyProperties();
                }

                if (shape.TextBody.ListStyle == null)
                {
                    shape.TextBody.ListStyle = new A.ListStyle();
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"文本形状验证失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复演示文稿结构的高级方法
        /// </summary>
        private void RepairPresentationStructure(PresentationDocument presentation)
        {
            try
            {
                _logger.LogInformation("开始高级文档结构修复...");

                var presentationPart = presentation.PresentationPart;
                if (presentationPart?.Presentation == null) return;

                // 修复幻灯片关系
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    RepairSlideStructure(slidePart);
                }

                // 确保演示文稿有有效的幻灯片ID列表
                var slideIdList = presentationPart.Presentation.SlideIdList;
                if (slideIdList == null)
                {
                    slideIdList = new SlideIdList();
                    presentationPart.Presentation.AppendChild(slideIdList);
                }

                // 验证幻灯片ID的一致性
                var slideIds = slideIdList.Elements<SlideId>().ToList();
                var slideParts = presentationPart.SlideParts.ToList();

                if (slideIds.Count != slideParts.Count)
                {
                    _logger.LogWarning($"幻灯片ID数量({slideIds.Count})与幻灯片部件数量({slideParts.Count})不匹配，正在修复...");

                    // 重建幻灯片ID列表
                    slideIdList.RemoveAllChildren();
                    uint slideId = 256;

                    foreach (var slidePart in slideParts)
                    {
                        var newSlideId = new SlideId
                        {
                            Id = slideId++,
                            RelationshipId = presentationPart.GetIdOfPart(slidePart)
                        };
                        slideIdList.AppendChild(newSlideId);
                    }
                }

                _logger.LogInformation("高级文档结构修复完成");
            }
            catch (Exception ex)
            {
                _logger.LogError($"高级文档结构修复失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复幻灯片结构
        /// </summary>
        private void RepairSlideStructure(SlidePart slidePart)
        {
            try
            {
                if (slidePart?.Slide == null) return;

                // 确保幻灯片有CommonSlideData
                if (slidePart.Slide.CommonSlideData == null)
                {
                    slidePart.Slide.CommonSlideData = new CommonSlideData();
                }

                // 确保有ShapeTree
                if (slidePart.Slide.CommonSlideData.ShapeTree == null)
                {
                    slidePart.Slide.CommonSlideData.ShapeTree = new ShapeTree();
                }

                var shapeTree = slidePart.Slide.CommonSlideData.ShapeTree;

                // 确保ShapeTree有必要的属性
                if (shapeTree.NonVisualGroupShapeProperties == null)
                {
                    shapeTree.NonVisualGroupShapeProperties = new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties();
                }

                if (shapeTree.GroupShapeProperties == null)
                {
                    shapeTree.GroupShapeProperties = new DocumentFormat.OpenXml.Presentation.GroupShapeProperties();
                }

                // 修复所有形状
                foreach (var shape in shapeTree.Elements<Shape>())
                {
                    ValidateAndFixTextShape(shape);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"修复幻灯片结构失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 深度结构修复：修复可能导致文档损坏的问题
        /// </summary>
        private void DeepStructureRepair(PresentationDocument presentation)
        {
            try
            {
                _logger.LogInformation("开始深度结构修复...");

                var presentationPart = presentation.PresentationPart;
                if (presentationPart?.Presentation == null) return;

                // 1. 修复幻灯片大小
                RepairSlideSize(presentationPart);

                // 2. 修复文本运行属性
                RepairTextRunProperties(presentationPart);

                // 3. 清理空的文本元素
                CleanupEmptyTextElements(presentationPart);

                // 4. 修复段落结构
                RepairParagraphStructure(presentationPart);

                _logger.LogInformation("深度结构修复完成");
            }
            catch (Exception ex)
            {
                _logger.LogError($"深度结构修复失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复幻灯片大小
        /// </summary>
        private void RepairSlideSize(PresentationPart presentationPart)
        {
            try
            {
                var presentation = presentationPart.Presentation;
                if (presentation.SlideSize == null)
                {
                    presentation.SlideSize = new SlideSize()
                    {
                        Cx = 9144000, // 标准16:9宽度
                        Cy = 6858000  // 标准16:9高度
                    };
                    _logger.LogDebug("已修复幻灯片大小");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"修复幻灯片大小失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复文本运行属性
        /// </summary>
        private void RepairTextRunProperties(PresentationPart presentationPart)
        {
            try
            {
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    foreach (var run in slidePart.Slide.Descendants<A.Run>())
                    {
                        if (run.RunProperties == null)
                        {
                            run.RunProperties = new A.RunProperties();
                        }

                        // 确保字号有效
                        if (run.RunProperties.FontSize != null && run.RunProperties.FontSize < 100)
                        {
                            run.RunProperties.FontSize = 1200; // 默认12pt
                        }
                    }
                }
                _logger.LogDebug("已修复文本运行属性");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"修复文本运行属性失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 清理空的文本元素
        /// </summary>
        private void CleanupEmptyTextElements(PresentationPart presentationPart)
        {
            try
            {
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    var emptyRuns = slidePart.Slide.Descendants<A.Run>()
                        .Where(r => !r.Elements<A.Text>().Any() ||
                                   r.Elements<A.Text>().All(t => string.IsNullOrWhiteSpace(t.Text)))
                        .ToList();

                    foreach (var emptyRun in emptyRuns)
                    {
                        // 为空的Run添加空文本，而不是删除
                        if (!emptyRun.Elements<A.Text>().Any())
                        {
                            emptyRun.AppendChild(new A.Text { Text = "" });
                        }
                    }
                }
                _logger.LogDebug("已清理空的文本元素");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"清理空的文本元素失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 修复段落结构
        /// </summary>
        private void RepairParagraphStructure(PresentationPart presentationPart)
        {
            try
            {
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>())
                    {
                        // 确保段落有至少一个Run
                        if (!paragraph.Elements<A.Run>().Any())
                        {
                            var emptyRun = new A.Run();
                            emptyRun.AppendChild(new A.Text { Text = "" });
                            paragraph.AppendChild(emptyRun);
                        }

                        // 确保段落属性存在
                        if (paragraph.ParagraphProperties == null)
                        {
                            paragraph.ParagraphProperties = new A.ParagraphProperties();
                        }
                    }
                }
                _logger.LogDebug("已修复段落结构");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"修复段落结构失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理双语对照模式：复制每一页，原页删除译文，复制页删除原文（重载方法）
        /// </summary>
        private void ProcessBilingualModeAsync(PresentationPart presentationPart, List<TranslationResult> translationResults)
        {
            ProcessBilingualModeAsync(presentationPart, (List<PPTTextEntry>?)null);
        }

        /// <summary>
        /// 处理双语对照模式：复制每一页，原页删除译文，复制页删除原文
        /// </summary>
        private void ProcessBilingualModeAsync(PresentationPart presentationPart, List<PPTTextEntry>? textEntries)
        {
            try
            {
                _logger.LogInformation("双语对照模式：正在处理幻灯片...");
                UpdateProgress(0.80, "双语对照模式：正在复制和分离原文译文...");

                var slides = presentationPart.SlideParts.ToList();
                var slidesToProcess = new List<SlidePart>(slides);

                foreach (var originalSlide in slidesToProcess)
                {
                    // 复制幻灯片
                    var clonedSlide = CloneSlide(presentationPart, originalSlide);

                    // 原页：删除译文段落
                    RemoveTranslatedParagraphs(originalSlide, textEntries);
                    if (_translateSmartArt)
                    {
                        RemoveTranslatedTextFromSmartArt(originalSlide);
                    }

                    // 复制页：删除原文段落
                    RemoveOriginalParagraphs(clonedSlide, textEntries);
                    if (_translateSmartArt)
                    {
                        RemoveOriginalTextFromSmartArt(clonedSlide);
                    }

                    _logger.LogDebug($"已处理双语对照幻灯片");
                }

                _logger.LogInformation($"双语对照模式处理完成，共处理 {slidesToProcess.Count} 张幻灯片");
                UpdateProgress(0.85, "双语对照模式处理完成");
            }
            catch (Exception ex)
            {
                _logger.LogError($"双语对照模式处理失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 处理仅译文模式：删除原页，只保留译文页（重载方法）
        /// </summary>
        private void ProcessTranslationOnlyModeAsync(PresentationPart presentationPart, List<TranslationResult> translationResults)
        {
            ProcessTranslationOnlyModeAsync(presentationPart, (List<PPTTextEntry>?)null);
        }

        /// <summary>
        /// 处理仅译文模式：删除原页，只保留译文页
        /// </summary>
        private void ProcessTranslationOnlyModeAsync(PresentationPart presentationPart, List<PPTTextEntry>? textEntries)
        {
            try
            {
                _logger.LogInformation("仅译文模式：正在处理幻灯片...");
                UpdateProgress(0.80, "仅译文模式：正在删除原文段落...");

                var slides = presentationPart.SlideParts.ToList();
                foreach (var slide in slides)
                {
                    // 删除原文段落，只保留译文段落
                    RemoveOriginalParagraphs(slide, textEntries);
                    if (_translateSmartArt)
                    {
                        RemoveOriginalTextFromSmartArt(slide);
                    }
                }

                _logger.LogInformation($"仅译文模式处理完成，共处理 {slides.Count} 张幻灯片");
                UpdateProgress(0.85, "仅译文模式处理完成");
            }
            catch (Exception ex)
            {
                _logger.LogError($"仅译文模式处理失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 删除译文段落，只保留原文段落（重载方法）
        /// </summary>
        private void RemoveTranslatedParagraphs(SlidePart slidePart, List<TranslationResult> translationResults)
        {
            RemoveTranslatedParagraphs(slidePart, (List<PPTTextEntry>?)null);
        }

        /// <summary>
        /// 删除译文段落，只保留原文段落
        /// </summary>
        private void RemoveTranslatedParagraphs(SlidePart slidePart, List<PPTTextEntry>? textEntries)
        {
            try
            {
                var slide = slidePart?.Slide;
                var shapes = slide?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();

                foreach (var shape in shapes)
                {
                    if (shape.TextBody != null)
                    {
                        var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();

                        foreach (var paragraph in paragraphs)
                        {
                            var runs = paragraph.Elements<A.Run>().ToList();

                            // 检查是否包含分隔符##-----##
                            var separatorFound = false;
                            var separatorIndex = -1;

                            for (int i = 0; i < runs.Count; i++)
                            {
                                var run = runs[i];
                                var textElement = run.Elements<A.Text>().FirstOrDefault();
                                
                                if (textElement != null && textElement.Text.Contains("##-----##"))
                                {
                                    separatorFound = true;
                                    separatorIndex = i;
                                    break;
                                }
                            }

                            if (separatorFound && separatorIndex >= 0)
                            {
                                // 找到分隔符，删除译文部分（分隔符之后的所有内容）
                                var runsToRemove = new List<A.Run>();
                                for (int j = separatorIndex; j < runs.Count; j++)
                                {
                                    runsToRemove.Add(runs[j]);
                                }
                                
                                // 删除译文部分
                                foreach (var runToRemove in runsToRemove)
                                {
                                    runToRemove.Remove();
                                }
                                
                                _logger.LogDebug($"已删除译文，保留原文");
                            }
                            else
                            {
                                // 没有找到分隔符，保留所有内容（原文）
                                _logger.LogDebug($"未找到分隔符，保留所有内容");
                            }
                        }
                    }

                    // 处理表格
                    var tables = shape.Descendants().Where(e => e.LocalName == "tbl");
                    foreach (var table in tables)
                    {
                        var rows = table.Elements().Where(e => e.LocalName == "tr").ToList();
                        foreach (var row in rows)
                        {
                            var cells = row.Elements().Where(e => e.LocalName == "tc").ToList();
                            foreach (var cell in cells)
                            {
                                var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                                if (textBody != null)
                                {
                                    var paragraphs = textBody.Elements<A.Paragraph>().ToList();
                                    foreach (var paragraph in paragraphs)
                                    {
                                        var runs = paragraph.Elements<A.Run>().ToList();

                                        // 检查是否包含分隔符##-----##
                                        var separatorFound = false;
                                        var separatorIndex = -1;

                                        for (int i = 0; i < runs.Count; i++)
                                        {
                                            var run = runs[i];
                                            var textElement = run.Elements<A.Text>().FirstOrDefault();
                                            
                                            if (textElement != null && textElement.Text.Contains("##-----##"))
                                            {
                                                separatorFound = true;
                                                separatorIndex = i;
                                                break;
                                            }
                                        }

                                        if (separatorFound && separatorIndex >= 0)
                                        {
                                            // 找到分隔符，删除译文部分（分隔符之后的所有内容）
                                            var runsToRemove = new List<A.Run>();
                                            for (int j = separatorIndex; j < runs.Count; j++)
                                            {
                                                runsToRemove.Add(runs[j]);
                                            }
                                            
                                            // 删除译文部分
                                            foreach (var runToRemove in runsToRemove)
                                            {
                                                runToRemove.Remove();
                                            }
                                            
                                            _logger.LogDebug($"表格单元格已删除译文，保留原文");
                                        }
                                        else
                                        {
                                            // 没有找到分隔符，保留所有内容（原文）
                                            _logger.LogDebug($"表格单元格未找到分隔符，保留所有内容");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"删除译文段落失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 删除原文段落，只保留译文段落（重载方法）
        /// </summary>
        private void RemoveOriginalParagraphs(SlidePart slidePart, List<TranslationResult> translationResults)
        {
            RemoveOriginalParagraphs(slidePart, (List<PPTTextEntry>?)null);
        }

        /// <summary>
        /// 删除原文段落，只保留译文段落
        /// </summary>
        private void RemoveOriginalParagraphs(SlidePart slidePart, List<PPTTextEntry>? textEntries)
        {
            try
            {
                var slide = slidePart?.Slide;
                var shapes = slide?.Descendants<Shape>() ?? Enumerable.Empty<Shape>();

                foreach (var shape in shapes)
                {
                    if (shape.TextBody != null)
                    {
                        var paragraphs = shape.TextBody.Elements<A.Paragraph>().ToList();

                        foreach (var paragraph in paragraphs)
                        {
                            var runs = paragraph.Elements<A.Run>().ToList();

                            // 检查是否包含分隔符##-----##
                            var separatorFound = false;
                            var separatorIndex = -1;

                            for (int i = 0; i < runs.Count; i++)
                            {
                                var run = runs[i];
                                var textElement = run.Elements<A.Text>().FirstOrDefault();
                                
                                if (textElement != null && textElement.Text.Contains("##-----##"))
                                {
                                    separatorFound = true;
                                    separatorIndex = i;
                                    break;
                                }
                            }

                            if (separatorFound && separatorIndex >= 0)
                            {
                                // 找到分隔符，删除原文部分（分隔符之前的所有内容）
                                var runsToRemove = new List<A.Run>();
                                for (int j = 0; j <= separatorIndex; j++)
                                {
                                    runsToRemove.Add(runs[j]);
                                }
                                
                                // 删除原文部分和分隔符
                                foreach (var runToRemove in runsToRemove)
                                {
                                    runToRemove.Remove();
                                }
                                
                                _logger.LogDebug($"已删除原文，保留译文");
                            }
                            else
                            {
                                // 没有找到分隔符，保留所有内容（可能只有原文或只有译文）
                                _logger.LogDebug($"未找到分隔符，保留所有内容");
                            }
                        }
                    }

                    // 处理表格
                    var tables = shape.Descendants().Where(e => e.LocalName == "tbl");
                    foreach (var table in tables)
                    {
                        var rows = table.Elements().Where(e => e.LocalName == "tr").ToList();
                        foreach (var row in rows)
                        {
                            var cells = row.Elements().Where(e => e.LocalName == "tc").ToList();
                            foreach (var cell in cells)
                            {
                                var textBody = cell.Elements().FirstOrDefault(e => e.LocalName == "txBody");
                                if (textBody != null)
                                {
                                    var paragraphs = textBody.Elements<A.Paragraph>().ToList();
                                    foreach (var paragraph in paragraphs)
                                    {
                                        var runs = paragraph.Elements<A.Run>().ToList();

                                        // 检查是否包含分隔符##-----##
                                        var separatorFound = false;
                                        var separatorIndex = -1;

                                        for (int i = 0; i < runs.Count; i++)
                                        {
                                            var run = runs[i];
                                            var textElement = run.Elements<A.Text>().FirstOrDefault();
                                            
                                            if (textElement != null && textElement.Text.Contains("##-----##"))
                                            {
                                                separatorFound = true;
                                                separatorIndex = i;
                                                break;
                                            }
                                        }

                                        if (separatorFound && separatorIndex >= 0)
                                        {
                                            // 找到分隔符，删除原文部分（分隔符之前的所有内容）
                                            var runsToRemove = new List<A.Run>();
                                            for (int j = 0; j <= separatorIndex; j++)
                                            {
                                                runsToRemove.Add(runs[j]);
                                            }
                                            
                                            // 删除原文部分和分隔符
                                            foreach (var runToRemove in runsToRemove)
                                            {
                                                runToRemove.Remove();
                                            }
                                            
                                            _logger.LogDebug($"表格单元格已删除原文，保留译文");
                                        }
                                        else
                                        {
                                            // 没有找到分隔符，保留所有内容（可能只有原文或只有译文）
                                            _logger.LogDebug($"表格单元格未找到分隔符，保留所有内容");
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"删除原文段落失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 复制幻灯片，并将其插入到原幻灯片之后
        /// </summary>
        private SlidePart CloneSlide(PresentationPart presentationPart, SlidePart sourceSlidePart)
        {
            try
            {
                // 创建新的幻灯片部分
                var newSlidePart = presentationPart.AddNewPart<SlidePart>();

                // 深度复制源幻灯片的内容
                newSlidePart.Slide = (Slide)sourceSlidePart.Slide.CloneNode(true);

                // 复制幻灯片布局引用
                if (sourceSlidePart.SlideLayoutPart != null)
                {
                    newSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
                }

                // 复制备注部分（如果存在）
                if (sourceSlidePart.NotesSlidePart != null)
                {
                    var newNotesSlidePart = newSlidePart.AddNewPart<NotesSlidePart>();
                    newNotesSlidePart.NotesSlide = (NotesSlide)sourceSlidePart.NotesSlidePart.NotesSlide.CloneNode(true);

                    // 复制备注的主布局引用
                    if (sourceSlidePart.NotesSlidePart.NotesMasterPart != null)
                    {
                        newNotesSlidePart.AddPart(sourceSlidePart.NotesSlidePart.NotesMasterPart);
                    }
                }

                // 复制所有图像和其他媒体部分
                var imagePartMapping = new Dictionary<string, string>();
                foreach (var imagePart in sourceSlidePart.ImageParts)
                {
                    var newImagePart = newSlidePart.AddImagePart(imagePart.ContentType);
                    using (var stream = imagePart.GetStream())
                    {
                        newImagePart.FeedData(stream);
                    }

                    var oldImageRelId = sourceSlidePart.GetIdOfPart(imagePart);
                    var newImageRelId = newSlidePart.GetIdOfPart(newImagePart);
                    if (oldImageRelId != null && newImageRelId != null)
                    {
                        imagePartMapping[oldImageRelId] = newImageRelId;
                    }
                }

                // 更新幻灯片中图片引用的RelationshipId
                UpdateImageRelationships(newSlidePart.Slide, imagePartMapping);

                // 复制所有图表部分
                foreach (var chartPart in sourceSlidePart.ChartParts)
                {
                    var newChartPart = newSlidePart.AddNewPart<ChartPart>();
                    newChartPart.ChartSpace = (ChartSpace)chartPart.ChartSpace.CloneNode(true);
                }

                // 将新幻灯片添加到演示文稿的幻灯片列表中，插入到原幻灯片之后
                var slideIdList = presentationPart.Presentation.SlideIdList;
                if (slideIdList == null)
                {
                    slideIdList = new SlideIdList();
                    presentationPart.Presentation.AppendChild(slideIdList);
                }

                // 找到原幻灯片在列表中的位置
                var sourceSlideId = slideIdList.Elements<SlideId>()
                    .FirstOrDefault(s => s.RelationshipId == presentationPart.GetIdOfPart(sourceSlidePart));

                // 生成新的幻灯片ID
                var slideIdElements = slideIdList.Elements<SlideId>().ToList();
                var maxSlideId = slideIdElements.Any() ? slideIdElements.Max(s => s.Id?.Value ?? 255) : 255;
                var newSlideId = new SlideId();
                newSlideId.Id = (uint)(maxSlideId + 1);
                var relationshipId = presentationPart.GetIdOfPart(newSlidePart);
                if (relationshipId != null)
                {
                    newSlideId.RelationshipId = relationshipId;
                }

                if (sourceSlideId != null)
                {
                    // 在原幻灯片之后插入新幻灯片
                    sourceSlideId.InsertAfterSelf(newSlideId);
                    _logger.LogDebug($"已在原幻灯片之后插入新幻灯片，新幻灯片ID: {newSlideId.Id}");
                }
                else
                {
                    // 如果找不到原幻灯片，添加到列表末尾
                    slideIdList.AppendChild(newSlideId);
                    _logger.LogDebug($"已将新幻灯片添加到列表末尾，新幻灯片ID: {newSlideId.Id}");
                }

                return newSlidePart;
            }
            catch (Exception ex)
            {
                _logger.LogError($"复制幻灯片失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 更新幻灯片中图片引用的RelationshipId
        /// </summary>
        private void UpdateImageRelationships(Slide slide, Dictionary<string, string> imagePartMapping)
        {
            try
            {
                var pictureElements = slide.Descendants<P.Picture>();
                foreach (var picture in pictureElements)
                {
                    var blip = picture.Descendants<A.Blip>().FirstOrDefault();
                    if (blip != null && blip.Embed != null && blip.Embed.Value != null && imagePartMapping.TryGetValue(blip.Embed.Value, out var newRelId))
                    {
                        blip.Embed.Value = newRelId;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"更新图片引用失败: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 翻译结果类
    /// </summary>
    public class TranslationResult
    {
        public string OriginalText { get; set; } = string.Empty;
        public string TranslatedText { get; set; } = string.Empty;
        public string Position { get; set; } = string.Empty;
        public string ElementType { get; set; } = string.Empty;
    }

    /// <summary>
    /// PPT文本条目，用于记录文本及其位置信息
    /// </summary>
    public class PPTTextEntry
    {
        public string OriginalText { get; set; } = string.Empty;
        public string ProcessedText { get; set; } = string.Empty;
        public string TranslatedText { get; set; } = string.Empty;
        public int SlideNumber { get; set; }
        public string Position { get; set; } = string.Empty;
        public string ElementType { get; set; } = string.Empty;
        public string ElementId { get; set; } = string.Empty;
        public int ParagraphIndex { get; set; } = -1;
        public int RowIndex { get; set; } = -1;
        public int ColumnIndex { get; set; } = -1;

        public string UniqueId { get; set; } = string.Empty;
        public string ContentHash { get; set; } = string.Empty;
        public string ParentElementId { get; set; } = string.Empty;
        public string TextType { get; set; } = string.Empty;
        public Dictionary<string, string> FormatProperties { get; set; } = new Dictionary<string, string>();
        public List<string> RunProperties { get; set; } = new List<string>();
        public bool IsPreserved { get; set; } = false;
        public string TranslationStatus { get; set; } = "pending";
        public DateTime ExtractionTime { get; set; } = DateTime.Now;
        public string ContextInfo { get; set; } = string.Empty;
    }
}
