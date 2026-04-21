using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;

namespace DocumentTranslator.Services.Translation.Tests
{
    /// <summary>
    /// PPT翻译功能测试类
    /// </summary>
    public class PPTTranslationTests
    {
        private readonly ILogger<PPTTranslationTests> _logger;
        private readonly TranslationService _translationService;
        private readonly PPTProcessor _pptProcessor;

        public PPTTranslationTests()
        {
            _logger = NullLogger<PPTTranslationTests>.Instance;
            
            // 创建模拟翻译服务
            var translationLogger = NullLogger<TranslationService>.Instance;
            _translationService = new MockTranslationService(translationLogger);
            
            // 创建PPT处理器，使用简单的logger
            var pptLogger = NullLogger<PPTProcessor>.Instance;
            _pptProcessor = new PPTProcessor(_translationService, pptLogger);
        }

        /// <summary>
        /// 运行所有测试
        /// </summary>
        public async Task RunAllTestsAsync()
        {
            Console.WriteLine("🚀 C#版PPT翻译功能测试开始");
            Console.WriteLine(new string('=', 60));

            try
            {
                // 测试PPT处理器创建
                if (!TestPPTProcessorCreation())
                {
                    Console.WriteLine("❌ PPT处理器创建测试失败，测试终止");
                    return;
                }

                // 测试PPT翻译选项设置
                if (!TestPPTTranslationOptions())
                {
                    Console.WriteLine("❌ PPT翻译选项设置测试失败，测试终止");
                    return;
                }

                // 创建示例PPT文件
                var testFilePath = CreateSamplePPTFile();
                if (string.IsNullOrEmpty(testFilePath))
                {
                    Console.WriteLine("❌ 示例PPT文件创建失败，测试终止");
                    return;
                }

                // 测试PPT文件验证
                if (!TestPPTFileValidation(testFilePath))
                {
                    Console.WriteLine("❌ PPT文件验证测试失败，测试终止");
                    return;
                }

                // 测试PPT信息提取
                if (!TestPPTInfoExtraction(testFilePath))
                {
                    Console.WriteLine("❌ PPT信息提取测试失败，测试终止");
                    return;
                }

                // 测试PPT翻译处理流程
                if (!await TestPPTTranslationProcessAsync(testFilePath))
                {
                    Console.WriteLine("❌ PPT翻译处理测试失败，测试终止");
                    return;
                }

                Console.WriteLine("\n" + new string('=', 60));
                Console.WriteLine("🎉 所有C#版PPT翻译功能测试通过！");
                Console.WriteLine(new string('=', 60));

                // 清理测试文件
                CleanupTestFiles(testFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"💥 测试过程中发生错误: {ex.Message}");
                Console.WriteLine($"堆栈跟踪: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 测试PPT处理器创建
        /// </summary>
        private bool TestPPTProcessorCreation()
        {
            Console.WriteLine(new string('=', 50));
            Console.WriteLine("测试PPT处理器创建");
            Console.WriteLine(new string('=', 50));

            try
            {
                if (_pptProcessor == null)
                {
                    Console.WriteLine("❌ PPT处理器创建失败");
                    return false;
                }

                Console.WriteLine("✅ PPT处理器创建成功");
                Console.WriteLine($"   - 术语库支持: {_pptProcessor.GetType().GetField("_useTerminology", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(_pptProcessor)}");
                Console.WriteLine($"   - 输出格式: {_pptProcessor.GetType().GetField("_outputFormat", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(_pptProcessor)}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ PPT处理器创建测试失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 测试PPT翻译选项设置
        /// </summary>
        private bool TestPPTTranslationOptions()
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine("测试PPT翻译选项设置");
            Console.WriteLine(new string('=', 50));

            try
            {
                // 设置翻译选项
                _pptProcessor.SetTranslationOptions(
                    useTerminology: true,
                    preprocessTerms: true,
                    outputFormat: "bilingual",
                    preserveFormatting: true,
                    translateNotes: true,
                    translateCharts: true,
                    translateSmartArt: true,
                    autoAdjustLayout: true
                );

                Console.WriteLine("✅ PPT翻译选项设置成功");
                Console.WriteLine("   - 术语库: 已启用");
                Console.WriteLine("   - 预处理: 已启用");
                Console.WriteLine("   - 输出格式: 双语对照");
                Console.WriteLine("   - 保留格式: 已启用");
                Console.WriteLine("   - 翻译备注: 已启用");
                Console.WriteLine("   - 翻译图表: 已启用");
                Console.WriteLine("   - 翻译SmartArt: 已启用");
                Console.WriteLine("   - 自动调整: 已启用");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ PPT翻译选项设置测试失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建示例PPT文件
        /// </summary>
        private string CreateSamplePPTFile()
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine("创建示例PPT文件");
            Console.WriteLine(new string('=', 50));

            try
            {
                // 这里应该使用OpenXML创建示例PPT文件
                // 由于复杂性，暂时创建一个简单的测试文件
                var testFilePath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "测试PPT文件.pptx");
                
                // 创建一个简单的测试文件（实际项目中应该使用OpenXML创建真实的PPT）
                File.WriteAllText(testFilePath, "这是一个测试PPT文件");
                
                Console.WriteLine($"✅ 示例PPT文件创建成功: {testFilePath}");
                return testFilePath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ 示例PPT文件创建失败: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 测试PPT文件验证
        /// </summary>
        private bool TestPPTFileValidation(string testFilePath)
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine("测试PPT文件验证功能");
            Console.WriteLine(new string('=', 50));

            try
            {
                if (!File.Exists(testFilePath))
                {
                    Console.WriteLine($"❌ 测试文件不存在: {testFilePath}");
                    return false;
                }

                // 验证PPT文件
                var validationResult = _pptProcessor.ValidatePPTFile(testFilePath);

                if ((bool)validationResult["valid"])
                {
                    Console.WriteLine("✅ PPT文件验证成功");
                    Console.WriteLine($"   - 文件大小: {Convert.ToInt64(validationResult["file_size"]) / 1024.0:F2} KB");
                    Console.WriteLine($"   - 幻灯片数量: {validationResult["slide_count"]}");
                    
                    var supportedElements = validationResult["supported_elements"] as List<string>;
                    if (supportedElements != null && supportedElements.Count > 0)
                    {
                        Console.WriteLine($"   - 支持的元素: {string.Join(", ", supportedElements)}");
                    }
                }
                else
                {
                    Console.WriteLine($"❌ PPT文件验证失败: {validationResult["error"]}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ PPT文件验证测试失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 测试PPT信息提取
        /// </summary>
        private bool TestPPTInfoExtraction(string testFilePath)
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine("测试PPT信息提取功能");
            Console.WriteLine(new string('=', 50));

            try
            {
                // 获取幻灯片数量
                var slideCount = _pptProcessor.GetSlideCount(testFilePath);
                Console.WriteLine($"✅ 幻灯片数量获取成功: {slideCount} 张");

                // 获取PPT详细信息
                var slideInfo = _pptProcessor.GetSlideInfo(testFilePath);
                if (slideInfo.Count > 0)
                {
                    Console.WriteLine("✅ PPT详细信息获取成功:");
                    Console.WriteLine($"   - 幻灯片数量: {slideInfo.GetValueOrDefault("slide_count", 0)}");
                    Console.WriteLine($"   - 幻灯片尺寸: {slideInfo.GetValueOrDefault("slide_size", "未知")}");
                    Console.WriteLine($"   - 包含备注: {slideInfo.GetValueOrDefault("has_notes", false)}");
                    Console.WriteLine($"   - 包含图表: {slideInfo.GetValueOrDefault("has_charts", false)}");
                    Console.WriteLine($"   - 包含表格: {slideInfo.GetValueOrDefault("has_tables", false)}");
                    Console.WriteLine($"   - 包含SmartArt: {slideInfo.GetValueOrDefault("has_smartart", false)}");
                }
                else
                {
                    Console.WriteLine("❌ PPT详细信息获取失败");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ PPT信息提取测试失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 测试PPT翻译处理流程
        /// </summary>
        private async Task<bool> TestPPTTranslationProcessAsync(string testFilePath)
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine("测试PPT翻译处理流程");
            Console.WriteLine(new string('=', 50));

            try
            {
                // 模拟术语库
                var terminology = new Dictionary<string, Dictionary<string, string>>
                {
                    ["英语"] = new Dictionary<string, string>
                    {
                        ["人工智能"] = "Artificial Intelligence",
                        ["机器学习"] = "Machine Learning",
                        ["深度学习"] = "Deep Learning",
                        ["神经网络"] = "Neural Network",
                        ["自然语言处理"] = "Natural Language Processing"
                    }
                };

                // 设置进度回调
                void ProgressCallback(double progress, string message)
                {
                    Console.WriteLine($"   进度: {progress:P1} - {message}");
                }

                _pptProcessor.SetProgressCallback(ProgressCallback);

                Console.WriteLine("开始PPT翻译处理...");

                // 执行翻译
                var outputPath = await _pptProcessor.ProcessDocumentAsync(
                    filePath: testFilePath,
                    targetLanguage: "英语",
                    terminology: terminology["英语"]
                );

                if (!string.IsNullOrEmpty(outputPath) && File.Exists(outputPath))
                {
                    Console.WriteLine("✅ PPT翻译处理成功");
                    Console.WriteLine($"   输出文件: {outputPath}");
                    return true;
                }
                else
                {
                    Console.WriteLine("❌ PPT翻译处理失败");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ PPT翻译处理测试失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 清理测试文件
        /// </summary>
        private void CleanupTestFiles(string testFilePath)
        {
            try
            {
                if (File.Exists(testFilePath))
                {
                    File.Delete(testFilePath);
                    Console.WriteLine($"✅ 测试文件已清理: {testFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ 测试文件清理失败: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 模拟翻译服务
    /// </summary>
    public class MockTranslationService : TranslationService
    {
        public MockTranslationService(ILogger<TranslationService> logger) : base(logger)
        {
        }

        public override async Task<string> TranslateTextAsync(string text, Dictionary<string, string> terminology, 
            string sourceLang, string targetLang, string prompt = null)
        {
            // 模拟翻译结果
            var translatedText = $"[TRANSLATED] {text}";
            return await Task.FromResult(translatedText);
        }
    }
} 