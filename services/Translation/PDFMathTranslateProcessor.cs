using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 基于PDFMathTranslate的PDF文档处理器
    /// 使用pdf2zh Python模块实现无损PDF翻译，支持：
    /// - 保持原始排版和格式
    /// - LaTeX公式保护
    /// - 图片和表格保留
    /// - 双语对照输出
    /// </summary>
    public class PDFMathTranslateProcessor
    {
        private readonly ILogger<PDFMathTranslateProcessor> _logger;
        private readonly ITranslator _translator;
        private readonly ILoggerFactory _loggerFactory;
        private readonly ConfigurationManager _configManager;
        private bool _useTerminology = true;
        private bool _preprocessTerms = true;
        private bool _exportPdf = true;
        private string _sourceLang = "zh";
        private string _targetLang = "en";
        private string _outputFormat = "bilingual";
        private Action<double, string> _progressCallback;
        private string _pythonRuntimePath;
        private string _pdfTranslateWrapperPath;
        private DateTime _lastProgressUpdate = DateTime.MinValue;
        private const double _minProgressUpdateInterval = 100; // 最小进度更新间隔100ms

        public PDFMathTranslateProcessor(ILogger<PDFMathTranslateProcessor> logger, ITranslator translator, ILoggerFactory loggerFactory = null, ConfigurationManager configManager = null)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _translator = translator ?? throw new ArgumentNullException(nameof(translator));
            _loggerFactory = loggerFactory ?? Microsoft.Extensions.Logging.Abstractions.NullLoggerFactory.Instance;
            _configManager = configManager;
            DetectPythonRuntimePath();
        }

        private void DetectPythonRuntimePath()
        {
            try
            {
                var baseDir = PathHelper.GetSafeBaseDirectory();
                var possiblePythonPaths = new[]
                {
                    Path.Combine(baseDir, "pdf2zh", "build", "runtime", "python.exe"),
                    Path.Combine(baseDir, "pdf2zh", "build", "python.exe"),
                    Path.Combine(Directory.GetParent(baseDir)?.Parent?.Parent?.Parent?.Parent?.FullName ?? "", "pdf2zh", "build", "runtime", "python.exe"),
                    Path.Combine(Directory.GetCurrentDirectory(), "pdf2zh", "build", "runtime", "python.exe"),
                    @"d:\AI_project\document-term-translator\pdf2zh\build\runtime\python.exe"
                };

                foreach (var path in possiblePythonPaths)
                {
                    if (File.Exists(path))
                    {
                        _pythonRuntimePath = path;
                        _logger.LogInformation($"检测到Python运行时路径: {_pythonRuntimePath}");
                        break;
                    }
                }

                var possibleWrapperPaths = new[]
                {
                    Path.Combine(baseDir, "pdf2zh", "pdf_translate_wrapper.py"),
                    Path.Combine(baseDir, "pdf2zh", "pdf_translate_wrapper.py"),
                    Path.Combine(Directory.GetParent(baseDir)?.Parent?.Parent?.Parent?.Parent?.FullName ?? "", "pdf2zh", "pdf_translate_wrapper.py"),
                    Path.Combine(Directory.GetCurrentDirectory(), "pdf2zh", "pdf_translate_wrapper.py"),
                    @"d:\AI_project\document-term-translator\pdf2zh\pdf_translate_wrapper.py"
                };

                foreach (var path in possibleWrapperPaths)
                {
                    if (File.Exists(path))
                    {
                        _pdfTranslateWrapperPath = path;
                        _logger.LogInformation($"检测到PDF翻译包装器路径: {_pdfTranslateWrapperPath}");
                        break;
                    }
                }

                if (string.IsNullOrEmpty(_pythonRuntimePath))
                {
                    _logger.LogWarning("未检测到Python运行时，请确保pdf2zh/build/runtime/python.exe存在");
                }

                if (string.IsNullOrEmpty(_pdfTranslateWrapperPath))
                {
                    _logger.LogWarning("未检测到PDF翻译包装器，请确保pdf2zh/pdf_translate_wrapper.py存在");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "检测Python运行时路径时出错");
            }
        }

        public void SetProgressCallback(Action<double, string> callback)
        {
            _progressCallback = callback;
        }

        public void SetTranslationOptions(bool useTerminology = true, bool preprocessTerms = true,
            bool exportPdf = true, string sourceLang = "zh", string targetLang = "en", string outputFormat = "bilingual")
        {
            _useTerminology = useTerminology;
            _preprocessTerms = preprocessTerms;
            _exportPdf = exportPdf;
            _sourceLang = sourceLang;
            _targetLang = targetLang;
            _outputFormat = outputFormat;

            _logger.LogInformation($"PDFMathTranslate处理器配置:");
            _logger.LogInformation($"  源语言: {sourceLang}");
            _logger.LogInformation($"  目标语言: {targetLang}");
            _logger.LogInformation($"  输出格式: {outputFormat}");
            _logger.LogInformation($"  使用术语库: {useTerminology}");
            _logger.LogInformation($"  导出PDF: {exportPdf}");
        }

        public void SetPdf2zhPath(string path)
        {
            _pythonRuntimePath = path;
            _logger.LogInformation("Python运行时路径设置为: {Path}", path);
        }

        public async Task<string> ProcessDocumentAsync(string filePath, string targetLanguage,
            Dictionary<string, string> terminology)
        {
            UpdateProgress(0.01, "开始处理PDF文档...");

            if (!File.Exists(filePath))
            {
                _logger.LogError($"文件不存在: {filePath}");
                throw new FileNotFoundException("文件不存在");
            }

            var outputDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "输出");
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var fileName = Path.GetFileNameWithoutExtension(filePath);
            var timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var outputFileName = $"{fileName}_翻译_{timeStamp}.pdf";
            var outputPath = Path.Combine(outputDir, outputFileName);

            _logger.LogInformation($"PDFMathTranslate处理器输出配置:");
            _logger.LogInformation($"  原始文件: {filePath}");
            _logger.LogInformation($"  输出路径: {outputPath}");
            _logger.LogInformation($"  目标语言: {targetLanguage}");

            try
            {
                await ProcessPDFWithPdf2zhAsync(filePath, outputPath, targetLanguage, terminology);

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

        private async Task ProcessPDFWithPdf2zhAsync(string inputPath, string outputPath,
            string targetLanguage, Dictionary<string, string> terminology)
        {
            UpdateProgress(0.05, "检查Python运行时和pdf2zh模块...");

            var pdf2zhAvailable = await CheckPdf2zhAvailableAsync();
            if (!pdf2zhAvailable)
            {
                _logger.LogError("Python运行时或pdf2zh模块不可用");
                throw new InvalidOperationException("Python运行时或pdf2zh模块不可用");
            }

            UpdateProgress(0.1, "准备翻译参数...");

            var langPair = GetLanguagePair(_sourceLang, targetLanguage);
            var outputDir = Path.GetDirectoryName(outputPath);

            var arguments = BuildPdf2zhArguments(inputPath, outputDir, langPair, terminology);

            _logger.LogInformation("执行PDF翻译命令: {Path} {Arguments}", _pythonRuntimePath, arguments);

            UpdateProgress(0.2, "开始翻译PDF文档...");

            var success = await ExecutePdf2zhAsync(arguments, inputPath);

            if (!success)
            {
                _logger.LogError("PDF翻译执行失败");
                throw new InvalidOperationException("PDF翻译执行失败");
            }

            UpdateProgress(0.9, "检查输出文件...");

            var fileName = Path.GetFileNameWithoutExtension(inputPath);
            var expectedMonoPath = Path.Combine(outputDir, $"{fileName}-mono.pdf");
            var expectedDualPath = Path.Combine(outputDir, $"{fileName}-dual.pdf");

            if (!File.Exists(expectedMonoPath) && !File.Exists(expectedDualPath))
            {
                _logger.LogError($"输出文件未生成: {expectedMonoPath} 或 {expectedDualPath}");
                throw new FileNotFoundException("翻译输出文件未生成");
            }

            var outputFilePath = _outputFormat == "bilingual" ? expectedDualPath : expectedMonoPath;
            if (File.Exists(outputFilePath))
            {
                File.Copy(outputFilePath, outputPath, true);
                var fileInfo = new FileInfo(outputPath);
                _logger.LogInformation($"翻译完成，输出文件大小: {fileInfo.Length / 1024.0:F2} KB");
            }

            UpdateProgress(0.95, "清理临时文件...");
        }

        private async Task<bool> CheckPdf2zhAvailableAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_pythonRuntimePath))
                {
                    _logger.LogError("Python运行时路径未设置");
                    return false;
                }

                if (!File.Exists(_pythonRuntimePath))
                {
                    _logger.LogError($"Python运行时文件不存在: {_pythonRuntimePath}");
                    return false;
                }

                if (string.IsNullOrEmpty(_pdfTranslateWrapperPath))
                {
                    _logger.LogError("PDF翻译包装器路径未设置");
                    return false;
                }

                if (!File.Exists(_pdfTranslateWrapperPath))
                {
                    _logger.LogError($"PDF翻译包装器文件不存在: {_pdfTranslateWrapperPath}");
                    return false;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = _pythonRuntimePath,
                    Arguments = $"--version",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                _logger.LogInformation($"开始检查Python运行时可用性 (路径: {_pythonRuntimePath})...");

                using var process = Process.Start(startInfo);
                if (process == null)
                {
                    _logger.LogWarning($"无法启动Python进程: {_pythonRuntimePath}");
                    return false;
                }

                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                await process.WaitForExitAsync();

                _logger.LogInformation($"Python检查结果 - 退出码: {process.ExitCode}");
                if (!string.IsNullOrEmpty(output))
                {
                    _logger.LogInformation($"标准输出: {output.Trim()}");
                }
                if (!string.IsNullOrEmpty(error))
                {
                    _logger.LogWarning($"标准错误: {error.Trim()}");
                }

                var isAvailable = process.ExitCode == 0;
                _logger.LogInformation($"Python运行时可用性: {(isAvailable ? "可用" : "不可用")}");

                return isAvailable;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "检查Python运行时可用性时出错");
                return false;
            }
        }

        private string BuildPdf2zhArguments(string inputPath, string outputDir, string langPair,
            Dictionary<string, string> terminology)
        {
            var args = new StringBuilder();

            args.Append($"\"{_pdfTranslateWrapperPath}\"");
            args.Append($" \"{inputPath}\"");
            args.Append($" \"{outputDir}\"");

            var langParts = langPair.Split('2');
            if (langParts.Length >= 2)
            {
                args.Append($" --lang-in {langParts[0]}");
                args.Append($" --lang-out {langParts[1]}");
            }

            args.Append(" --service openai");
            args.Append(" --thread 4");

            var envConfig = GetTranslatorApiConfig();

            var envJson = JsonSerializer.Serialize(envConfig);
            var escapedJson = envJson.Replace("\"", "\\\"");
            args.Append($" --env \"{escapedJson}\"");

            return args.ToString();
        }

        private Dictionary<string, string> GetTranslatorApiConfig()
        {
            var config = new Dictionary<string, string>();

            try
            {
                if (_configManager == null)
                {
                    _logger.LogWarning("ConfigurationManager未设置，使用默认配置");
                    config["OPENAI_BASE_URL"] = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                    config["OPENAI_API_KEY"] = "";
                    config["OPENAI_MODEL"] = "glm-4-flash";
                    return config;
                }

                var translatorType = _translator?.GetType().Name.ToLower() ?? "";

                if (translatorType.Contains("zhipuai"))
                {
                    var apiKey = _configManager.GetApiKey("zhipuai");
                    var translatorConfig = _configManager.GetTranslatorConfig("zhipuai");
                    config["OPENAI_BASE_URL"] = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                    config["OPENAI_API_KEY"] = apiKey ?? "";
                    config["OPENAI_MODEL"] = translatorConfig?.Model ?? "glm-4-flash";
                    _logger.LogInformation($"使用智谱AI配置: BaseURL={config["OPENAI_BASE_URL"]}, Model={config["OPENAI_MODEL"]}");
                }
                else if (translatorType.Contains("hunyuan"))
                {
                    var apiKey = _configManager.GetApiKey("hunyuantranslator");
                    var translatorConfig = _configManager.GetTranslatorConfig("hunyuantranslator");
                    config["OPENAI_BASE_URL"] = translatorConfig?.ApiUrl ?? "https://hunyuan.tencentcloudapi.com/v1/chat/completions";
                    config["OPENAI_API_KEY"] = apiKey ?? "";
                    config["OPENAI_MODEL"] = translatorConfig?.Model ?? "Hunyuan-MT-7B";
                    _logger.LogInformation($"使用腾讯混元配置: BaseURL={config["OPENAI_BASE_URL"]}, Model={config["OPENAI_MODEL"]}");
                }
                else if (translatorType.Contains("siliconflow"))
                {
                    var apiKey = _configManager.GetApiKey("siliconflow");
                    var translatorConfig = _configManager.GetTranslatorConfig("siliconflow");
                    config["OPENAI_BASE_URL"] = "https://api.siliconflow.cn/v1/chat/completions";
                    config["OPENAI_API_KEY"] = apiKey ?? "";
                    config["OPENAI_MODEL"] = translatorConfig?.Model ?? "deepseek-ai/DeepSeek-V3";
                    _logger.LogInformation($"使用硅基流动配置: BaseURL={config["OPENAI_BASE_URL"]}, Model={config["OPENAI_MODEL"]}");
                }
                else if (translatorType.Contains("ollama"))
                {
                    var translatorConfig = _configManager.GetTranslatorConfig("ollama");
                    config["OPENAI_BASE_URL"] = "http://localhost:11434/v1/chat/completions";
                    config["OPENAI_API_KEY"] = "ollama";
                    config["OPENAI_MODEL"] = translatorConfig?.Model ?? "gemma3:4b";
                    _logger.LogInformation($"使用Ollama配置: BaseURL={config["OPENAI_BASE_URL"]}, Model={config["OPENAI_MODEL"]}");
                }
                else if (translatorType.Contains("rwkv"))
                {
                    var translatorConfig = _configManager.GetTranslatorConfig("rwkv");
                    config["OPENAI_BASE_URL"] = translatorConfig?.ApiUrl ?? "http://localhost:8000/v1/chat/completions";
                    config["OPENAI_API_KEY"] = "";
                    config["OPENAI_MODEL"] = translatorConfig?.Model ?? "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118";
                    _logger.LogInformation($"使用RWKV配置: BaseURL={config["OPENAI_BASE_URL"]}, Model={config["OPENAI_MODEL"]}");
                }
                else
                {
                    _logger.LogWarning($"未知的翻译器类型: {translatorType}，使用默认配置");
                    config["OPENAI_BASE_URL"] = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                    config["OPENAI_API_KEY"] = "";
                    config["OPENAI_MODEL"] = "glm-4-flash";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "获取翻译器API配置失败，使用默认配置");
                config["OPENAI_BASE_URL"] = "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                config["OPENAI_API_KEY"] = "";
                config["OPENAI_MODEL"] = "glm-4-flash";
            }

            return config;
        }

        private string CreateTerminologyFile(Dictionary<string, string> terminology)
        {
            try
            {
                var tempDir = Path.GetTempPath();
                var tempFile = Path.Combine(tempDir, $"terminology_{Guid.NewGuid()}.txt");

                using var writer = new StreamWriter(tempFile, false, Encoding.UTF8);
                foreach (var term in terminology)
                {
                    writer.WriteLine($"{term.Key}\t{term.Value}");
                }

                _logger.LogInformation($"创建术语库文件: {tempFile}");
                return tempFile;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "创建术语库文件失败");
                return null;
            }
        }

        private async Task<bool> ExecutePdf2zhAsync(string arguments, string inputPath)
        {
            try
            {
                if (string.IsNullOrEmpty(_pythonRuntimePath))
                {
                    _logger.LogError("Python运行时路径未设置");
                    return false;
                }

                if (!File.Exists(_pythonRuntimePath))
                {
                    _logger.LogError("Python运行时文件不存在: {Path}", _pythonRuntimePath);
                    return false;
                }

                var startInfo = new ProcessStartInfo
                {
                    FileName = _pythonRuntimePath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    WorkingDirectory = PathHelper.GetSafeBaseDirectory()
                };

                _logger.LogInformation("执行PDF翻译命令: {Path} {Arguments}", _pythonRuntimePath, arguments);

                using var process = Process.Start(startInfo);
                if (process == null)
                {
                    _logger.LogError("无法启动Python进程: {Path}", _pythonRuntimePath);
                    return false;
                }

                var outputBuilder = new StringBuilder();
                var errorBuilder = new StringBuilder();

                process.OutputDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        outputBuilder.AppendLine(e.Data);
                        ParseProgressFromOutput(e.Data);
                    }
                };

                process.ErrorDataReceived += (sender, e) =>
                {
                    if (e.Data != null)
                    {
                        errorBuilder.AppendLine(e.Data);
                    }
                };

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                await process.WaitForExitAsync();

                var output = outputBuilder.ToString();
                var error = errorBuilder.ToString();

                _logger.LogInformation("PDF翻译输出:\n{Output}", output);
                if (!string.IsNullOrEmpty(error))
                {
                    _logger.LogWarning("PDF翻译错误:\n{Error}", error);
                }

                return process.ExitCode == 0;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "执行PDF翻译时出错");
                return false;
            }
        }

        private void ParseProgressFromOutput(string output)
        {
            try
            {
                var progressMatch = Regex.Match(output, @"(\d+)%");
                if (progressMatch.Success)
                {
                    var progress = int.Parse(progressMatch.Groups[1].Value) / 100.0;
                    UpdateProgress(progress, $"翻译进度: {progressMatch.Groups[1].Value}%");
                }

                var pageMatch = Regex.Match(output, @"Page\s+(\d+)/(\d+)");
                if (pageMatch.Success)
                {
                    var currentPage = int.Parse(pageMatch.Groups[1].Value);
                    var totalPages = int.Parse(pageMatch.Groups[2].Value);
                    var progress = 0.2 + (0.7 * currentPage / totalPages);
                    UpdateProgress(progress, $"正在翻译第 {currentPage}/{totalPages} 页...");
                }
            }
            catch
            {
            }
        }

        private string GetLanguagePair(string sourceLang, string targetLang)
        {
            var source = sourceLang.ToLower();
            var target = targetLang.ToLower();

            if (source == "zh" && target == "en")
                return "zh2en";
            if (source == "en" && target == "zh")
                return "en2zh";
            if (source == "zh" && target == "ja")
                return "zh2ja";
            if (source == "ja" && target == "zh")
                return "ja2zh";
            if (source == "zh" && target == "ko")
                return "zh2ko";
            if (source == "ko" && target == "zh")
                return "ko2zh";
            if (source == "zh" && target == "es")
                return "zh2es";
            if (source == "es" && target == "zh")
                return "es2zh";
            if (source == "zh" && target == "fr")
                return "zh2fr";
            if (source == "fr" && target == "zh")
                return "fr2zh";
            if (source == "zh" && target == "de")
                return "zh2de";
            if (source == "de" && target == "zh")
                return "de2zh";
            if (source == "zh" && target == "ru")
                return "zh2ru";
            if (source == "ru" && target == "zh")
                return "ru2zh";

            return "zh2en";
        }

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
                    _logger.LogInformation($"进度: {progress:P2} - {message}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "更新进度失败");
            }
        }
    }
}
