using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// 模型管理服务
    /// 负责扫描、检测和管理RWKV模型文件
    /// </summary>
    public class ModelManagementService
    {
        private readonly ILogger _logger;
        private readonly string _modelsDirectory;
        private readonly string _vocabFilePath;
        private readonly string _engineDirectory;
        private readonly string _llamaCppBaseDirectory;
        private static readonly string[] EngineDirectoryNames =
        {
            "rwkv_lightning_libtorch_win",
            "rwkv_lightning_libtorch_win",
            "rwkv_lightning_libtorch"
        };

        // 模型参数量估算正则表达式
        private static readonly Regex ParameterPattern = new Regex(@"(\d+(?:\.\d+)?)\s*[Bb]", RegexOptions.Compiled);
        private static readonly Regex ContextPattern = new Regex(@"ctx(\d+)", RegexOptions.Compiled);

        public ModelManagementService(ILogger logger, string? modelsDirectory = null, string? engineDirectory = null)
        {
            _logger = logger;
            
            // 使用 GetExecutableDirectory() 而不是 GetSafeBaseDirectory()，这在单文件部署时更可靠
            var baseDir = PathHelper.GetExecutableDirectory();
            baseDir = Path.GetFullPath(baseDir);
            
            var projectRoot = PathHelper.FindProjectRoot(baseDir);
            
            _modelsDirectory = modelsDirectory ?? ResolveModelsDirectory(projectRoot, baseDir);
            
            if (engineDirectory != null)
            {
                _engineDirectory = engineDirectory;
            }
            else
            {
                _engineDirectory = ResolveEngineDirectory(projectRoot, baseDir);
            }
            
            _vocabFilePath = Path.Combine(_engineDirectory, "rwkv_vocab_v20230424.txt");
            
            // llama.cpp 目录
            _llamaCppBaseDirectory = ResolveLlamaCppDirectory(projectRoot, baseDir);
            
            _logger.LogInformation("推理引擎目录: {Path}", _engineDirectory);
            _logger.LogInformation("模型目录: {Path}", _modelsDirectory);
            _logger.LogInformation("llama.cpp目录: {Path}", _llamaCppBaseDirectory);
        }
        
        private static string ResolveModelsDirectory(string projectRoot, string baseDir)
        {
            var possibleModelPaths = new[]
            {
                Path.Combine(baseDir, "rwkv_models"),
                Path.Combine(projectRoot, "rwkv_models")
            };

            return possibleModelPaths.FirstOrDefault(Directory.Exists) ?? possibleModelPaths[0];
        }

        private static string ResolveEngineDirectory(string projectRoot, string baseDir)
        {
            var candidateRoots = new List<string>
            {
                baseDir,
                projectRoot
            };

            var parent = Directory.GetParent(baseDir);
            for (int i = 0; i < 3 && parent != null; i++)
            {
                candidateRoots.Add(parent.FullName);
                parent = parent.Parent;
            }

            var possibleEnginePaths = candidateRoots
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .SelectMany(root => GetEngineCandidatePaths(root))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return possibleEnginePaths.FirstOrDefault(HasUsableEngineFiles)
                ?? possibleEnginePaths.FirstOrDefault(Directory.Exists)
                ?? possibleEnginePaths[0];
        }

        private static bool HasUsableEngineFiles(string engineDirectory)
        {
            return Directory.Exists(engineDirectory)
                && File.Exists(Path.Combine(engineDirectory, "rwkv_lightning.exe"))
                && File.Exists(Path.Combine(engineDirectory, "rwkv_vocab_v20230424.txt"));
        }

        private static string ResolveLlamaCppDirectory(string projectRoot, string baseDir)
        {
            var candidateRoots = new List<string> { baseDir, projectRoot };
            var parent = Directory.GetParent(baseDir);
            for (int i = 0; i < 3 && parent != null; i++)
            {
                candidateRoots.Add(parent.FullName);
                parent = parent.Parent;
            }

            var possiblePaths = candidateRoots
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Select(root => Path.Combine(root, "llama_cpp"))
                .Select(Path.GetFullPath)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return possiblePaths.FirstOrDefault(Directory.Exists) ?? possiblePaths.FirstOrDefault() ?? "";
        }

        private static IEnumerable<string> GetEngineCandidatePaths(string root)
        {
            foreach (var name in EngineDirectoryNames)
            {
                var engineRoot = Path.Combine(root, name);
                yield return engineRoot;
                yield return Path.Combine(engineRoot, "dist_native", "bin");
                yield return Path.Combine(engineRoot, "build_native", "Release");
                yield return Path.Combine(engineRoot, "build", "Release");
                yield return Path.Combine(engineRoot, "bin", "Release");
            }

            yield return Path.Combine(root, "transplantation", "dist_native", "bin");
            yield return Path.Combine(root, "transplantation", "rwkv_lightning_libtorch", "dist_native", "bin");
            yield return Path.Combine(root, "transplantation", "rwkv_lightning_libtorch", "build_native", "Release");
            yield return Path.Combine(root, "transplantation", "rwkv_lightning_libtorch", "build", "Release");
        }

        /// <summary>
        /// 词汇表文件路径
        /// </summary>
        public string VocabFilePath => _vocabFilePath;

        public string ModelsDirectory => _modelsDirectory;

        /// <summary>
        /// 推理引擎可执行文件路径（默认 rwkv_lightning.exe）
        /// </summary>
        public string EngineExePath => Path.Combine(_engineDirectory, "rwkv_lightning.exe");

        /// <summary>
        /// 获取指定推理工具的可执行文件路径
        /// </summary>
        /// <param name="toolName">工具名称（rwkv_lightning, benchmark, llama-cuda, llama-cpu, llama-vulkan）</param>
        /// <returns>可执行文件路径</returns>
        public string GetToolPath(string toolName = "rwkv_lightning")
        {
            return toolName switch
            {
                "benchmark" => Path.Combine(_engineDirectory, "benchmark.exe"),
                "llama-cuda" => Path.Combine(_llamaCppBaseDirectory, "cuda", "llama-server.exe"),
                "llama-sycl" => Path.Combine(_llamaCppBaseDirectory, "sycl", "llama-server.exe"),
                "llama-vulkan" => Path.Combine(_llamaCppBaseDirectory, "vulkan", "llama-server.exe"),
                "llama-cpu" => Path.Combine(_llamaCppBaseDirectory, "cpu", "llama-server.exe"),
                _ => Path.Combine(_engineDirectory, "rwkv_lightning.exe")
            };
        }

        /// <summary>
        /// 获取llama.cpp工具的工作目录（用于设置DLL搜索路径）
        /// </summary>
        public string GetLlamaCppWorkingDirectory(string toolName)
        {
            return toolName switch
            {
                "llama-cuda" => Path.Combine(_llamaCppBaseDirectory, "cuda"),
                "llama-sycl" => Path.Combine(_llamaCppBaseDirectory, "sycl"),
                "llama-vulkan" => Path.Combine(_llamaCppBaseDirectory, "vulkan"),
                "llama-cpu" => Path.Combine(_llamaCppBaseDirectory, "cpu"),
                _ => ""
            };
        }

        /// <summary>
        /// 检查指定推理工具是否存在
        /// </summary>
        /// <param name="toolName">工具名称（rwkv_lightning, benchmark, llama-cuda, llama-cpu, llama-vulkan）</param>
        public bool ToolExists(string toolName = "rwkv_lightning")
        {
            var toolPath = GetToolPath(toolName);
            return File.Exists(toolPath);
        }

        /// <summary>
        /// 判断是否为llama.cpp推理工具
        /// </summary>
        public bool IsLlamaCppTool(string toolName)
        {
            return toolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";
        }

        /// <summary>
        /// 检查词汇表文件是否存在
        /// </summary>
        public bool VocabFileExists => File.Exists(_vocabFilePath);

        /// <summary>
        /// 检查推理引擎是否存在
        /// </summary>
        public bool EngineExists => File.Exists(EngineExePath);

        public async Task<string> ImportModelAsync(string sourceFilePath, bool overwrite = false)
        {
            if (string.IsNullOrWhiteSpace(sourceFilePath))
            {
                throw new ArgumentNullException(nameof(sourceFilePath));
            }

            if (!File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException("模型文件不存在", sourceFilePath);
            }

            var extension = Path.GetExtension(sourceFilePath);
            if (!string.Equals(extension, ".pth", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".safetensors", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".st", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".gguf", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("仅支持导入 .pth、.safetensors、.st 或 .gguf 模型文件");
            }

            Directory.CreateDirectory(_modelsDirectory);

            var targetPath = Path.Combine(_modelsDirectory, Path.GetFileName(sourceFilePath));
            if (File.Exists(targetPath) && !overwrite)
            {
                throw new IOException($"目标模型已存在: {targetPath}");
            }

            await Task.Run(() => File.Copy(sourceFilePath, targetPath, overwrite));
            _logger.LogInformation($"模型文件已导入: {targetPath}");
            return targetPath;
        }

        /// <summary>
        /// 扫描所有可用的模型
        /// </summary>
        public async Task<List<ModelInfo>> ScanModelsAsync()
        {
            return await Task.Run(() =>
            {
                var models = new List<ModelInfo>();

                try
                {
                    if (!Directory.Exists(_modelsDirectory))
                    {
                        _logger.LogWarning($"模型目录不存在: {_modelsDirectory}");
                        return models;
                    }

                    // 扫描.safetensors文件（.safetensors 和 .st）
                    var safetensorsFiles = Directory.GetFiles(_modelsDirectory, "*.safetensors", SearchOption.TopDirectoryOnly);
                    foreach (var file in safetensorsFiles)
                    {
                        var model = CreateModelInfo(file, ModelFormat.SafeTensors);
                        model.State = ModelState.Available;
                        models.Add(model);
                    }

                    var stFiles = Directory.GetFiles(_modelsDirectory, "*.st", SearchOption.TopDirectoryOnly);
                    foreach (var file in stFiles)
                    {
                        var model = CreateModelInfo(file, ModelFormat.SafeTensors);
                        model.State = ModelState.Available;
                        models.Add(model);
                    }

                    // 扫描.pth文件
                    var pthFiles = Directory.GetFiles(_modelsDirectory, "*.pth", SearchOption.TopDirectoryOnly);
                    foreach (var file in pthFiles)
                    {
                        var model = CreateModelInfo(file, ModelFormat.PyTorch);
                        
                        // 检查是否已有转换后的文件
                        var convertedPath = GetConvertedFilePath(file);
                        if (File.Exists(convertedPath))
                        {
                            model.State = ModelState.Available;
                            model.ConvertedFilePath = convertedPath;
                            model.FilePath = convertedPath;
                            model.Format = ModelFormat.SafeTensors;
                        }
                        else
                        {
                            model.State = ModelState.NeedsConversion;
                        }
                        
                        models.Add(model);
                    }

                    // 扫描.gguf文件
                    var ggufFiles = Directory.GetFiles(_modelsDirectory, "*.gguf", SearchOption.TopDirectoryOnly);
                    foreach (var file in ggufFiles)
                    {
                        var model = CreateModelInfo(file, ModelFormat.GGUF);
                        model.State = ModelState.Available;
                        models.Add(model);
                    }

                    // 按文件大小降序排序（大模型优先）
                    models = models.OrderByDescending(m => m.FileSizeBytes).ToList();

                    _logger.LogInformation($"扫描完成，找到 {models.Count} 个模型文件");
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "扫描模型文件时发生异常");
                }

                return models;
            });
        }

        /// <summary>
        /// 创建模型信息对象
        /// </summary>
        private ModelInfo CreateModelInfo(string filePath, ModelFormat format)
        {
            var fileInfo = new FileInfo(filePath);
            var model = new ModelInfo
            {
                FilePath = filePath,
                Format = format,
                FileSizeBytes = fileInfo.Length
            };

            // 从文件名提取参数量
            var paramMatch = ParameterPattern.Match(model.ModelName);
            if (paramMatch.Success && double.TryParse(paramMatch.Groups[1].Value, out var paramsNum))
            {
                model.EstimatedParameters = paramMatch.Groups[1].Value + "B";
                model.EstimatedMemoryGB = EstimateMemoryRequirement(paramsNum);
            }

            // 从文件名提取上下文长度
            var ctxMatch = ContextPattern.Match(model.ModelName);
            if (ctxMatch.Success && int.TryParse(ctxMatch.Groups[1].Value, out var ctxLen))
            {
                model.ContextLength = ctxLen;
            }

            return model;
        }

        /// <summary>
        /// 估算模型显存需求
        /// </summary>
        private double EstimateMemoryRequirement(double parametersBillion)
        {
            // 粗略估算：FP16模型约需要 2GB/B参数 + 上下文缓存
            // 1.5B ≈ 3GB + 缓存
            // 2.9B ≈ 6GB + 缓存
            // 7B ≈ 14GB + 缓存
            return parametersBillion * 2.0 + 0.5;
        }

        /// <summary>
        /// 获取转换后的文件路径
        /// </summary>
        public string GetConvertedFilePath(string pthPath)
        {
            var dir = Path.GetDirectoryName(pthPath) ?? "";
            var name = Path.GetFileNameWithoutExtension(pthPath);
            return Path.Combine(dir, $"{name}.st");
        }

        /// <summary>
        /// 获取模型推理所需的总显存估算
        /// </summary>
        public double GetEstimatedMemoryForConcurrency(ModelInfo model, int concurrency)
        {
            // 基础显存 + 每并发额外显存
            double baseMemory = model.EstimatedMemoryGB ?? 3.0;
            double perConcurrencyMemory = model.EstimatedParameters switch
            {
                "1.5B" => 0.1,
                "2.9B" => 0.2,
                "7B" or "7.2B" => 0.5,
                "13B" or "13.3B" => 1.0,
                _ => 0.2
            };

            return baseMemory + (concurrency * perConcurrencyMemory);
        }

        /// <summary>
        /// 检查模型是否可用（已转换或本身就是safetensors格式）
        /// </summary>
        public bool IsModelReady(ModelInfo model)
        {
            if ((model.Format == ModelFormat.SafeTensors || model.Format == ModelFormat.GGUF) && model.State == ModelState.Available)
            {
                return File.Exists(model.FilePath);
            }

            if (model.Format == ModelFormat.PyTorch && !string.IsNullOrEmpty(model.ConvertedFilePath))
            {
                return File.Exists(model.ConvertedFilePath);
            }

            return false;
        }

        /// <summary>
        /// 获取模型的实际可用路径
        /// </summary>
        public string GetModelPath(ModelInfo model)
        {
            if (model.Format == ModelFormat.SafeTensors || model.Format == ModelFormat.GGUF)
            {
                return model.FilePath;
            }

            if (!string.IsNullOrEmpty(model.ConvertedFilePath) && File.Exists(model.ConvertedFilePath))
            {
                return model.ConvertedFilePath;
            }

            // 返回预期转换后的路径
            return GetConvertedFilePath(model.FilePath);
        }
    }
}
