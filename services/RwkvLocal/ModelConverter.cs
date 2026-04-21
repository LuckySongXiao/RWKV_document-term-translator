using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// 模型转换服务
    /// 负责将.pth格式模型转换为.safetensors格式
    /// 直接使用 rwkv_lightning_libtorch/convert_safetensors.py 进行转换
    /// C++ 加载器会在加载时自动应用剩余的运行时转换（转置 ffn.value.weight 等）
    /// </summary>
    public class ModelConverter
    {
        private const string WslCondaEnvironmentName = "rwkv";
        private const string WslCondaExecutablePath = "/root/miniconda3/bin/conda";
        private static readonly string[] PreferredPythonVersions = { "3.12", "3.11", "3.10" };

        private readonly ILogger _logger;
        private readonly string _baseDirectory;
        private readonly string _converterScriptPath;
        private readonly string _localPythonPackagesPath;
        private string? _pythonPath;
        private string _pythonArgumentsPrefix = string.Empty;
        private string? _pythonVersion;
        private bool? _pythonAvailable;
        private PythonRuntimeKind _pythonRuntimeKind;

        public event EventHandler<ConversionProgressEventArgs>? ProgressChanged;

        public ModelConverter(ILogger logger, string? converterScriptPath = null)
        {
            _logger = logger;

            var baseDir = PathHelper.GetSafeBaseDirectory();
            _baseDirectory = Path.GetFullPath(baseDir);

            _converterScriptPath = converterScriptPath ?? Path.Combine(_baseDirectory, "rwkv_lightning_libtorch", "convert_safetensors.py");
            _localPythonPackagesPath = Path.Combine(_baseDirectory, "python");
        }

        /// <summary>
        /// 检测Python环境是否可用
        /// </summary>
        public async Task<bool> CheckPythonEnvironmentAsync()
        {
            if (_pythonAvailable.HasValue)
            {
                return _pythonAvailable.Value;
            }

            try
            {
                var wslVersion = await TryGetWslCondaPythonVersionAsync();
                if (!string.IsNullOrWhiteSpace(wslVersion))
                {
                    _pythonPath = "wsl";
                    _pythonArgumentsPrefix = $"conda:{WslCondaEnvironmentName}";
                    _pythonVersion = wslVersion;
                    _pythonRuntimeKind = PythonRuntimeKind.WslConda;
                    _pythonAvailable = true;
                    _logger.LogInformation($"检测到 WSL Conda Python 环境: {_pythonArgumentsPrefix} - Python {_pythonVersion}");
                    return true;
                }

                var candidates = new (string FileName, string ArgumentPrefix)[]
                {
                    (Path.Combine(_baseDirectory, "python", "python.exe"), string.Empty),
                    (Path.Combine(_baseDirectory, "python", "Scripts", "python.exe"), string.Empty),
                    ("py", "-3.12"),
                    ("py", "-3.11"),
                    ("py", "-3.10"),
                    ("python3.12", string.Empty),
                    ("python3.11", string.Empty),
                    ("python3.10", string.Empty),
                    ("python", string.Empty),
                    ("python3", string.Empty),
                    ("py", string.Empty)
                };

                var detected = new List<(string FileName, string ArgumentPrefix, string Version)>();

                foreach (var candidate in candidates)
                {
                    var version = await TryGetPythonVersionAsync(candidate.FileName, candidate.ArgumentPrefix);
                    if (!string.IsNullOrWhiteSpace(version))
                    {
                        detected.Add((candidate.FileName, candidate.ArgumentPrefix, version));
                    }
                }

                var selected = detected
                    .FirstOrDefault(item => PreferredPythonVersions.Any(version => item.Version.StartsWith(version, StringComparison.OrdinalIgnoreCase)));

                if (string.IsNullOrWhiteSpace(selected.FileName) && detected.Count > 0)
                {
                    selected = detected[0];
                }

                if (!string.IsNullOrWhiteSpace(selected.FileName))
                {
                    _pythonPath = selected.FileName;
                    _pythonArgumentsPrefix = selected.ArgumentPrefix;
                    _pythonVersion = selected.Version;
                    _pythonRuntimeKind = PythonRuntimeKind.Standard;
                    _pythonAvailable = true;
                    _logger.LogInformation($"检测到Python环境: {FormatPythonCommand(_pythonPath, _pythonArgumentsPrefix)} - Python {_pythonVersion}");
                    return true;
                }

                _pythonAvailable = false;
                _logger.LogWarning("未检测到可用的Python环境");
            }
            catch (Exception ex)
            {
                _pythonAvailable = false;
                _logger.LogError(ex, "检测Python环境时发生异常");
            }

            return _pythonAvailable.Value;
        }

        /// <summary>
        /// 获取Python路径
        /// </summary>
        public string? PythonPath => _pythonPath;

        public string LocalPythonPackagesPath => _localPythonPackagesPath;

        /// <summary>
        /// 检查转换脚本是否存在
        /// </summary>
        public bool ConverterScriptExists => File.Exists(_converterScriptPath);

        /// <summary>
        /// 转换模型
        /// </summary>
        /// <param name="model">要转换的模型信息</param>
        /// <returns>转换后的模型路径，失败返回null</returns>
        public async Task<string?> ConvertAsync(ModelInfo model)
        {
            if (model.Format != ModelFormat.PyTorch)
            {
                _logger.LogWarning($"模型 {model.ModelName} 不是PyTorch格式，无需转换");
                return model.FilePath;
            }

            var outputPath = GetOutputPath(model.FilePath);

            if (File.Exists(outputPath))
            {
                _logger.LogInformation($"转换后的模型已存在: {outputPath}");
                model.ConvertedFilePath = outputPath;
                model.FilePath = outputPath;
                model.State = ModelState.Available;
                model.Format = ModelFormat.SafeTensors;
                return outputPath;
            }

            if (!await CheckPythonEnvironmentAsync())
            {
                model.State = ModelState.ConversionFailed;
                model.ErrorMessage = "Python环境不可用（convert_safetensors.py转换需要Python）";
                _logger.LogError(model.ErrorMessage);
                return null;
            }

            if (!await EnsurePythonDependenciesAsync(model))
            {
                model.State = ModelState.ConversionFailed;
                _logger.LogError(model.ErrorMessage);
                return null;
            }

            model.State = ModelState.Converting;
            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
            {
                ModelName = model.ModelName,
                Progress = 0,
                Status = "使用 convert_safetensors.py 转换（C++加载器将自动应用运行时转换）..."
            });

            try
            {
                var success = await RunConvertSafetensorsAsync(model.FilePath, outputPath, model);

                if (success)
                {
                    model.ConvertedFilePath = outputPath;
                    model.FilePath = outputPath;
                    model.State = ModelState.Available;
                    model.Format = ModelFormat.SafeTensors;

                    if (File.Exists(outputPath))
                    {
                        model.FileSizeBytes = new FileInfo(outputPath).Length;
                    }

                    _logger.LogInformation($"转换成功: {outputPath}");
                    return outputPath;
                }
                else
                {
                    model.State = ModelState.ConversionFailed;
                    model.ErrorMessage = "转换进程返回失败";
                    return null;
                }
            }
            catch (Exception ex)
            {
                model.State = ModelState.ConversionFailed;
                model.ErrorMessage = $"转换异常: {ex.Message}";
                _logger.LogError(ex, $"转换模型 {model.ModelName} 时发生异常");
                return null;
            }
        }

        /// <summary>
        /// 运行 convert_safetensors.py 转换
        /// </summary>
        private async Task<bool> RunConvertSafetensorsAsync(string inputPath, string outputPath, ModelInfo model)
        {
            if (_pythonRuntimeKind == PythonRuntimeKind.WslConda)
            {
                return await RunWslConversionAsync(inputPath, outputPath, model);
            }

            var tcs = new TaskCompletionSource<bool>();

            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonPath,
                Arguments = BuildPythonArguments($"\"{_converterScriptPath}\" --input \"{inputPath}\" --output \"{outputPath}\""),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(_converterScriptPath)
            };

            startInfo.Environment["PYTHONPATH"] = BuildPythonPathEnvironment();
            startInfo.Environment["PIP_DISABLE_PIP_VERSION_CHECK"] = "1";

            var process = new Process { StartInfo = startInfo };
            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            process.OutputDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    outputBuilder.AppendLine(e.Data);
                    _logger.LogDebug($"[转换输出] {e.Data}");
                    TryParseProgress(e.Data, model.ModelName);
                }
            };

            process.ErrorDataReceived += (s, e) =>
            {
                if (!string.IsNullOrEmpty(e.Data))
                {
                    errorBuilder.AppendLine(e.Data);
                    _logger.LogWarning($"[转换警告] {e.Data}");
                }
            };

            process.Exited += (s, e) =>
            {
                tcs.TrySetResult(process.ExitCode == 0 && File.Exists(outputPath));
                process.Dispose();
            };

            process.EnableRaisingEvents = true;

            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
            {
                ModelName = model.ModelName,
                Progress = 10,
                Status = "正在加载模型..."
            });

            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();

            var timeoutTask = Task.Delay(TimeSpan.FromMinutes(30));
            var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

            if (completedTask == timeoutTask)
            {
                try
                {
                    process.Kill();
                }
                catch { }

                _logger.LogError("模型转换超时");
                return false;
            }

            var success = await tcs.Task;
            if (!success)
            {
                var error = errorBuilder.ToString();
                var output = outputBuilder.ToString();
                var message = string.IsNullOrWhiteSpace(error) ? output : error;
                model.ErrorMessage = $"转换失败: {message}".Trim();
                _logger.LogError(model.ErrorMessage);
            }

            return success;
        }

        private async Task<bool> RunWslConversionAsync(string inputPath, string outputPath, ModelInfo model)
        {
            var wslScriptPath = ConvertWindowsPathToWsl(_converterScriptPath);
            var wslInputPath = ConvertWindowsPathToWsl(inputPath);
            var wslOutputPath = ConvertWindowsPathToWsl(outputPath);
            var command = $"if command -v {WslCondaExecutablePath} >/dev/null 2>&1; then {WslCondaExecutablePath} run -n {WslCondaEnvironmentName} python '{EscapeForBashSingleQuotedString(wslScriptPath)}' --input '{EscapeForBashSingleQuotedString(wslInputPath)}' --output '{EscapeForBashSingleQuotedString(wslOutputPath)}'; elif command -v conda >/dev/null 2>&1; then conda run -n {WslCondaEnvironmentName} python '{EscapeForBashSingleQuotedString(wslScriptPath)}' --input '{EscapeForBashSingleQuotedString(wslInputPath)}' --output '{EscapeForBashSingleQuotedString(wslOutputPath)}'; else echo 'conda not found' >&2; exit 1; fi";
            var result = await RunWslCommandAsync(command, timeoutMinutes: 30);

            if (result.Success && File.Exists(outputPath))
            {
                return true;
            }

            var message = string.IsNullOrWhiteSpace(result.Error) ? result.Output : result.Error;
            model.ErrorMessage = $"WSL Python 转换失败: {message}".Trim();
            _logger.LogError(model.ErrorMessage);
            return false;
        }

        /// <summary>
        /// 尝试解析转换进度
        /// </summary>
        private void TryParseProgress(string line, string modelName)
        {
            if (line.Contains("loading") || line.Contains("Loading"))
            {
                ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
                {
                    ModelName = modelName,
                    Progress = 20,
                    Status = "加载模型权重..."
                });
            }
            else if (line.Contains("converting") || line.Contains("Converting"))
            {
                ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
                {
                    ModelName = modelName,
                    Progress = 50,
                    Status = "转换权重格式..."
                });
            }
            else if (line.Contains("saving") || line.Contains("Saving"))
            {
                ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
                {
                    ModelName = modelName,
                    Progress = 80,
                    Status = "保存转换结果..."
                });
            }
        }

        /// <summary>
        /// 获取输出路径
        /// </summary>
        private string GetOutputPath(string inputPath)
        {
            var dir = Path.GetDirectoryName(inputPath) ?? "";
            var name = Path.GetFileNameWithoutExtension(inputPath);
            return Path.Combine(dir, $"{name}.st");
        }

        /// <summary>
        /// 取消转换（如果正在进行）
        /// </summary>
        public void CancelConversion()
        {
        }

        private async Task<bool> EnsurePythonDependenciesAsync(ModelInfo model)
        {
            if (_pythonRuntimeKind != PythonRuntimeKind.WslConda)
            {
                Directory.CreateDirectory(_localPythonPackagesPath);
            }

            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
            {
                ModelName = model.ModelName,
                Progress = 5,
                Status = "检测 Python 依赖..."
            });

            var missingPackages = await GetMissingPackagesAsync();
            if (missingPackages.Count == 0)
            {
                _logger.LogInformation($"Python依赖已就绪: {GetPythonRuntimeDisplayName()}");
                return true;
            }

            _logger.LogInformation($"检测到缺少Python依赖: {string.Join(", ", missingPackages)}");
            ProgressChanged?.Invoke(this, new ConversionProgressEventArgs
            {
                ModelName = model.ModelName,
                Progress = 8,
                Status = $"安装依赖: {string.Join(", ", missingPackages)}"
            });

            if (!await EnsurePipAvailableAsync())
            {
                model.ErrorMessage = "pip 不可用，无法自动安装模型转换依赖";
                return false;
            }

            var installResult = await InstallPackagesAsync(missingPackages);
            if (!installResult.Success)
            {
                var errorText = string.IsNullOrWhiteSpace(installResult.Error) ? installResult.Output : installResult.Error;
                if (!string.IsNullOrWhiteSpace(_pythonVersion) &&
                    (_pythonVersion.StartsWith("3.14", StringComparison.OrdinalIgnoreCase) ||
                     _pythonVersion.StartsWith("3.13", StringComparison.OrdinalIgnoreCase)))
                {
                    model.ErrorMessage = $"自动安装依赖失败，当前 Python {_pythonVersion} 可能不兼容 torch。请安装 Python 3.10-3.12 后重试。{Environment.NewLine}{errorText}".Trim();
                }
                else
                {
                    model.ErrorMessage = $"自动安装依赖失败: {errorText}".Trim();
                }
                _logger.LogError(model.ErrorMessage);
                return false;
            }

            var stillMissingPackages = await GetMissingPackagesAsync();
            if (stillMissingPackages.Count > 0)
            {
                model.ErrorMessage = $"依赖安装后仍缺少: {string.Join(", ", stillMissingPackages)}";
                _logger.LogError(model.ErrorMessage);
                return false;
            }

            _logger.LogInformation($"Python依赖已安装到: {GetPythonRuntimeDisplayName()}");
            return true;
        }

        private static readonly (string ModuleName, string PackageName)[] RequiredPackages =
        {
            ("torch", "torch"),
            ("safetensors", "safetensors")
        };

        private async Task<List<string>> GetMissingPackagesAsync()
        {
            if (string.IsNullOrWhiteSpace(_pythonPath))
            {
                return RequiredPackages.Select(item => item.PackageName).ToList();
            }

            var missingPackages = new List<string>();

            foreach (var (moduleName, packageName) in RequiredPackages)
            {
                string checkScript;
                ProcessExecutionResult result;

                if (_pythonRuntimeKind == PythonRuntimeKind.WslConda)
                {
                    checkScript = $"import {moduleName}";
                    result = await RunWslCommandAsync($"python -c '{checkScript}'", timeoutMinutes: 1);
                }
                else
                {
                    var localPathLiteral = ToPythonStringLiteral(_localPythonPackagesPath);
                    checkScript = $"import sys;sys.path.insert(0,r{localPathLiteral});import {moduleName}";
                    result = await RunPythonCommandAsync($"-c \"{checkScript}\"", includeLocalPythonPath: false);
                }

                if (!result.Success)
                {
                    missingPackages.Add(packageName);
                }
            }

            return missingPackages;
        }

        private async Task<bool> EnsurePipAvailableAsync()
        {
            var pipVersionResult = await RunPythonCommandAsync("-m pip --version", includeLocalPythonPath: false);
            if (pipVersionResult.Success)
            {
                return true;
            }

            _logger.LogWarning("当前Python环境缺少pip，正在尝试启用 ensurepip");
            var ensurePipResult = await RunPythonCommandAsync("-m ensurepip --upgrade", includeLocalPythonPath: false, timeoutMinutes: 5);
            if (ensurePipResult.Success)
            {
                return true;
            }

            _logger.LogError($"ensurepip 执行失败: {ensurePipResult.Error}");
            return false;
        }

        private async Task<ProcessExecutionResult> InstallPackagesAsync(List<string> packages)
        {
            if (_pythonRuntimeKind == PythonRuntimeKind.WslConda)
            {
                var wslInstallArguments = $"-m pip install --upgrade --extra-index-url https://download.pytorch.org/whl/cpu {string.Join(" ", packages)}";
                return await RunPythonCommandAsync(wslInstallArguments, includeLocalPythonPath: false, timeoutMinutes: 30);
            }

            var installArguments = $"-m pip install --upgrade --target \"{_localPythonPackagesPath}\" --extra-index-url https://download.pytorch.org/whl/cpu {string.Join(" ", packages)}";
            return await RunPythonCommandAsync(installArguments, includeLocalPythonPath: false, timeoutMinutes: 30);
        }

        private async Task<string?> TryGetWslCondaPythonVersionAsync()
        {
            var command = $"if command -v {WslCondaExecutablePath} >/dev/null 2>&1; then {WslCondaExecutablePath} run -n {WslCondaEnvironmentName} python --version; elif command -v conda >/dev/null 2>&1; then conda run -n {WslCondaEnvironmentName} python --version; fi";
            var result = await RunWslCommandAsync(command, timeoutMinutes: 1);
            if (!result.Success)
            {
                return null;
            }

            var versionText = string.IsNullOrWhiteSpace(result.Output) ? result.Error : result.Output;
            if (string.IsNullOrWhiteSpace(versionText))
            {
                return null;
            }

            const string prefix = "Python ";
            return versionText.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? versionText.Substring(prefix.Length).Trim()
                : versionText.Trim();
        }

        private async Task<string?> TryGetPythonVersionAsync(string fileName, string argumentPrefix)
        {
            try
            {
                if (Path.IsPathRooted(fileName) && !File.Exists(fileName))
                {
                    return null;
                }

                var arguments = string.IsNullOrWhiteSpace(argumentPrefix)
                    ? "--version"
                    : $"{argumentPrefix} --version";

                var startInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true
                };

                using var process = Process.Start(startInfo);
                if (process == null)
                {
                    return null;
                }

                var waitTask = process.WaitForExitAsync();
                var completedTask = await Task.WhenAny(waitTask, Task.Delay(TimeSpan.FromSeconds(5)));
                if (completedTask != waitTask)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                    }

                    return null;
                }

                await waitTask;

                if (process.ExitCode != 0)
                {
                    return null;
                }

                var output = (await process.StandardOutput.ReadToEndAsync()).Trim();
                var error = (await process.StandardError.ReadToEndAsync()).Trim();
                var versionText = string.IsNullOrWhiteSpace(output) ? error : output;
                if (string.IsNullOrWhiteSpace(versionText))
                {
                    return null;
                }

                const string prefix = "Python ";
                return versionText.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                    ? versionText.Substring(prefix.Length).Trim()
                    : versionText.Trim();
            }
            catch
            {
                return null;
            }
        }

        private async Task<ProcessExecutionResult> RunPythonCommandAsync(string arguments, bool includeLocalPythonPath, int timeoutMinutes = 10)
        {
            if (string.IsNullOrWhiteSpace(_pythonPath))
            {
                return new ProcessExecutionResult(false, string.Empty, "Python环境不可用");
            }

            if (_pythonRuntimeKind == PythonRuntimeKind.WslConda)
            {
                return await RunWslPythonCommandAsync(arguments, timeoutMinutes);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonPath,
                Arguments = BuildPythonArguments(arguments),
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(_converterScriptPath)
            };

            if (includeLocalPythonPath)
            {
                startInfo.Environment["PYTHONPATH"] = BuildPythonPathEnvironment();
            }

            startInfo.Environment["PIP_DISABLE_PIP_VERSION_CHECK"] = "1";

            using var process = new Process { StartInfo = startInfo };
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();
            var waitTask = process.WaitForExitAsync();
            var completedTask = await Task.WhenAny(waitTask, Task.Delay(TimeSpan.FromMinutes(timeoutMinutes)));

            if (completedTask != waitTask)
            {
                try
                {
                    process.Kill(true);
                }
                catch
                {
                }

                return new ProcessExecutionResult(false, string.Empty, "执行超时");
            }

            await waitTask;
            var output = await outputTask;
            var error = await errorTask;

            return new ProcessExecutionResult(process.ExitCode == 0, output.Trim(), error.Trim());
        }

        private async Task<ProcessExecutionResult> RunWslPythonCommandAsync(string arguments, int timeoutMinutes)
        {
            var command = $"if command -v {WslCondaExecutablePath} >/dev/null 2>&1; then {WslCondaExecutablePath} run -n {WslCondaEnvironmentName} python {arguments}; elif command -v conda >/dev/null 2>&1; then conda run -n {WslCondaEnvironmentName} python {arguments}; else echo 'conda not found' >&2; exit 1; fi";
            return await RunWslCommandAsync(command, timeoutMinutes);
        }

        private async Task<ProcessExecutionResult> RunWslCommandAsync(string bashCommand, int timeoutMinutes)
        {
            var wslExecutablePath = GetWslExecutablePath();
            var startInfo = new ProcessStartInfo
            {
                FileName = wslExecutablePath,
                Arguments = $"bash -lc \"{EscapeForBashDoubleQuotedString(bashCommand)}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = Path.GetDirectoryName(_converterScriptPath)
            };

            using var process = new Process { StartInfo = startInfo };
            process.Start();

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();
            var waitTask = process.WaitForExitAsync();
            var completedTask = await Task.WhenAny(waitTask, Task.Delay(TimeSpan.FromMinutes(timeoutMinutes)));

            if (completedTask != waitTask)
            {
                try
                {
                    process.Kill(true);
                }
                catch
                {
                }

                return new ProcessExecutionResult(false, string.Empty, "WSL 执行超时");
            }

            await waitTask;
            var output = await outputTask;
            var error = await errorTask;
            return new ProcessExecutionResult(process.ExitCode == 0, output.Trim(), error.Trim());
        }

        private string BuildPythonArguments(string arguments)
        {
            return string.IsNullOrWhiteSpace(_pythonArgumentsPrefix)
                ? arguments
                : $"{_pythonArgumentsPrefix} {arguments}";
        }

        private string BuildPythonPathEnvironment()
        {
            var existingPythonPath = Environment.GetEnvironmentVariable("PYTHONPATH");
            return string.IsNullOrWhiteSpace(existingPythonPath)
                ? _localPythonPackagesPath
                : $"{_localPythonPackagesPath}{Path.PathSeparator}{existingPythonPath}";
        }

        private static string EscapeForBashDoubleQuotedString(string value)
        {
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("$", "\\$")
                .Replace("`", "\\`");
        }

        private static string EscapeForBashSingleQuotedString(string value)
        {
            return value.Replace("'", "'\"'\"'");
        }

        private static string ToPythonStringLiteral(string value)
        {
            return $"'{value.Replace("\\", "\\\\").Replace("'", "\\'")}'";
        }

        private static string FormatPythonCommand(string? fileName, string argumentPrefix)
        {
            return string.IsNullOrWhiteSpace(argumentPrefix) ? fileName ?? string.Empty : $"{fileName} {argumentPrefix}";
        }

        private static string ConvertWindowsPathToWsl(string path)
        {
            var fullPath = Path.GetFullPath(path);
            var root = Path.GetPathRoot(fullPath);
            if (string.IsNullOrWhiteSpace(root) || root.Length < 2 || root[1] != ':')
            {
                return fullPath.Replace('\\', '/');
            }

            var drive = char.ToLowerInvariant(root[0]);
            var relativePath = fullPath.Substring(root.Length).Replace('\\', '/');
            return $"/mnt/{drive}/{relativePath}";
        }

        private string GetPythonRuntimeDisplayName()
        {
            return _pythonRuntimeKind == PythonRuntimeKind.WslConda
                ? $"WSL conda 环境 {WslCondaEnvironmentName}"
                : _localPythonPackagesPath;
        }

        private static string GetWslExecutablePath()
        {
            var systemRoot = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            var candidatePaths = new[]
            {
                Path.Combine(systemRoot, "System32", "wsl.exe"),
                Path.Combine(systemRoot, "Sysnative", "wsl.exe"),
                "wsl.exe",
                "wsl"
            };

            foreach (var candidatePath in candidatePaths)
            {
                if (!Path.IsPathRooted(candidatePath) || File.Exists(candidatePath))
                {
                    return candidatePath;
                }
            }

            return "wsl.exe";
        }
    }

    /// <summary>
    /// 转换进度事件参数
    /// </summary>
    public class ConversionProgressEventArgs : EventArgs
    {
        public string ModelName { get; set; } = string.Empty;
        public int Progress { get; set; }
        public string Status { get; set; } = string.Empty;
    }

    internal sealed class ProcessExecutionResult
    {
        public ProcessExecutionResult(bool success, string output, string error)
        {
            Success = success;
            Output = output;
            Error = error;
        }

        public bool Success { get; }
        public string Output { get; }
        public string Error { get; }
    }

    internal enum PythonRuntimeKind
    {
        Standard,
        WslConda
    }
}
