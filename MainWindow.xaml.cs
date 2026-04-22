using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using DocumentTranslator.Services.Translation;
using DocumentTranslator.Services.Logging;
using DocumentTranslator.Services;
using DocumentTranslator.Services.RwkvLocal;
using DocumentTranslator.Services.RwkvLocal.Models;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using IOPath = System.IO.Path;
using IOFile = System.IO.File;
using IODirectory = System.IO.Directory;
using IOFileInfo = System.IO.FileInfo;
using DocumentTranslator.Helpers;

namespace DocumentTranslator
{
    public partial class MainWindow : Window
    {
        private Dictionary<string, string> _languageMap;
        private TranslationService _translationService;
        private DocumentProcessorFactory _documentProcessorFactory;
        private TermExtractor _termExtractor;
        private ConfigurationManager _configurationManager;
        private ILogger<MainWindow> _logger;
        private ILoggerFactory _loggerFactory;
        private TranslationOutputConfig _outputConfig = new TranslationOutputConfig();
        private string _currentEngine = "rwkv";
        private bool _isTranslating = false;
        private readonly Queue<string> _logMessages = new Queue<string>(40);
        private string _historyRecordPath;
        private readonly object _logLock = new object();
        private DispatcherTimer _uiUpdateTimer;
        private volatile bool _logsChanged = false;

        // 本地模型服务
        private GpuResourceService? _gpuResourceService;
        private ModelManagementService? _modelManagementService;
        private ConcurrencyCalculator? _concurrencyCalculator;
        private ModelBenchmarkCache? _benchmarkCache;
        private RwkvProcessService? _rwkvProcessService;
        private LlamaCppProcessService? _llamaCppProcessService;
        private ModelConverter? _modelConverter;
        private ModelDownloadService? _modelDownloadService;

        // 当前使用的推理工具名称
        private string _currentToolName = "rwkv_lightning";
        
        // 防止 SelectionChanged 事件递归调用
        private bool _suppressToolSelectionEvents = false;

        // GGUF模型默认下载地址
        private const string DefaultGgufModelScopeUrl = "https://www.modelscope.cn/models/shoumenchougou/RWKV7-G1e-1.5B-GGUF/files";
        // SafeTensors模型默认下载地址
        private const string DefaultStModelScopeUrl = "https://www.modelscope.cn/models/AlicLi/rwkv7-g1-libtorch-st/files";

        // 本地模型数据
        private List<GpuInfo> _availableGpus = new List<GpuInfo>();
        private List<ModelInfo> _availableModels = new List<ModelInfo>();

        // 下载相关
        private CancellationTokenSource? _downloadCts;

        public MainWindow()
        {
            InitializeComponent();

            // 安全加载窗口图标（避免XAML中Icon属性加载失败导致窗口无法显示）
            try
            {
                var iconPath = IOPath.Combine(AppDomain.CurrentDomain.BaseDirectory, "logo.ico");
                if (IOFile.Exists(iconPath))
                {
                    using var stream = IOFile.OpenRead(iconPath);
                    this.Icon = BitmapFrame.Create(stream);
                }
            }
            catch { }

            // 设置窗口标题
            this.Title = "RWKV文档术语翻译助手 v3.1";

            // 初始化UI更新定时器
            _uiUpdateTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            _uiUpdateTimer.Tick += (s, e) => UpdateLogUI();
            _uiUpdateTimer.Start();

            // 绑定事件处理器（轻量级，不阻塞UI）
            ChineseToForeign.Checked += OnTranslationDirectionChanged;
            ForeignToChinese.Checked += OnTranslationDirectionChanged;
            LanguageCombo.SelectionChanged += OnLanguageSelectionChanged;
            BilingualOutput.Checked += OnOutputFormatChanged;
            TranslationOnlyOutput.Checked += OnOutputFormatChanged;

            // 初始化并发建议与滑块
            try
            {
                SuggestedParallelText.Text = "4";
                ParallelismSlider.Value = 4;
                ParallelismValueText.Text = $"并发: 4";
                ParallelismInput.Text = "4";
            }
            catch { }

            ParallelismSlider.ValueChanged += (s, e) =>
            {
                int value = (int)Math.Round(ParallelismSlider.Value);
                ParallelismValueText.Text = $"并发: {value}";
                ParallelismInput.Text = value.ToString();
            };

            ParallelismInput.TextChanged += (s, e) =>
            {
                if (int.TryParse(ParallelismInput.Text, out int value))
                {
                    value = Math.Max((int)ParallelismSlider.Minimum, Math.Min((int)ParallelismSlider.Maximum, value));
                    ParallelismSlider.Value = value;
                    ParallelismValueText.Text = $"并发: {value}";
                }
            };

            // 窗口已准备好显示，在窗口 Loaded 后进行重型初始化
            this.Loaded += async (s, e) =>
            {
                await InitializeHeavyServicesAsync();
            };
        }

        /// <summary>
        /// 异步初始化重型服务（DI容器、翻译服务、GPU服务等）
        /// 在后台线程执行，不阻塞窗口显示
        /// </summary>
        private async Task InitializeHeavyServicesAsync()
        {
            try
            {
                await Task.Run(async () =>
                {
                    _historyRecordPath = IOPath.Combine(PathHelper.GetSafeBaseDirectory(), "History_record");
                    if (!IODirectory.Exists(_historyRecordPath))
                    {
                        IODirectory.CreateDirectory(_historyRecordPath);
                    }

                    // 初始化依赖注入容器（这是最耗时的操作）
                    var services = new ServiceCollection();
                    ConfigureServices(services);
                    var serviceProvider = services.BuildServiceProvider();

                    // 获取服务实例
                    _translationService = serviceProvider.GetRequiredService<TranslationService>();
                    _documentProcessorFactory = serviceProvider.GetRequiredService<DocumentProcessorFactory>();
                    _termExtractor = serviceProvider.GetRequiredService<TermExtractor>();
                    _configurationManager = serviceProvider.GetRequiredService<ConfigurationManager>();
                    _logger = serviceProvider.GetRequiredService<ILogger<MainWindow>>();
                    _loggerFactory = serviceProvider.GetRequiredService<ILoggerFactory>();

                    // 初始化语言映射
                    _languageMap = new Dictionary<string, string>
                    {
                        {"英语", "en"}, {"日语", "ja"}, {"韩语", "ko"}, {"法语", "fr"},
                        {"德语", "de"}, {"西班牙语", "es"}, {"越南语", "vi"}, {"意大利语", "it"},
                        {"俄语", "ru"}, {"葡萄牙语", "pt"}, {"荷兰语", "nl"}, {"阿拉伯语", "ar"},
                        {"泰语", "th"}, {"印尼语", "id"}, {"马来语", "ms"}, {"土耳其语", "tr"},
                        {"波兰语", "pl"}, {"捷克语", "cs"}, {"匈牙利语", "hu"}, {"希腊语", "el"},
                        {"瑞典语", "sv"}, {"挪威语", "no"}, {"丹麦语", "da"}, {"芬兰语", "fi"}
                    };
                });

                // 在UI线程更新UI相关部分
                await Dispatcher.InvokeAsync(() =>
                {
                    UpdateLanguageMappingFromTerminology();
                    UpdateSuggestedParallelism();
                    LogMessage("🟢 系统初始化完成");
                });

                // 异步初始化本地模型服务
                await InitializeLocalModelServicesAsync();
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "系统初始化失败");
                await Dispatcher.InvokeAsync(() => LogMessage($"⚠️ 系统初始化失败: {ex.Message}"));
            }
        }

        /// <summary>
        /// 异步初始化本地模型服务
        /// </summary>
        private async Task InitializeLocalModelServicesAsync()
        {
            try
            {
                // 在后台线程初始化GPU服务（NVML加载可能较慢）
                await Task.Run(() =>
                {
                    _gpuResourceService = new GpuResourceService(_logger);
                    _modelManagementService = new ModelManagementService(_logger);
                    _benchmarkCache = new ModelBenchmarkCache(_logger);
                    _concurrencyCalculator = new ConcurrencyCalculator(_logger, _gpuResourceService, _modelManagementService, _benchmarkCache);
                    _rwkvProcessService = new RwkvProcessService(_logger, _modelManagementService, _gpuResourceService, _concurrencyCalculator);
                    _llamaCppProcessService = new LlamaCppProcessService(_logger, _modelManagementService, _gpuResourceService, _concurrencyCalculator);
                    _modelConverter = new ModelConverter(_logger);
                    _modelDownloadService = new ModelDownloadService(_logger);
                });

                // 订阅服务状态变化事件
                _rwkvProcessService.StatusChanged += OnRwkvServiceStatusChanged;
                _rwkvProcessService.BenchmarkOutputReceived += OnBenchmarkOutputReceived;
                _llamaCppProcessService.StatusChanged += OnLlamaCppServiceStatusChanged;
                _modelConverter.ProgressChanged += OnModelConversionProgress;
                _modelDownloadService.StatusChanged += OnDownloadStatusChanged;
                _modelDownloadService.ProgressChanged += OnDownloadProgressChanged;
                _modelDownloadService.DownloadCompleted += OnDownloadCompleted;

                LogMessage("🎮 本地模型服务初始化完成");

                // 异步初始化GPU和模型列表
                await InitializeGpuAndModelsAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化本地模型服务失败");
                LogMessage($"⚠️ 本地模型服务初始化失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 初始化本地模型服务（保留同步版本用于兼容）
        /// </summary>
        private void InitializeLocalModelServices()
        {
            try
            {
                // 创建服务实例
                _gpuResourceService = new GpuResourceService(_logger);
                _modelManagementService = new ModelManagementService(_logger);
                _benchmarkCache = new ModelBenchmarkCache(_logger);
                _concurrencyCalculator = new ConcurrencyCalculator(_logger, _gpuResourceService, _modelManagementService, _benchmarkCache);
                _rwkvProcessService = new RwkvProcessService(_logger, _modelManagementService, _gpuResourceService, _concurrencyCalculator);
                _llamaCppProcessService = new LlamaCppProcessService(_logger, _modelManagementService, _gpuResourceService, _concurrencyCalculator);
                _modelConverter = new ModelConverter(_logger);
                _modelDownloadService = new ModelDownloadService(_logger);

                // 订阅服务状态变化事件
                _rwkvProcessService.StatusChanged += OnRwkvServiceStatusChanged;
                _rwkvProcessService.BenchmarkOutputReceived += OnBenchmarkOutputReceived;
                _llamaCppProcessService.StatusChanged += OnLlamaCppServiceStatusChanged;
                _modelConverter.ProgressChanged += OnModelConversionProgress;
                _modelDownloadService.StatusChanged += OnDownloadStatusChanged;
                _modelDownloadService.ProgressChanged += OnDownloadProgressChanged;
                _modelDownloadService.DownloadCompleted += OnDownloadCompleted;

                // 异步初始化GPU和模型列表
                _ = InitializeGpuAndModelsAsync();

                LogMessage("🎮 本地模型服务初始化完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化本地模型服务失败");
                LogMessage($"⚠️ 本地模型服务初始化失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 异步初始化GPU和模型列表
        /// </summary>
        private async Task InitializeGpuAndModelsAsync()
        {
            try
            {
                // 检测GPU
                _availableGpus = await _gpuResourceService!.GetAllGpusAsync();
                
                await Dispatcher.InvokeAsync(() =>
                {
                    // 抑制事件，防止初始化期间触发递归
                    _suppressToolSelectionEvents = true;
                    try
                    {
                        GpuSelectorCombo.Items.Clear();
                        if (_availableGpus.Count > 0)
                        {
                            foreach (var gpu in _availableGpus)
                            {
                                GpuSelectorCombo.Items.Add(new ComboBoxItem 
                                { 
                                    Content = gpu.DisplayName, 
                                    Tag = gpu 
                                });
                            }
                            GpuSelectorCombo.SelectedIndex = 0;
                            LogMessage($"🎮 检测到 {_availableGpus.Count} 个NVIDIA GPU");
                            
                            // 显示GPU显存信息
                            var firstGpu = _availableGpus[0];
                            GpuMemoryText.Text = $"显存: {firstGpu.UsedMemoryGB:F2}/{firstGpu.TotalMemoryGB:F1}GB";
                        }
                        else
                        {
                            GpuSelectorCombo.Items.Add(new ComboBoxItem { Content = "⚠️ 未检测到NVIDIA GPU" });
                            GpuSelectorCombo.SelectedIndex = 0;
                            LogMessage("⚠️ 未检测到NVIDIA GPU，已自动切换到CPU模式");
                            // 无GPU时自动切换到CPU模式
                            ComputeModeCombo.SelectedIndex = 1; // CPU
                        }
                        
                        // 初始化推理工具列表
                        RefreshInferenceTools();
                    }
                    finally
                    {
                        _suppressToolSelectionEvents = false;
                    }
                });

                // 扫描模型
                _availableModels = await _modelManagementService!.ScanModelsAsync();

                // 根据当前推理工具类型过滤模型格式
                var initIsLlamaCpp = _currentToolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";
                var initFilteredModels = _availableModels.Where(m =>
                    initIsLlamaCpp ? m.Format == ModelFormat.GGUF : m.Format != ModelFormat.GGUF
                ).ToList();
                
                await Dispatcher.InvokeAsync(() =>
                {
                    ModelSelectorCombo.Items.Clear();
                    if (initFilteredModels.Count > 0)
                    {
                        foreach (var model in initFilteredModels)
                        {
                            ModelSelectorCombo.Items.Add(new ComboBoxItem 
                            { 
                                Content = model.DisplayName, 
                                Tag = model 
                            });
                        }
                        ModelSelectorCombo.SelectedIndex = 0;
                        LogMessage($"📦 扫描到 {initFilteredModels.Count} 个模型文件");
                    }
                    else
                    {
                        ModelSelectorCombo.Items.Add(new ComboBoxItem { Content = "⚠️ 未找到模型文件" });
                        ModelSelectorCombo.SelectedIndex = 0;
                        LogMessage("⚠️ 未在rwkv_models目录找到模型文件");
                        LocalModelInfoText.Text = initIsLlamaCpp
                            ? "请将 .gguf 模型文件放入 rwkv_models 目录"
                            : "请将 .safetensors 或 .pth 模型文件放入 rwkv_models 目录";
                    }

                    // 检查推理工具
                    var rwkvToolExists = _modelManagementService.ToolExists("rwkv_lightning");
                    var benchmarkToolExists = _modelManagementService.ToolExists("benchmark");
                    var llamaCudaExists = _modelManagementService.ToolExists("llama-cuda");
                    var llamaCpuExists = _modelManagementService.ToolExists("llama-cpu");
                    var llamaSyclExists = _modelManagementService.ToolExists("llama-sycl");
                    var llamaVulkanExists = _modelManagementService.ToolExists("llama-vulkan");
                    
                    var toolStatus = new List<string>();
                    if (rwkvToolExists) toolStatus.Add("rwkv_lightning.exe✅");
                    if (benchmarkToolExists) toolStatus.Add("benchmark.exe✅");
                    if (llamaCudaExists) toolStatus.Add("llama-server[CUDA]✅");
                    if (llamaSyclExists) toolStatus.Add("llama-server[SYCL]✅");
                    if (llamaVulkanExists) toolStatus.Add("llama-server[Vulkan]✅");
                    if (llamaCpuExists) toolStatus.Add("llama-server[CPU]✅");

                    if (toolStatus.Count == 0)
                    {
                        LogMessage("⚠️ 未找到任何推理工具");
                        LocalModelInfoText.Text = "⚠️ 推理工具不存在";
                    }
                    else if (!rwkvToolExists && !benchmarkToolExists && !_modelManagementService.VocabFileExists)
                    {
                        // RWKV工具不存在但llama.cpp可能存在
                        LogMessage($"🔧 可用推理工具: {string.Join(", ", toolStatus)}");
                        LogMessage("⚠️ RWKV词汇表文件不存在，RWKV推理工具不可用");
                    }
                    else
                    {
                        LogMessage($"🔧 可用推理工具: {string.Join(", ", toolStatus)}");
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化GPU和模型列表失败");
                await Dispatcher.InvokeAsync(() =>
                {
                    LogMessage($"❌ 初始化GPU和模型列表失败: {ex.Message}");
                });
            }
        }

        /// <summary>
        /// RWKV服务状态变化事件处理
        /// </summary>
        private void OnRwkvServiceStatusChanged(object? sender, ServiceStatusEventArgs e)
        {
            Dispatcher.BeginInvoke(() =>
            {
                UpdateLocalServiceStatusUI(e.Status);
            });
        }

        /// <summary>
        /// llama.cpp 服务状态变化事件处理
        /// </summary>
        private void OnLlamaCppServiceStatusChanged(object? sender, ServiceStatusEventArgs e)
        {
            Dispatcher.BeginInvoke(() =>
            {
                UpdateLocalServiceStatusUI(e.Status);
            });
        }

        /// <summary>
        /// Benchmark测试结果输出事件处理
        /// </summary>
        private void OnBenchmarkOutputReceived(object? sender, string output)
        {
            Dispatcher.BeginInvoke(() =>
            {
                LogMessage("📊 ===== Benchmark测试结果 =====");
                // 按行分割输出并逐行显示
                var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var line in lines)
                {
                    LogMessage($"📊 {line}");
                }
                LogMessage("📊 ===== 测试完成 =====");
            });
        }

        /// <summary>
        /// 模型转换进度事件处理
        /// </summary>
        private void OnModelConversionProgress(object? sender, ConversionProgressEventArgs e)
        {
            Dispatcher.BeginInvoke(() =>
            {
                LocalModelInfoText.Text = $"🔄 转换中: {e.ModelName} - {e.Status} ({e.Progress}%)";
            });
        }

        /// <summary>
        /// 更新本地服务状态UI
        /// </summary>
        private void UpdateLocalServiceStatusUI(RwkvServiceStatus status)
        {
            LocalServiceStatusText.Text = status.StateDisplayText;
            LocalServiceStatusText.Foreground = new System.Windows.Media.SolidColorBrush(
                (System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(status.StateColor));

            StartLocalModelBtn.IsEnabled = status.CanStart;
            StopLocalModelBtn.IsEnabled = status.CanStop;

            if (status.State == ServiceState.Running)
            {
                SuggestedConcurrencyText.Text = status.MaxConcurrency.ToString();
                LocalModelInfoText.Text = $"✅ 运行中 - 端口:{status.Port}, 并发:{status.MaxConcurrency}";
                
                if (status.CurrentGpuStatus != null)
                {
                    GpuMemoryText.Text = $"显存: {status.CurrentGpuStatus.UsedMemoryGB:F2}/{status.CurrentGpuStatus.TotalMemoryGB:F1}GB";
                }

                // 更新API地址显示（根据推理工具类型使用不同的API路径）
                var isLlamaCpp = _currentToolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";
                ApiUrlDisplay.Text = isLlamaCpp 
                    ? status.ApiUrl + "/v1/chat/completions" 
                    : status.ApiUrl + "/translate/v1/batch-translate";
            }
            else if (status.State == ServiceState.Error)
            {
                LocalModelInfoText.Text = $"❌ {status.ErrorMessage ?? "未知错误"}";
            }
        }

        /// <summary>
        /// 启动本地模型
        /// </summary>
        private async void StartLocalModel(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取选择的推理工具
                var selectedToolItem = InferenceToolCombo.SelectedItem as ComboBoxItem;
                var toolName = selectedToolItem?.Tag as string ?? ResolveToolName(selectedToolItem?.Content as string ?? "");
                _currentToolName = toolName;
                var isLlamaCpp = toolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";

                var selectedModel = GetSelectedModel();

                if (selectedModel == null)
                {
                    LogMessage("❌ 未选择模型");
                    MessageBox.Show("请选择一个模型", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // llama.cpp 工具只支持 GGUF 格式
                if (isLlamaCpp && selectedModel.Format != ModelFormat.GGUF)
                {
                    MessageBox.Show(
                        "llama.cpp 推理工具仅支持 GGUF 格式模型。\n\n请选择 .gguf 格式的模型，或切换到 rwkv_lightning 推理工具。",
                        "模型格式不匹配", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // rwkv_lightning/benchmark 不支持 GGUF 格式
                if (!isLlamaCpp && selectedModel.Format == ModelFormat.GGUF)
                {
                    MessageBox.Show(
                        "RWKV 推理工具不支持 GGUF 格式模型。\n\n请切换到 llama-server 推理工具来使用 GGUF 模型。",
                        "模型格式不匹配", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 获取选中的GPU（仅GPU模式需要）
                GpuInfo? selectedGpu = null;
                var isGpuMode = IsGpuComputeMode();
                
                if (isGpuMode)
                {
                    var selectedGpuItem = GpuSelectorCombo.SelectedItem as ComboBoxItem;
                    selectedGpu = selectedGpuItem?.Tag as GpuInfo;
                    
                    if (selectedGpu == null && _availableGpus.Count > 0)
                    {
                        selectedGpu = _availableGpus[0];
                    }
                    
                    if (selectedGpu == null)
                    {
                        LogMessage("❌ 未选择GPU");
                        MessageBox.Show("请选择一个GPU", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                // 检查模型是否需要转换
                if (selectedModel.NeedsConversion)
                {
                    var result = MessageBox.Show(
                        $"模型 '{selectedModel.ModelName}' 需要转换为safetensors格式。\n\n是否现在转换？（需要Python环境）",
                        "模型转换",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        selectedModel = await ConvertModelAsync(selectedModel);
                        if (selectedModel == null) return;
                    }
                    else
                    {
                        return;
                    }
                }

                LogMessage($"🚀 正在启动本地模型: {selectedModel.ModelName} (工具: {toolName})");
                LocalModelInfoText.Text = "🟡 正在启动服务...";

                if (isLlamaCpp)
                {
                    // 使用 llama.cpp 推理服务
                    if (_llamaCppProcessService == null)
                    {
                        LogMessage("❌ llama.cpp 服务未初始化");
                        return;
                    }

                    // 读取配置的对话模板
                    var chatTemplate = GetLlamaCppChatTemplate();
                    
                    // 获取 BSZ 测试选项和增量值
                    bool enableBszTest = EnableBszTestCheck.IsChecked == true;
                    int bszIncrement = 10;
                    if (int.TryParse(BszIncrementText.Text, out int parsedIncrement) && parsedIncrement > 0)
                    {
                        bszIncrement = parsedIncrement;
                    }

                    if (enableBszTest)
                    {
                        LogMessage($"🧪 已启用 BSZ 上限测试，增量值: {bszIncrement}");
                    }

                    var success = await _llamaCppProcessService.StartAsync(selectedModel, selectedGpu, toolName, chatTemplate: chatTemplate, enableBszTest: enableBszTest, bszIncrement: bszIncrement);

                    if (success)
                    {
                        var port = _llamaCppProcessService.Status.Port;
                        var concurrency = _llamaCppProcessService.Status.MaxConcurrency;
                        LogMessage($"✅ llama-server 启动成功 - 端口: {port}, 并发: {concurrency}");

                        // 切换到 llama_cpp 翻译器，使用续写模式（与 rwkv_lightning 翻译端点等效）
                        var llamaApiUrl = $"http://127.0.0.1:{port}/v1/completions";
                        UpdateLlamaCppApiUrl(llamaApiUrl);
                        _translationService.ReinitializeTranslator("llama_cpp");
                        _translationService.CurrentTranslatorType = "llama_cpp";
                        _currentEngine = "llama_cpp";
                        LogMessage($"🔗 已切换到 llama_cpp 翻译器，API: {llamaApiUrl}");

                        _translationService.SetSuggestedMaxParallelism(concurrency, "llama_cpp");
                        UpdateSuggestedParallelism();
                        LogMessage($"⚡ 已自动应用并发设置: {concurrency}");
                    }
                    else
                    {
                        LogMessage($"❌ llama-server 启动失败: {_llamaCppProcessService.Status.ErrorMessage}");
                    }
                }
                else
                {
                    // 使用 RWKV 推理服务
                    if (_rwkvProcessService == null)
                    {
                        LogMessage("❌ 本地模型服务未初始化");
                        return;
                    }

                    // 获取 BSZ 测试选项和增量值
                    bool enableBszTest = EnableBszTestCheck.IsChecked == true;
                    int bszIncrement = 10;
                    if (int.TryParse(BszIncrementText.Text, out int parsedIncrement) && parsedIncrement > 0)
                    {
                        bszIncrement = parsedIncrement;
                    }

                    if (enableBszTest)
                    {
                        LogMessage($"🧪 已启用 BSZ 上限测试，增量值: {bszIncrement}");
                    }

                    var success = await _rwkvProcessService.StartAsync(selectedModel, selectedGpu, toolName: toolName, enableBszTest: enableBszTest, bszIncrement: bszIncrement);

                    if (success)
                    {
                        LogMessage($"✅ 本地模型启动成功 - 端口: {_rwkvProcessService.Status.Port}, 并发: {_rwkvProcessService.Status.MaxConcurrency}");

                        // 切换到 rwkv 翻译器
                        var rwkvApiUrl = $"http://127.0.0.1:{_rwkvProcessService.Status.Port}/translate/v1/batch-translate";
                        UpdateRwkvApiUrl(rwkvApiUrl);
                        _translationService.ReinitializeTranslator("rwkv");
                        _translationService.CurrentTranslatorType = "rwkv";
                        _currentEngine = "rwkv";
                        LogMessage($"🔗 已切换到 rwkv 翻译器，API: {rwkvApiUrl}");

                        _translationService.SetSuggestedMaxParallelism(_rwkvProcessService.Status.MaxConcurrency, "rwkv");
                        UpdateSuggestedParallelism();
                        LogMessage($"⚡ 已自动应用计算并发设置: {_rwkvProcessService.Status.MaxConcurrency}");
                    }
                    else
                    {
                        LogMessage($"❌ 本地模型启动失败: {_rwkvProcessService.Status.ErrorMessage}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "启动本地模型失败");
                LogMessage($"❌ 启动本地模型失败: {ex.Message}");
                MessageBox.Show($"启动本地模型失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 停止本地模型
        /// </summary>
        private async void StopLocalModel(object sender, RoutedEventArgs e)
        {
            try
            {
                var isLlamaCpp = _currentToolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";

                if (isLlamaCpp && _llamaCppProcessService != null)
                {
                    LogMessage("⏹️ 正在停止 llama-server...");
                    await _llamaCppProcessService.StopAsync();
                    LogMessage("✅ llama-server 已停止");
                }
                else if (_rwkvProcessService != null)
                {
                    LogMessage("⏹️ 正在停止本地模型...");
                    await _rwkvProcessService.StopAsync();
                    LogMessage("✅ 本地模型已停止");
                }
                
                // 恢复API地址显示
                UpdateApiUrlDisplay();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "停止本地模型失败");
                LogMessage($"❌ 停止本地模型失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 刷新本地模型列表
        /// </summary>
        private async void RefreshLocalModels(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("🔄 正在刷新GPU和模型列表...");
                await InitializeGpuAndModelsAsync();
                LogMessage("✅ 刷新完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "刷新本地模型列表失败");
                LogMessage($"❌ 刷新失败: {ex.Message}");
            }
        }

        private ModelInfo? GetSelectedModel()
        {
            var selectedModelItem = ModelSelectorCombo.SelectedItem as ComboBoxItem;
            var selectedModel = selectedModelItem?.Tag as ModelInfo;

            if (selectedModel == null && _availableModels.Count > 0)
            {
                selectedModel = _availableModels[0];
            }

            return selectedModel;
        }

        private async Task RefreshModelSelectorAsync(string? preferredModelPath = null)
        {
            if (_modelManagementService == null) return;

            _availableModels = await _modelManagementService.ScanModelsAsync();

            // 根据当前推理工具类型过滤模型格式
            var isLlamaCpp = _currentToolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";
            var filteredModels = _availableModels.Where(m =>
                isLlamaCpp ? m.Format == ModelFormat.GGUF : m.Format != ModelFormat.GGUF
            ).ToList();

            await Dispatcher.InvokeAsync(() =>
            {
                ModelSelectorCombo.Items.Clear();
                if (filteredModels.Count > 0)
                {
                    foreach (var model in filteredModels)
                    {
                        var item = new ComboBoxItem
                        {
                            Content = model.DisplayName,
                            Tag = model
                        };
                        ModelSelectorCombo.Items.Add(item);

                        if (!string.IsNullOrWhiteSpace(preferredModelPath) &&
                            (string.Equals(model.FilePath, preferredModelPath, StringComparison.OrdinalIgnoreCase) ||
                             string.Equals(model.ConvertedFilePath, preferredModelPath, StringComparison.OrdinalIgnoreCase)))
                        {
                            ModelSelectorCombo.SelectedItem = item;
                        }
                    }

                    if (ModelSelectorCombo.SelectedIndex < 0)
                    {
                        ModelSelectorCombo.SelectedIndex = 0;
                    }
                }
                else
                {
                    ModelSelectorCombo.Items.Add(new ComboBoxItem { Content = "⚠️ 未找到模型文件" });
                    ModelSelectorCombo.SelectedIndex = 0;
                }
            });
        }

        private async Task<ModelInfo?> ConvertModelAsync(ModelInfo selectedModel)
        {
            if (_modelConverter == null)
            {
                LogMessage("❌ 模型转换服务未初始化");
                return null;
            }

            if (!selectedModel.NeedsConversion && selectedModel.Format != ModelFormat.PyTorch)
            {
                LogMessage($"ℹ️ 模型 {selectedModel.ModelName} 不是 .pth 格式，无需转换");
                return selectedModel;
            }

            LocalModelInfoText.Text = "🔄 正在转换模型...";
            var convertedPath = await _modelConverter.ConvertAsync(selectedModel);

            if (string.IsNullOrEmpty(convertedPath))
            {
                LocalModelInfoText.Text = "❌ 模型转换失败";
                MessageBox.Show($"模型转换失败：{selectedModel.ErrorMessage ?? "请检查 Python 环境与转换脚本"}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            LogMessage($"✅ 模型转换成功: {convertedPath}");
            await RefreshModelSelectorAsync(convertedPath);

            return _availableModels.FirstOrDefault(m =>
                string.Equals(m.FilePath, convertedPath, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(m.ConvertedFilePath, convertedPath, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(m.ModelName, selectedModel.ModelName, StringComparison.OrdinalIgnoreCase));
        }

        private async void ImportLocalModel(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_modelManagementService == null)
                {
                    LogMessage("❌ 模型管理服务未初始化");
                    return;
                }

                var openFileDialog = new OpenFileDialog
                {
                    Title = "选择要导入的 RWKV 模型",
                    Filter = "RWKV模型|*.pth;*.safetensors;*.gguf|PyTorch模型|*.pth|SafeTensors模型|*.safetensors|GGUF模型|*.gguf|所有文件|*.*",
                    Multiselect = false,
                    CheckFileExists = true
                };

                if (openFileDialog.ShowDialog() != true) return;

                var sourcePath = openFileDialog.FileName;
                var fileName = IOPath.GetFileName(sourcePath);
                var targetPath = IOPath.Combine(_modelManagementService.ModelsDirectory, fileName);

                if (IOFile.Exists(targetPath))
                {
                    var overwriteResult = MessageBox.Show(
                        $"模型目录中已存在同名文件：{fileName}\n\n是否覆盖？",
                        "导入模型",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (overwriteResult != MessageBoxResult.Yes) return;
                }

                LocalModelInfoText.Text = "📂 正在导入模型...";
                var importedPath = await _modelManagementService.ImportModelAsync(sourcePath, overwrite: true);
                LogMessage($"✅ 模型已导入: {importedPath}");

                await RefreshModelSelectorAsync(importedPath);

                var importedModel = _availableModels.FirstOrDefault(m =>
                    string.Equals(m.FilePath, importedPath, StringComparison.OrdinalIgnoreCase));

                if (string.Equals(IOPath.GetExtension(importedPath), ".pth", StringComparison.OrdinalIgnoreCase) && importedModel != null)
                {
                    var convertNow = MessageBox.Show(
                        $"已导入 .pth 模型：{importedModel.ModelName}\n\n是否立即转换为 .safetensors？",
                        "转换模型",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (convertNow == MessageBoxResult.Yes)
                    {
                        var convertedModel = await ConvertModelAsync(importedModel);
                        if (convertedModel != null)
                        {
                            LocalModelInfoText.Text = $"✅ 已导入并转换: {convertedModel.ModelName}";
                        }
                        return;
                    }

                    LocalModelInfoText.Text = $"🔄 已导入 .pth 模型，启动前可点击“转换选中.pth”";
                }
                else
                {
                    LocalModelInfoText.Text = $"✅ 已导入模型: {fileName}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导入本地模型失败");
                LogMessage($"❌ 导入模型失败: {ex.Message}");
                MessageBox.Show($"导入模型失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ConvertSelectedPthModel(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedModel = GetSelectedModel();
                if (selectedModel == null)
                {
                    MessageBox.Show("请先选择一个模型", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (selectedModel.Format != ModelFormat.PyTorch && !selectedModel.NeedsConversion)
                {
                    MessageBox.Show("当前选中的模型不是 .pth 文件，无需转换", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var result = MessageBox.Show(
                    $"是否将模型 '{selectedModel.ModelName}' 转换为 safetensors 格式？\n\n该操作需要本机可用的 Python 环境。",
                    "转换 .pth 模型",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes) return;

                var convertedModel = await ConvertModelAsync(selectedModel);
                if (convertedModel != null)
                {
                    LocalModelInfoText.Text = $"✅ 模型转换完成: {convertedModel.ModelName}";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "转换选中模型失败");
                LogMessage($"❌ 转换选中模型失败: {ex.Message}");
                MessageBox.Show($"转换模型失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 从 ModelScope 下载模型
        /// </summary>
        private async void DownloadModelFromSite(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_modelDownloadService == null)
                {
                    LogMessage("❌ 下载服务未初始化");
                    return;
                }

                // 从输入框获取ModelScope URL并更新下载服务配置
                var modelScopeUrl = ModelScopeUrlText.Text.Trim();
                if (string.IsNullOrEmpty(modelScopeUrl))
                {
                    MessageBox.Show("请输入ModelScope下载地址", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!_modelDownloadService.SetModelScopeConfigFromUrl(modelScopeUrl))
                {
                    MessageBox.Show(
                        "无法解析ModelScope地址，请检查格式。\n\n正确格式示例:\nhttps://www.modelscope.cn/models/AlicLi/rwkv7-g1-libtorch-st/files",
                        "地址格式错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 使用解析后的URL作为显示地址
                var displayUrl = _modelDownloadService.GetModelScopeFilesUrl();

                // 根据当前推理工具类型确定模型格式提示
                var selectedToolItem = InferenceToolCombo.SelectedItem as ComboBoxItem;
                var selectedToolContent = selectedToolItem?.Content as string ?? "";
                var currentTool = ResolveToolName(selectedToolContent);
                var isLlamaCpp = currentTool is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";
                var formatHint = isLlamaCpp
                    ? "模型为 GGUF 格式，可使用 llama-server 推理。"
                    : "模型为 safetensors 格式，可直接使用。";

                // 先获取文件列表
                LogMessage("📥 正在获取 ModelScope 文件列表...");
                LocalModelInfoText.Text = "正在获取文件列表...";
                
                var allFiles = await _modelDownloadService.GetModelFilesAsync();
                
                if (allFiles.Count == 0)
                {
                    MessageBox.Show("无法获取文件列表，请检查ModelScope仓库地址是否正确。", "获取失败", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    LocalModelInfoText.Text = "获取文件列表失败";
                    return;
                }

                // 弹出模型选择对话框
                var selectWindow = new Windows.ModelSelectWindow(allFiles, displayUrl, formatHint);
                selectWindow.Owner = this;
                var dialogResult = selectWindow.ShowDialog();
                
                if (dialogResult != true || selectWindow.SelectedFiles.Count == 0)
                {
                    LogMessage("⏹️ 已取消模型下载");
                    LocalModelInfoText.Text = "下载已取消";
                    return;
                }
                
                var selectedFiles = selectWindow.SelectedFiles;
                var totalSize = selectedFiles.Sum(f => f.SizeBytes);
                var sizeStr = totalSize >= 1073741824 
                    ? $"{totalSize / 1073741824.0:F2} GB" 
                    : $"{totalSize / 1048576.0:F1} MB";

                // 确认下载
                var result = MessageBox.Show(
                    $"即将下载 {selectedFiles.Count} 个模型文件（共 {sizeStr}）到 rwkv_models 目录。\n\n" +
                    $"下载源: {displayUrl}\n" +
                    $"{formatHint}\n\n" +
                    "是否继续？",
                    "模型下载",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);

                if (result != MessageBoxResult.Yes) return;

                // 获取模型保存目录（使用项目根目录）
                var baseDir = PathHelper.GetSafeBaseDirectory();
                var projectRoot = PathHelper.FindProjectRoot(baseDir);
                var modelsDir = IOPath.Combine(projectRoot, "rwkv_models");

                // 准备下载UI
                await Dispatcher.InvokeAsync(() =>
                {
                    DownloadModelBtn.IsEnabled = false;
                    CancelDownloadBtn.IsEnabled = true;
                    ModelDownloadProgress.Visibility = Visibility.Visible;
                    ModelDownloadProgress.Value = 0;
                    DownloadProgressText.Visibility = Visibility.Visible;
                    DownloadProgressText.Text = "正在连接 ModelScope...";
                });

                // 创建下载进度跟踪
                var progress = new Progress<Services.RwkvLocal.DownloadProgressEventArgs>(args =>
                {
                    Dispatcher.BeginInvoke(() =>
                    {
                        ModelDownloadProgress.Value = args.OverallProgress > 0 ? args.OverallProgress : args.Progress;
                        DownloadProgressText.Text = args.Status;
                    });
                });

                // 开始下载
                _downloadCts = new CancellationTokenSource();

                LogMessage("📥 开始从 ModelScope 下载模型...");
                LocalModelInfoText.Text = "正在下载模型，请稍候...";

                var downloadedFiles = await _modelDownloadService.DownloadModelFilesAsync(
                    selectedFiles,
                    modelsDir,
                    progress,
                    _downloadCts.Token);

                LogMessage($"✅ 模型下载完成！共 {downloadedFiles.Count} 个文件");
                LocalModelInfoText.Text = $"下载完成！共 {downloadedFiles.Count} 个文件已保存到 rwkv_models";

                // 刷新模型列表
                await InitializeGpuAndModelsAsync();
            }
            catch (OperationCanceledException)
            {
                LogMessage("⏹️ 模型下载已取消");
                LocalModelInfoText.Text = "下载已取消";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "下载模型失败");
                LogMessage($"❌ 下载模型失败: {ex.Message}");
                MessageBox.Show($"下载模型失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                await Dispatcher.InvokeAsync(() =>
                {
                    DownloadModelBtn.IsEnabled = true;
                    CancelDownloadBtn.IsEnabled = false;
                    ModelDownloadProgress.Visibility = Visibility.Collapsed;
                    DownloadProgressText.Visibility = Visibility.Collapsed;
                    _downloadCts?.Dispose();
                    _downloadCts = null;
                });
            }
        }

        /// <summary>
        /// 取消模型下载
        /// </summary>
        private void CancelModelDownload(object sender, RoutedEventArgs e)
        {
            try
            {
                _downloadCts?.Cancel();
                LogMessage("⏹️ 已请求取消下载...");
                CancelDownloadBtn.IsEnabled = false;
                CancelDownloadBtn.Content = "取消中...";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "取消下载失败");
            }
        }

        /// <summary>
        /// 下载状态变更事件处理
        /// </summary>
        private void OnDownloadStatusChanged(object? sender, string status)
        {
            Dispatcher.BeginInvoke(() =>
            {
                LocalModelInfoText.Text = status;
            });
        }

        /// <summary>
        /// 下载进度变更事件处理
        /// </summary>
        private void OnDownloadProgressChanged(object? sender, Services.RwkvLocal.DownloadProgressEventArgs args)
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (args.TotalFiles > 0)
                {
                    DownloadProgressText.Text = $"[{args.OverallFileIndex}/{args.TotalFiles}] {args.Status}";
                }
                else
                {
                    DownloadProgressText.Text = args.Status;
                }
            });
        }

        /// <summary>
        /// 下载完成事件处理
        /// </summary>
        private void OnDownloadCompleted(object? sender, DownloadCompletedEventArgs args)
        {
            Dispatcher.BeginInvoke(() =>
            {
                if (args.Success)
                {
                    LogMessage($"✅ 文件下载完成: {args.FileName}");
                }
                else
                {
                    LogMessage($"❌ 文件下载失败: {args.FileName} - {args.ErrorMessage}");
                }
            });
        }

        /// <summary>
        /// 从术语库动态更新语言映射和UI
        /// </summary>
        private void UpdateLanguageMappingFromTerminology()
        {
            try
            {
                // 获取术语库中的所有语种
                var terminologyLanguages = _termExtractor.GetSupportedLanguages();

                // 为术语库中的新语种添加默认语言代码映射
                foreach (var language in terminologyLanguages)
                {
                    if (!_languageMap.ContainsKey(language))
                    {
                        // 为新语种生成语言代码
                        var languageCode = GenerateLanguageCode(language);
                        _languageMap[language] = languageCode;
                        LogMessage($"🌐 为新语种 '{language}' 添加语言代码映射: {languageCode}");
                    }
                }

                // 更新UI中的语言选择器
                UpdateLanguageComboBox();

                LogMessage($"📚 已更新语言映射，当前支持 {_languageMap.Count} 种语言");
            }
            catch (Exception ex)
            {
                LogMessage($"⚠️ 更新语言映射失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 为新语种生成语言代码
        /// </summary>
        private string GenerateLanguageCode(string languageName)
        {
            // 常见语种的映射
            var commonMappings = new Dictionary<string, string>
            {
                {"葡萄牙语", "pt"}, {"荷兰语", "nl"}, {"阿拉伯语", "ar"},
                {"泰语", "th"}, {"印尼语", "id"}, {"马来语", "ms"}, {"土耳其语", "tr"},
                {"波兰语", "pl"}, {"捷克语", "cs"}, {"匈牙利语", "hu"}, {"希腊语", "el"},
                {"瑞典语", "sv"}, {"挪威语", "no"}, {"丹麦语", "da"}, {"芬兰语", "fi"},
                {"罗马尼亚语", "ro"}, {"保加利亚语", "bg"}, {"克罗地亚语", "hr"},
                {"斯洛伐克语", "sk"}, {"斯洛文尼亚语", "sl"}, {"立陶宛语", "lt"},
                {"拉脱维亚语", "lv"}, {"爱沙尼亚语", "et"}, {"乌克兰语", "uk"},
                {"白俄罗斯语", "be"}, {"塞尔维亚语", "sr"}, {"马其顿语", "mk"},
                {"阿尔巴尼亚语", "sq"}, {"格鲁吉亚语", "ka"}, {"亚美尼亚语", "hy"},
                {"希伯来语", "he"}, {"波斯语", "fa"}, {"乌尔都语", "ur"},
                {"印地语", "hi"}, {"孟加拉语", "bn"}, {"泰米尔语", "ta"},
                {"缅甸语", "my"}, {"高棉语", "km"}, {"老挝语", "lo"}
            };

            if (commonMappings.ContainsKey(languageName))
            {
                return commonMappings[languageName];
            }

            // 如果没有预定义映射，使用语种名称的前两个字符作为代码
            if (languageName.Length >= 2)
            {
                return languageName.Substring(0, 2).ToLower();
            }

            // 默认返回通用代码
            return "xx";
        }

        /// <summary>
        /// 更新语言选择下拉框
        /// </summary>
        private void UpdateLanguageComboBox()
        {
            if (LanguageCombo == null) return;

            var currentSelection = ((ComboBoxItem)LanguageCombo.SelectedItem)?.Content?.ToString();

            // 清空现有选项
            LanguageCombo.Items.Clear();

            // 添加所有支持的语种
            foreach (var language in _languageMap.Keys.OrderBy(x => x))
            {
                var item = new ComboBoxItem { Content = language };
                LanguageCombo.Items.Add(item);
            }

            // 恢复之前的选择，或选择默认语种
            if (!string.IsNullOrEmpty(currentSelection) && _languageMap.ContainsKey(currentSelection))
            {
                var itemToSelect = LanguageCombo.Items.Cast<ComboBoxItem>()
                    .FirstOrDefault(x => x.Content.ToString() == currentSelection);
                if (itemToSelect != null)
                {
                    LanguageCombo.SelectedItem = itemToSelect;
                }
            }
            else
            {
                // 选择默认语种（英语）
                var defaultItem = LanguageCombo.Items.Cast<ComboBoxItem>()
                    .FirstOrDefault(x => x.Content.ToString() == "英语");
                if (defaultItem != null)
                {
                    LanguageCombo.SelectedItem = defaultItem;
                }
                else if (LanguageCombo.Items.Count > 0)
                {
                    LanguageCombo.SelectedIndex = 0;
                }
            }
        }

        /// <summary>
        /// 配置依赖注入服务
        /// </summary>
        private void ConfigureServices(IServiceCollection services)
        {
            services.AddLogging(builder =>
            {
                builder.ClearProviders();
                builder.AddProvider(new CompositeLoggerProvider(LogMessageWrapper));
                builder.SetMinimumLevel(LogLevel.Information);
            });

            services.AddSingleton<ModelConfigurationManager>();
            services.AddSingleton<ConfigurationManager>();
            services.AddSingleton<TranslationService>();
            services.AddSingleton<DocumentProcessor>();
            services.AddSingleton<ExcelProcessor>();
            services.AddSingleton<DocumentProcessorFactory>();
            services.AddSingleton<TermExtractor>();
            services.AddSingleton<OcrInitializationService>();
            services.AddSingleton<AcrobatDetectionService>();
        }

        private void InitializeUI()
        {
            // 设置默认选中的引擎
            UpdateEngineSelection("rwkv");

            // 显示当前API地址
            UpdateApiUrlDisplay();

            // 绑定翻译方向切换事件
            ChineseToForeign.Checked += OnTranslationDirectionChanged;
            ForeignToChinese.Checked += OnTranslationDirectionChanged;

            // 绑定语言选择变化事件
            LanguageCombo.SelectionChanged += OnLanguageSelectionChanged;

            // 绑定输出格式切换事件
            BilingualOutput.Checked += OnOutputFormatChanged;
            TranslationOnlyOutput.Checked += OnOutputFormatChanged;

            // 设置窗口图标和标题
            this.Title = "RWKV文档术语翻译助手 v3.1";

            // 初始化翻译方向显示
            UpdateTranslationDirection();

            LogMessage("📊 界面初始化完成");

            // 初始化并发建议与滑块
            UpdateSuggestedParallelism();
            
            // 绑定滑块值变化事件
            ParallelismSlider.ValueChanged += (s, e) =>
            {
                int value = (int)Math.Round(ParallelismSlider.Value);
                ParallelismValueText.Text = $"并发: {value}";
                // 同步文本框值
                ParallelismInput.Text = value.ToString();
            };
            
            // 绑定文本框值变化事件
            ParallelismInput.TextChanged += (s, e) =>
            {
                if (int.TryParse(ParallelismInput.Text, out int value))
                {
                    // 确保值在有效范围内
                    value = Math.Max((int)ParallelismSlider.Minimum, Math.Min((int)ParallelismSlider.Maximum, value));
                    // 同步滑块值
                    ParallelismSlider.Value = value;
                    ParallelismValueText.Text = $"并发: {value}";
                }
            };

        }
        private void UpdateSuggestedParallelism()
        {
            try
            {
                var suggested = _translationService.GetSuggestedMaxParallelism();
                SuggestedParallelText.Text = suggested.ToString();
                ParallelismSlider.Value = suggested;
                ParallelismValueText.Text = $"并发: {suggested}";
                ParallelismInput.Text = suggested.ToString();
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "计算建议并发失败");
            }
        }

        private void DetectResources(object sender, RoutedEventArgs e)
        {
            UpdateSuggestedParallelism();
            LogMessage("🧮 已根据本机资源更新建议并发");
        }

        private void ApplyParallelism(object sender, RoutedEventArgs e)
        {
            try
            {
                int val = Math.Max(1, (int)Math.Round(ParallelismSlider.Value));
                _translationService.SetMaxParallelismOverride(val);
                
                // 同步每批次并发上限设置
                if (int.TryParse(BatchConcurrencyLimitText.Text, out var batchLimit) && batchLimit >= 0)
                {
                    ConcurrencyCalculator.BatchConcurrencyLimit = batchLimit;
                    _translationService.BatchConcurrencyLimit = batchLimit;
                }
                
                LogMessage($"💾 已应用并发设置: 最大并发={val}, 每批次上限={ConcurrencyCalculator.BatchConcurrencyLimit}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "应用并发设置失败");
                MessageBox.Show($"❌ 应用并发设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void UpdateLogUI()
        {
            if (!_logsChanged) return;

            List<string> messagesCopy;
            lock (_logLock)
            {
                messagesCopy = _logMessages.ToList();
                _logsChanged = false;
            }

            // 使用StringBuilder构建完整文本，一次性更新UI，避免多次渲染
            var sb = new StringBuilder();
            foreach (var msg in messagesCopy)
            {
                sb.AppendLine(msg);
            }
            
            LogTextBox.Text = sb.ToString();
            LogTextBox.ScrollToEnd();
        }

        private void LogMessage(string message, bool isSkipped = false, bool isFailure = false, string originalText = null, string modelResponse = null)
        {
            if (isSkipped)
            {
                return;
            }

            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var logLine = $"[{timestamp}] {message}";

            lock (_logLock)
            {
                if (_logMessages.Count >= 40)
                {
                    var oldMessage = _logMessages.Dequeue();
                    WriteToHistoryRecord(oldMessage);
                }
                _logMessages.Enqueue(logLine);
                _logsChanged = true;
            }

            // 已移除Dispatcher.Invoke，改由UpdateLogUI定时更新，解决UI卡顿问题

            if (isFailure && !string.IsNullOrEmpty(originalText) && !string.IsNullOrEmpty(modelResponse))
            {
                var failureLog = $"[{timestamp}] 翻译失败 - 原文: {originalText}, 模型返回: {modelResponse}";
                WriteToHistoryRecord(failureLog);
            }
        }

        private void LogMessageWrapper(string message)
        {
            LogMessage(message);
        }

        private void WriteToHistoryRecord(string message)
        {
            try
            {
                var dateStr = DateTime.Now.ToString("yyyyMMdd");
                var logFilePath = IOPath.Combine(_historyRecordPath, $"log_{dateStr}.txt");
                IOFile.AppendAllText(logFilePath, message + Environment.NewLine);
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "写入历史记录失败");
            }
        }

        private void UpdateEngineSelection(string engine)
        {
            _currentEngine = engine;

            LogMessage($"🔄 切换到 {engine} 引擎");
        }

        // 文件选择
        private void SelectFile(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择要翻译的文档",
                Filter = "Word文档 (*.docx)|*.docx|PDF文档 (*.pdf)|*.pdf|Excel文档 (*.xlsx;*.xls)|*.xlsx;*.xls|PowerPoint文档 (*.pptx)|*.pptx|所有支持格式|*.docx;*.pdf;*.xlsx;*.xls;*.pptx|所有文件 (*.*)|*.*"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                FilePathText.Text = openFileDialog.FileName;
                var fileInfo = new IOFileInfo(openFileDialog.FileName);

                // 检查文件格式是否支持
                var extension = IOPath.GetExtension(openFileDialog.FileName).ToLower();
                var supportedFormats = new[] { ".docx", ".pdf", ".xlsx", ".xls", ".pptx" };

                if (supportedFormats.Any(f => f == extension))
                {
                    StatusText.Text = $"📄 已选择文档 ({fileInfo.Length / 1024.0:F1} KB)";
                    LogMessage($"📁 选择文件: {openFileDialog.FileName}");

                    // 根据文件类型显示特定提示
                    switch (extension)
                    {
                        case ".docx":
                            LogMessage("📝 Word文档 - 支持完整格式保留和双语对照");
                            break;
                        case ".pdf":
                            LogMessage("📑 PDF文档 - 将提取文本内容进行翻译");
                            break;
                        case ".xlsx":
                        case ".xls":
                            LogMessage("📊 Excel文档 - 支持表格内容翻译");
                            break;
                        case ".pptx":
                            LogMessage("📽️ PowerPoint文档 - 支持幻灯片内容翻译");
                            break;
                    }
                }
                else
                {
                    StatusText.Text = "⚠️ 不支持的文件格式";
                    LogMessage($"❌ 不支持的文件格式: {extension}");
                    MessageBox.Show($"不支持的文件格式: {extension}\n\n支持的格式包括:\n• Word文档 (.docx)\n• PDF文档 (.pdf)\n• Excel文档 (.xlsx, .xls)\n• PowerPoint文档 (.pptx)",
                                  "文件格式错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    FilePathText.Text = "";
                }
            }
        }

        private void ClearFile(object sender, RoutedEventArgs e)
        {
            FilePathText.Text = "";
            StatusText.Text = "🟢 系统就绪";
            LogMessage("🗑️ 已清除文件选择");
        }

        /// <summary>
        /// 更新 rwkv_api.json 中的 api_url 配置
        /// </summary>
        private void UpdateRwkvApiUrl(string apiUrl)
        {
            try
            {
                var configPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "API_config", "rwkv_api.json");
                if (System.IO.File.Exists(configPath))
                {
                    var json = System.IO.File.ReadAllText(configPath);
                    var cfg = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                    if (cfg != null)
                    {
                        cfg["api_url"] = apiUrl;
                        var updatedJson = Newtonsoft.Json.JsonConvert.SerializeObject(cfg, Newtonsoft.Json.Formatting.Indented);
                        System.IO.File.WriteAllText(configPath, updatedJson);
                        _logger.LogInformation("已更新 rwkv_api.json 中的 api_url 为: {ApiUrl}", apiUrl);
                    }
                }

                // 同时更新 config.json 中的 rwkv_translator.api_url
                if (_configurationManager != null)
                {
                    var rwkvConfig = _configurationManager.GetConfig<Dictionary<string, object>>("rwkv_translator", new Dictionary<string, object>());
                    if (rwkvConfig != null)
                    {
                        rwkvConfig["api_url"] = apiUrl;
                        _configurationManager.SaveConfig("rwkv_translator", rwkvConfig);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "更新 RWKV API URL 配置失败");
            }
        }

        /// <summary>
        /// 更新 config.json 中 llama_cpp_translator 的 api_url 配置
        /// </summary>
        private void UpdateLlamaCppApiUrl(string apiUrl)
        {
            try
            {
                if (_configurationManager != null)
                {
                    var llamaConfig = _configurationManager.GetConfig<Dictionary<string, object>>("llama_cpp_translator", new Dictionary<string, object>());
                    if (llamaConfig != null)
                    {
                        llamaConfig["api_url"] = apiUrl;
                        _configurationManager.SaveConfig("llama_cpp_translator", llamaConfig);
                        _logger.LogInformation("已更新 llama_cpp_translator.api_url 为: {ApiUrl}", apiUrl);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "更新 LlamaCpp API URL 配置失败");
            }
        }

        /// <summary>
        /// 从配置中读取 llama_cpp 的对话模板名称
        /// </summary>
        private string GetLlamaCppChatTemplate()
        {
            try
            {
                if (_configurationManager != null)
                {
                    var llamaConfig = _configurationManager.GetConfig<Dictionary<string, object>>("llama_cpp_translator", null);
                    if (llamaConfig != null)
                    {
                        return llamaConfig.GetValueOrDefault("chat_template", "rwkv-world").ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "读取 LlamaCpp chat_template 配置失败");
            }
            return "rwkv-world";
        }

        private void UpdateApiUrlDisplay()
        {
            try
            {
                if (_configurationManager == null)
                {
                    ApiUrlDisplay.Text = "加载中...";
                    return;
                }
                var config = _configurationManager.GetTranslatorConfig("rwkv");
                ApiUrlDisplay.Text = config.ApiUrl;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "获取API地址失败");
                ApiUrlDisplay.Text = "获取失败";
            }
        }

        // 连接测试
        private async void TestConnection(object sender, RoutedEventArgs e)
        {
            TestStatusText.Text = "🔄 测试中...";
            TestStatusText.Foreground = System.Windows.Media.Brushes.Blue;

            try
            {
                // 切换到指定的翻译器
                _translationService.CurrentTranslatorType = _currentEngine;

                // 使用C#翻译服务进行连接测试
                var result = await _translationService.CurrentTranslator?.TestConnectionAsync();

                if (result == true)
                {
                    TestStatusText.Text = "✅ 测试成功";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Green;
                    TranslateButton.IsEnabled = true;
                    LogMessage($"✅ {_currentEngine} 连接测试成功");
                }
                else
                {
                    TestStatusText.Text = "❌ 测试失败";
                    TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
                    TranslateButton.IsEnabled = false;
                    LogMessage($"❌ {_currentEngine} 连接测试失败");
                }
            }
            catch (Exception ex)
            {
                TestStatusText.Text = "❌ 测试异常";
                TestStatusText.Foreground = System.Windows.Media.Brushes.Red;
                TranslateButton.IsEnabled = false;
                LogMessage($"❌ {_currentEngine} 连接测试异常: {ex.Message}");
            }
        }

        // 开始翻译
        private async void StartTranslation(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(FilePathText.Text))
            {
                MessageBox.Show("请先选择要翻译的文件！", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_isTranslating)
            {
                return;
            }

            _isTranslating = true;
            TranslateButton.IsEnabled = false;
            TranslationProgress.Visibility = Visibility.Visible;
            ProgressText.Visibility = Visibility.Visible;

            try
            {
                // 验证文件格式
                if (!ValidateFileFormat(FilePathText.Text))
                {
                    return;
                }

                LogMessage("🚀 开始翻译任务");
                StatusText.Text = "🔄 正在翻译中...";

                // 准备翻译请求
                var selectedLanguage = ((ComboBoxItem)LanguageCombo.SelectedItem)?.Content?.ToString() ?? "英语";
                var selectedModel = "default";

                // 确定输出格式
                var outputFormat = TranslationOnlyOutput.IsChecked == true ? "translation_only" : "bilingual";

                // 改进的语言代码处理逻辑
                string sourceLanguage, targetLanguageCode, targetLanguageName;

                if (ChineseToForeign.IsChecked == true)
                {
                    // 中文 → 外语
                    sourceLanguage = "zh";
                    targetLanguageCode = _languageMap.GetValueOrDefault(selectedLanguage, "en");
                    targetLanguageName = selectedLanguage;
                }
                else
                {
                    // 外语 → 中文
                    sourceLanguage = _languageMap.GetValueOrDefault(selectedLanguage, "en");
                    targetLanguageCode = "zh";
                    targetLanguageName = "中文";
                }

                LogMessage($"🔧 翻译配置详情:");
                LogMessage($"   📁 文件: {IOPath.GetFileName(FilePathText.Text)}");
                LogMessage($"   🤖 引擎: {_currentEngine}");
                LogMessage($"   🎯 模型: {selectedModel}");
                LogMessage($"   🌐 翻译方向: {sourceLanguage} → {targetLanguageCode}");
                LogMessage($"   📋 输出格式: {outputFormat}");
                LogMessage($"   📚 使用术语库: {(UseTerminology.IsChecked == true ? "是" : "否")}");
                LogMessage($"   ⚡ 术语预处理: {(PreprocessTerms.IsChecked == true ? "是" : "否")}");

                // 检查是否是PPT文件（批处理模式现在是默认的）
                var isPptFile = IOPath.GetExtension(FilePathText.Text).ToLower() == ".pptx";
                var usePptBatchProcess = isPptFile;

                if (usePptBatchProcess)
                {
                    LogMessage($"   📊 使用PPT批量处理模式: 是");
                }

                // 创建适当的文档处理器
                LogMessage("🔍 正在分析文件类型并创建处理器...");
                var documentProcessor = _documentProcessorFactory.CreateProcessor(FilePathText.Text);
                LogMessage($"✅ 文档处理器创建成功");

                // 设置翻译器和文档处理器选项
                LogMessage("⚙️ 配置翻译器和处理器选项...");
                _translationService.CurrentTranslatorType = _currentEngine;
                documentProcessor.SetTranslationOptions(
                    useTerminology: UseTerminology.IsChecked == true,
                    preprocessTerms: PreprocessTerms.IsChecked == true,
                    sourceLang: sourceLanguage,
                    targetLang: targetLanguageCode,
                    outputFormat: outputFormat
                );
                // 设置BOM表翻译选项（如果是Excel处理器）
                if (documentProcessor is ExcelProcessorAdapter excelAdapter)
                {
                    excelAdapter.Processor.UseBOMTranslation = UseBOMTranslation.IsChecked == true;
                    LogMessage($"📊 BOM表翻译功能: {(UseBOMTranslation.IsChecked == true ? "已启用" : "已禁用")}");
                }
                LogMessage("✅ 处理器选项配置完成");

                // 设置进度回调
                var startTime = DateTime.Now;
                var lastUpdateTime = DateTime.MinValue;
                const double minUpdateInterval = 100; // 最小更新间隔100ms

                documentProcessor.SetProgressCallback((progress, message) =>
                {
                    var now = DateTime.Now;
                    var elapsedSinceLastUpdate = (now - lastUpdateTime).TotalMilliseconds;

                    // 节流：只在进度变化较大或超过最小间隔时更新UI
                    if (elapsedSinceLastUpdate >= minUpdateInterval || progress >= 1.0)
                    {
                        lastUpdateTime = now;

                        // 使用BeginInvoke异步更新UI，不阻塞后台线程
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(() =>
                        {
                            TranslationProgress.Value = progress * 100;

                            // 计算预估剩余时间
                            var elapsed = DateTime.Now - startTime;
                            var estimatedTotal = progress > 0 ? TimeSpan.FromMilliseconds(elapsed.TotalMilliseconds / progress) : TimeSpan.Zero;
                            var remaining = estimatedTotal - elapsed;

                            var progressText = $"翻译进度: {progress * 100:F0}%";
                            if (progress > 0.05 && remaining.TotalSeconds > 0)
                            {
                                progressText += $" (预计剩余: {remaining:mm\\:ss})";
                            }

                            ProgressText.Text = progressText;
                            StatusText.Text = $"🔄 {message}";
                            LogMessage($"📊 进度: {progress * 100:F0}% - {message}");

                            // 更新状态栏
                            StatusBarText.Text = $"正在翻译... {progress * 100:F0}%";
                        }));
                    }
                });

                // 加载术语库
                var terminology = new Dictionary<string, string>();
                if (UseTerminology.IsChecked == true)
                {
                    // 根据翻译方向加载不同的术语库
                    var terminologyLanguageName = selectedLanguage;
                    
                    if (ChineseToForeign.IsChecked == true)
                    {
                        // 中文→外语：使用中译外术语库
                        terminology = _termExtractor.GetTermsForLanguage(terminologyLanguageName);
                        LogMessage($"📚 已加载中译外术语库（{terminologyLanguageName}）：{terminology.Count} 条");
                    }
                    else
                    {
                        // 外语→中文：直接使用外译中术语库，不进行镜像
                        terminology = _termExtractor.GetReverseTermsForLanguage(terminologyLanguageName);
                        LogMessage($"📚 已加载外译中术语库（{terminologyLanguageName}）：{terminology.Count} 条");
                    }

                    // 验证术语库加载结果
                    if (terminology.Count == 0)
                    {
                        LogMessage($"⚠️ 警告：术语库为空！请检查以下项目：");
                        LogMessage($"   - 术语库文件是否存在");
                        LogMessage($"   - 语言名称是否正确：'{terminologyLanguageName}'");
                        LogMessage($"   - 术语库中是否包含该语言的术语");
                    }
                    else
                    {
                        // 显示术语库样本
                        var sampleTerms = terminology.Take(3).Select(kv => $"'{kv.Key}' → '{kv.Value}'");
                        LogMessage($"   术语库样本: {string.Join(", ", sampleTerms)}");

                        if (PreprocessTerms.IsChecked == true)
                        {
                            LogMessage($"   🔧 术语预处理已启用");
                        }
                        else
                        {
                            LogMessage($"   ⚡ 术语预处理已关闭，将使用占位符策略");
                        }
                    }
                }

                // 调用C#翻译服务
                string outputPath;
                var filePath = FilePathText.Text; // Capture for thread safety

                // 在后台线程运行耗时任务，防止阻塞UI
                outputPath = await Task.Run(async () =>
                {
                    // 根据文件类型和批处理选项选择处理方式
                    if (usePptBatchProcess)
                    {
                        // 使用PPT批处理模式
                        LogMessage("📊 正在使用PPT批量处理模式...");

                        // 获取PPT处理器
                        var pptProcessorAdapter = documentProcessor as PPTProcessorAdapter;
                        if (pptProcessorAdapter == null)
                        {
                            LogMessage("❌ 错误：无法创建PPT处理器适配器");
                            throw new InvalidOperationException("无法创建PPT处理器，文件格式可能不支持");
                        }

                        // 从适配器中获取实际的处理器实例
                        var pptProcessor = pptProcessorAdapter.Processor;
                        LogMessage("✅ 成功获取PPT处理器实例");

                        // 调用批处理方法（保留副本现在是默认的）
                        LogMessage("🚀 开始执行PPT批处理翻译...");
                        var result = await pptProcessor.BatchProcessPPTDocumentAsync(
                            filePath,
                            targetLanguageName,
                            terminology,
                            createCopy: true
                        );
                        LogMessage("✅ PPT批处理翻译完成");
                        return result;
                    }
                    else
                    {
                        // 使用标准处理模式
                        LogMessage("🚀 使用标准处理模式进行翻译...");
                        var result = await documentProcessor.ProcessDocumentAsync(
                            filePath,
                            targetLanguageName,
                            terminology
                        );
                        LogMessage("✅ 标准模式翻译完成");
                        return result;
                    }
                });

                if (!string.IsNullOrEmpty(outputPath))
                {
                    var totalTime = DateTime.Now - startTime;
                    TranslationProgress.Value = 100;
                    ProgressText.Text = $"翻译进度: 100% (耗时: {totalTime:mm\\:ss})";
                    StatusText.Text = "✅ 翻译完成！";
                    StatusBarText.Text = "翻译完成";

                    LogMessage($"🎉 翻译任务完成！");
                    LogMessage($"   📁 输出文件: {outputPath}");
                    LogMessage($"   ⏱️ 总耗时: {totalTime:hh\\:mm\\:ss}");
                    LogMessage($"   📊 输出格式: {(TranslationOnlyOutput.IsChecked == true ? "仅译文" : "双语对照")}");

                    if (usePptBatchProcess)
                    {
                        LogMessage($"   📊 PPT批处理：已生成Excel数据文件和翻译后的PPT");
                    }

                    var message = $"翻译完成！\n\n📁 输出文件：{outputPath}\n⏱️ 耗时：{totalTime:hh\\:mm\\:ss}";

                    if (usePptBatchProcess)
                    {
                        message += $"\n📊 已生成Excel数据文件，包含原文和翻译结果";
                    }

                    MessageBox.Show(message, "翻译成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    throw new Exception("翻译失败：未生成输出文件");
                }
            }
            catch (Exception ex)
            {
                TranslationProgress.Value = 0;
                ProgressText.Text = "翻译进度: 0%";
                StatusText.Text = "❌ 翻译失败";

                // 详细的异常分析
                LogMessage($"❌ 翻译异常: {ex.GetType().Name} - {ex.Message}");
                if (ex.InnerException != null)
                {
                    LogMessage($"   内部异常: {ex.InnerException.Message}");
                }

                var suggestion = GetExceptionSuggestion(ex);
                var fullMessage = $"翻译失败：{ex.Message}";
                if (!string.IsNullOrEmpty(suggestion))
                {
                    fullMessage += $"\n\n💡 建议：{suggestion}";
                }

                MessageBox.Show(fullMessage, "翻译失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isTranslating = false;
                TranslateButton.IsEnabled = true;
                TranslationProgress.Visibility = Visibility.Collapsed;
                ProgressText.Visibility = Visibility.Collapsed;
            }
        }

        // 菜单事件
        private void OpenrwkvSettings(object sender, RoutedEventArgs e)
        {
            LogMessage("⚙️ 打开RWKV_lightning设置");
            OpenEngineConfig("rwkv");
        }

        private void OpenLlamaCppSettings(object sender, RoutedEventArgs e)
        {
            LogMessage("⚙️ 打开llama_cpp设置");
            OpenEngineConfig("llama_cpp");
        }

        private void OpenTranslationAssistant(object sender, RoutedEventArgs e)
        {
            LogMessage("🌐 打开翻译助手");
            try
            {
                var translationWindow = new TranslationAssistantWindow(_translationService, _loggerFactory);
                translationWindow.Owner = this;
                translationWindow.Show();
                LogMessage("✅ 翻译助手已打开");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开翻译助手失败: {ex.Message}");
                MessageBox.Show($"打开翻译助手失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenEngineConfig(string engineType)
        {
            try
            {
                var configWindow = new Windows.EngineConfigWindow
                {
                    Owner = this
                };
                configWindow.ShowDialog();
                LogMessage($"✅ {engineType} 配置窗口已关闭");

                UpdateApiUrlDisplay();
                LogMessage("🔄 已更新API地址显示");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开配置窗口失败: {ex.Message}");
                MessageBox.Show($"打开配置窗口失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ShowLicenseInfo(object sender, RoutedEventArgs e)
        {
            LogMessage("ℹ️ 显示授权信息");
            MessageBox.Show("授权信息：已授权\n作者留言：☭", "授权信息", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ShowAbout(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("RWKV文档术语翻译助手 @2024~2026\n\n支持Word、Excel文档\n\n Github - LuckySongXiao © 2024",
                          "关于", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>
        /// BSZ 增量输入框只允许数字
        /// </summary>
        private void BszIncrementText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // 只允许数字
            e.Handled = !Regex.IsMatch(e.Text, "^[0-9]+$");
        }

        /// <summary>
        /// 从推理工具ComboBox内容解析工具名称
        /// </summary>
        private static string ResolveToolName(string toolContent)
        {
            if (toolContent.Contains("CUDA") && toolContent.Contains("llama"))
                return "llama-cuda";
            if (toolContent.Contains("SYCL") && toolContent.Contains("llama"))
                return "llama-sycl";
            if (toolContent.Contains("Vulkan") && toolContent.Contains("llama"))
                return "llama-vulkan";
            if (toolContent.Contains("CPU") && toolContent.Contains("llama"))
                return "llama-cpu";
            if (toolContent.Contains("benchmark"))
                return "benchmark";
            return "rwkv_lightning";
        }

        /// <summary>
        /// 推理工具选择变化时，自动切换ModelScope下载地址
        /// <summary>
        /// 计算方式选择变化时，过滤推理工具列表并显示/隐藏GPU选择
        /// </summary>
        private void ComputeModeCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 防止递归调用或服务未初始化
            if (_suppressToolSelectionEvents || _gpuResourceService == null) return;
            
            try
            {
                var isGpuMode = IsGpuComputeMode();
                
                // 显示/隐藏GPU选择面板
                GpuSelectPanel.Visibility = isGpuMode ? Visibility.Visible : Visibility.Collapsed;
                
                // 根据计算方式刷新推理工具列表
                RefreshInferenceTools();
            }
            catch { }
        }

        /// <summary>
        /// 判断当前是否为GPU计算模式
        /// </summary>
        private bool IsGpuComputeMode()
        {
            var selectedMode = ComputeModeCombo.SelectedItem as ComboBoxItem;
            var modeTag = selectedMode?.Tag as string ?? "gpu";
            return modeTag == "gpu";
        }

        /// <summary>
        /// 根据当前计算方式刷新推理工具ComboBox
        /// GPU模式: rwkv_lightning, benchmark, llama-cuda, llama-sycl, llama-vulkan
        /// CPU模式: llama-cpu
        /// </summary>
        private void RefreshInferenceTools()
        {
            var isGpuMode = IsGpuComputeMode();
            var previousTool = _currentToolName;
            
            // 抑制事件，防止递归
            _suppressToolSelectionEvents = true;
            try
            {
                InferenceToolCombo.Items.Clear();
                
                if (isGpuMode)
                {
                    // GPU模式：显示所有GPU推理工具
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "rwkv_lightning.exe (NVIDIA GPU - 完整翻译API服务)", Tag = "rwkv_lightning" });
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "benchmark.exe (NVIDIA GPU - 性能测试)", Tag = "benchmark" });
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "llama-server.exe [CUDA] (GGUF模型 - NVIDIA GPU)", Tag = "llama-cuda" });
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "llama-server.exe [SYCL] (GGUF模型 - Intel GPU)", Tag = "llama-sycl" });
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "llama-server.exe [Vulkan] (GGUF模型 - AMD/通用GPU)", Tag = "llama-vulkan" });
                }
                else
                {
                    // CPU模式：仅显示CPU推理工具
                    InferenceToolCombo.Items.Add(new ComboBoxItem { Content = "llama-server.exe [CPU] (GGUF模型 - 纯CPU推理)", Tag = "llama-cpu" });
                }
                
                // 尝试恢复之前选择的工具，否则选第一项
                int targetIndex = 0;
                for (int i = 0; i < InferenceToolCombo.Items.Count; i++)
                {
                    var item = InferenceToolCombo.Items[i] as ComboBoxItem;
                    if ((item?.Tag as string) == previousTool)
                    {
                        targetIndex = i;
                        break;
                    }
                }
                InferenceToolCombo.SelectedIndex = targetIndex;
                
                // 手动更新当前工具名和描述
                var selectedItem = InferenceToolCombo.SelectedItem as ComboBoxItem;
                var newToolName = selectedItem?.Tag as string ?? "rwkv_lightning";
                var toolChanged = newToolName != previousTool;
                _currentToolName = newToolName;
                UpdateInferenceToolDescription();
                
                // 恢复标志后，如果工具发生了变化，需要刷新模型列表
                if (toolChanged)
                {
                    _suppressToolSelectionEvents = false;
                    _ = InitializeGpuAndModelsAsync();
                    return; // 已恢复标志，直接返回
                }
            }
            finally
            {
                _suppressToolSelectionEvents = false;
            }
        }

        /// <summary>
        /// 推理工具选择变化处理
        /// </summary>
        private void InferenceToolCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 防止递归调用或服务未初始化
            if (_suppressToolSelectionEvents || _gpuResourceService == null) return;
            
            try
            {
                var selectedToolItem = InferenceToolCombo.SelectedItem as ComboBoxItem;
                // 优先使用Tag获取工具名，回退到Content解析
                var toolName = selectedToolItem?.Tag as string ?? ResolveToolName(selectedToolItem?.Content as string ?? "");
                _currentToolName = toolName;
                
                UpdateInferenceToolDescription();

                // 切换推理工具后刷新本地模型列表（不同工具对应不同格式的模型）
                _ = InitializeGpuAndModelsAsync();
            }
            catch { }
        }

        /// <summary>
        /// 根据当前工具名更新描述文本和下载地址
        /// </summary>
        private void UpdateInferenceToolDescription()
        {
            var toolName = _currentToolName;
            var isLlamaCpp = toolName is "llama-cuda" or "llama-sycl" or "llama-vulkan" or "llama-cpu";

            if (isLlamaCpp)
            {
                ModelScopeUrlText.Text = DefaultGgufModelScopeUrl;
                InferenceToolDescription.Text = toolName switch
                {
                    "llama-cuda" => "💡 llama-server [CUDA] - NVIDIA GPU加速，需安装CUDA驱动",
                    "llama-sycl" => "💡 llama-server [SYCL] - Intel GPU加速（Arc独显/Xe核显）",
                    "llama-vulkan" => "💡 llama-server [Vulkan] - AMD/通用GPU加速",
                    "llama-cpu" => "💡 llama-server [CPU] - 纯CPU推理，无需GPU",
                    _ => "💡 llama-server.exe 使用GGUF格式模型，提供OpenAI兼容API"
                };
            }
            else
            {
                ModelScopeUrlText.Text = DefaultStModelScopeUrl;
                if (toolName == "benchmark")
                {
                    InferenceToolDescription.Text = "💡 benchmark.exe 仅用于性能测试，不提供翻译API服务";
                }
                else
                {
                    InferenceToolDescription.Text = "💡 rwkv_lightning.exe 提供完整的翻译API服务（NVIDIA GPU专属），使用safetensors格式模型";
                }
            }
        }

        private void OpenTranslationRulesEditor(object sender, RoutedEventArgs e)
        {
            LogMessage("🧩 打开翻译规则编辑器");
            try
            {
                var win = new Windows.TranslationRulesWindow();
                win.Owner = this;
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开翻译规则编辑器失败: {ex.Message}");
                MessageBox.Show($"打开翻译规则编辑器失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenTerminologyEditor(object sender, RoutedEventArgs e)
        {
            LogMessage("📝 打开术语库编辑器");
            try
            {
                var editorWindow = new Windows.TerminologyEditorWindow(_translationService)
                {
                    Owner = this
                };
                editorWindow.ShowDialog();
                LogMessage("✅ 术语库编辑器已关闭");

                // 术语库编辑器关闭后，重新加载术语库并更新语言映射
                _termExtractor.LoadTerminologyData();
                UpdateLanguageMappingFromTerminology();
                LogMessage("🔄 已刷新术语库和语言映射");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开术语库编辑器失败: {ex.Message}");
                MessageBox.Show($"打开术语库编辑器失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenOutputConfigWindow(object sender, RoutedEventArgs e)
        {
            LogMessage("⚙️ 打开输出格式配置");
            try
            {
                var configWindow = new OutputConfigWindow(_outputConfig)
                {
                    Owner = this
                };
                
                if (configWindow.ShowDialog() == true)
                {
                    LogMessage("✅ 输出格式配置已应用");
                    LogMessage($"📋 输出格式: {GetOutputFormatName(_outputConfig.Format)}");
                    LogMessage($"🔤 分段模式: {(_outputConfig.Segmentation == TranslationOutputConfig.SegmentationMode.Paragraph ? "按段落" : "按句子")}");
                    LogMessage($"🎨 术语高亮: {(_outputConfig.HighlightTerms ? "启用" : "禁用")}");
                    MessageBox.Show("输出格式配置已应用", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogMessage("ℹ️ 输出格式配置已取消");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开输出格式配置失败: {ex.Message}");
                MessageBox.Show($"打开输出格式配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

        private void SaveLog(object sender, RoutedEventArgs e)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Title = "保存日志文件",
                    Filter = "文本文件 (*.txt)|*.txt|日志文件 (*.log)|*.log|所有文件 (*.*)|*.*",
                    FileName = $"translation_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    IOFile.WriteAllText(saveFileDialog.FileName, LogTextBox.Text);
                    LogMessage($"💾 日志已保存到: {saveFileDialog.FileName}");
                    MessageBox.Show("✅ 日志保存成功！", "保存完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 保存日志失败: {ex.Message}");
                MessageBox.Show($"保存日志失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearLog(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("确定要清除所有日志吗？", "确认清除",
                                       MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                LogTextBox.Clear();
                LogMessage("🗑️ 日志已清除");
            }
        }

        private void OpenOutputDirectory(object sender, RoutedEventArgs e)
        {
            try
            {
                // 获取输出目录路径
                var outputDir = IOPath.Combine(PathHelper.GetSafeBaseDirectory(), "输出");

                // 如果输出目录不存在，则创建它
                if (!IODirectory.Exists(outputDir))
                {
                    IODirectory.CreateDirectory(outputDir);
                    LogMessage($"📁 创建输出目录: {outputDir}");
                }

                // 使用Windows资源管理器打开目录
                System.Diagnostics.Process.Start("explorer.exe", outputDir);
                LogMessage($"📁 已打开输出目录: {outputDir}");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ 打开输出目录失败: {ex.Message}");
                MessageBox.Show($"打开输出目录失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 翻译方向切换事件处理
        private void OnTranslationDirectionChanged(object sender, RoutedEventArgs e)
        {
            UpdateTranslationDirection();
        }

        // 语言选择变化事件处理
        private void OnLanguageSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateTranslationDirection();
        }

        // 更新翻译方向和语言代码映射
        private void UpdateTranslationDirection()
        {
            if (LanguageCombo.SelectedItem == null) return;

            var selectedLanguage = ((ComboBoxItem)LanguageCombo.SelectedItem).Content.ToString();
            var isChineseToForeign = ChineseToForeign.IsChecked == true;

            if (isChineseToForeign)
            {
                // 中文 → 外语
                StatusText.Text = $"🔄 翻译方向: 中文 → {selectedLanguage}";
                LogMessage($"🔄 设置翻译方向: 中文 → {selectedLanguage}");
            }
            else
            {
                // 外语 → 中文
                StatusText.Text = $"🔄 翻译方向: {selectedLanguage} → 中文";
                LogMessage($"🔄 设置翻译方向: {selectedLanguage} → 中文");
            }

            // 更新输出格式提示
            UpdateOutputFormatHint();
        }

        // 输出格式切换事件处理
        private void OnOutputFormatChanged(object sender, RoutedEventArgs e)
        {
            UpdateOutputFormatHint();
        }

        // 更新输出格式提示
        private void UpdateOutputFormatHint()
        {
            var outputFormat = TranslationOnlyOutput.IsChecked == true ? "仅译文" : "双语对照";
            var description = TranslationOnlyOutput.IsChecked == true
                ? "仅显示翻译结果，文档更简洁"
                : "原文和译文并排显示，便于对比检查";

            LogMessage($"📋 输出格式: {outputFormat} - {description}");
        }

        // 验证文件格式支持
        private bool ValidateFileFormat(string filePath)
        {
            var extension = IOPath.GetExtension(filePath).ToLower();
            var supportedFormats = new Dictionary<string, string>
            {
                { ".docx", "Word文档" },
                { ".pdf", "PDF文档" },
                { ".xlsx", "Excel工作簿" },
                { ".xls", "Excel工作簿(旧版)" },
                { ".pptx", "PowerPoint演示文稿" }
            };

            if (supportedFormats.ContainsKey(extension))
            {
                LogMessage($"✅ 文件格式验证通过: {supportedFormats[extension]} ({extension})");
                return true;
            }
            else
            {
                var supportedList = string.Join(", ", supportedFormats.Values);
                LogMessage($"❌ 不支持的文件格式: {extension}");
                MessageBox.Show($"不支持的文件格式: {extension}\n\n支持的格式包括:\n{supportedList}",
                              "文件格式错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }






        // 错误建议方法
        private static string GetErrorSuggestion(string errorMessage)
        {
            if (string.IsNullOrEmpty(errorMessage)) return "";

            var lowerError = errorMessage.ToLower();

            if (lowerError.Contains("api") && lowerError.Contains("key"))
                return "请检查API密钥是否正确配置，可通过菜单栏的设置选项进行配置";

            if (lowerError.Contains("network") || lowerError.Contains("connection") || lowerError.Contains("timeout"))
                return "请检查网络连接是否正常，或尝试切换到其他翻译引擎";

            if (lowerError.Contains("quota") || lowerError.Contains("limit"))
                return "API配额已用完，请检查账户余额或等待配额重置";

            if (lowerError.Contains("model") || lowerError.Contains("不支持"))
                return "当前模型可能不可用，请尝试切换到其他模型";

            if (lowerError.Contains("file") || lowerError.Contains("文件"))
                return "请检查文件是否存在且未被其他程序占用";

            if (lowerError.Contains("python"))
                return "Python环境可能有问题，请确保Python已正确安装并配置";

            return "请检查配置设置，或尝试重新启动程序";
        }

        private static string GetExceptionSuggestion(Exception ex)
        {
            var exceptionType = ex.GetType().Name;

            switch (exceptionType)
            {
                case "HttpRequestException":
                    return "网络请求失败，请检查网络连接或API服务状态";
                case "TaskCanceledException":
                    return "请求超时，请检查网络连接或尝试增加超时时间";
                case "FileNotFoundException":
                    return "文件未找到，请确认文件路径正确且文件存在";
                case "UnauthorizedAccessException":
                    return "文件访问被拒绝，请检查文件权限或关闭占用文件的程序";
                case "JsonException":
                    return "数据解析错误，可能是API返回格式异常";
                case "ArgumentException":
                    return "参数错误，请检查输入的配置参数是否正确";
                default:
                    return GetErrorSuggestion(ex.Message);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            // 停止所有本地推理服务
            _llamaCppProcessService?.StopAsync().Wait(TimeSpan.FromSeconds(5));
            _llamaCppProcessService?.Dispose();
            _rwkvProcessService?.StopAsync().Wait(TimeSpan.FromSeconds(5));
            _rwkvProcessService?.Dispose();
            _gpuResourceService?.Dispose();

            // 清理资源
            _translationService?.StopCurrentOperations();
            base.OnClosed(e);
        }

        /// <summary>
        /// 测试PPT翻译输出模式功能
        /// </summary>
        private async Task TestPPTTranslationModes()
        {
            try
            {
                _logger.LogInformation("=== 开始测试PPT翻译输出模式功能 ===");

                // 创建PPT处理器 - 创建专用的日志记录器
                var loggerFactory = LoggerFactory.Create(builder => { });
                var pptLogger = loggerFactory.CreateLogger<PPTProcessor>();
                var pptProcessor = new PPTProcessor(_translationService, pptLogger);

                // 测试文件路径（请确保存在）
                var testFile = "test_input.pptx";
                if (!IOFile.Exists(testFile))
                {
                    _logger.LogWarning($"测试文件不存在: {testFile}");
                    _logger.LogInformation("请将测试PPT文件命名为 'test_input.pptx' 并放在程序目录下");
                    return;
                }

                // 示例术语库
                var terminology = new Dictionary<string, string>
                {
                    { "人工智能", "Artificial Intelligence" },
                    { "机器学习", "Machine Learning" },
                    { "深度学习", "Deep Learning" },
                    { "神经网络", "Neural Network" }
                };

                // 测试1：双语输出模式（自动生成纯翻译版本）
                _logger.LogInformation("--- 测试双语输出模式 ---");
                pptProcessor.SetTranslationOptions(
                    useTerminology: true,
                    preprocessTerms: true,
                    outputFormat: "bilingual",  // 双语模式
                    preserveFormatting: true,
                    translateNotes: true,
                    translateCharts: true,
                    translateSmartArt: true,
                    autoAdjustLayout: true
                );

                var bilingualPath = await pptProcessor.ProcessDocumentAsync(
                    filePath: testFile,
                    targetLanguage: "英语",
                    terminology: terminology
                );

                _logger.LogInformation($"✅ 双语版本已生成: {bilingualPath}");

                // 等待一段时间
                await Task.Delay(2000);

                // 测试2：手动生成纯翻译版本
                _logger.LogInformation("--- 测试手动生成纯翻译版本 ---");
                if (IOFile.Exists(bilingualPath))
                {
                    var translationOnlyPath = pptProcessor.GenerateTranslationOnlyFromBilingual(bilingualPath);
                    _logger.LogInformation($"✅ 手动生成的纯翻译版本: {translationOnlyPath}");
                }

                _logger.LogInformation("🎉 PPT翻译模式测试完成！");
            }
            catch (Exception ex)
            {
                _logger.LogError($"❌ PPT翻译模式测试失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 菜单项：测试PPT翻译模式
        /// </summary>
        private async void TestPPTTranslationModes(object sender, RoutedEventArgs e)
        {
            try
            {
                // 禁用界面
                IsEnabled = false;
                StatusText.Text = "正在测试PPT翻译模式...";

                await TestPPTTranslationModes();

                StatusText.Text = "PPT翻译模式测试完成";
                MessageBox.Show("PPT翻译模式测试完成！请查看日志了解详细信息。", "测试完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                StatusText.Text = "PPT翻译模式测试失败";
                MessageBox.Show($"PPT翻译模式测试失败：{ex.Message}", "测试失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // 恢复界面
                IsEnabled = true;
            }
        }


    }

    /// <summary>
    /// 翻译请求类（简化版，用于兼容性）
    /// </summary>
    public class TranslationRequest
    {
        public string FilePath { get; set; }
        public string TargetLanguage { get; set; }
        public string TargetLanguageCode { get; set; }
        public string SourceLanguage { get; set; }
        public string Engine { get; set; }
        public string Model { get; set; }
        public bool UseTerminology { get; set; }
        public bool PreprocessTerms { get; set; }
        public string OutputFormat { get; set; }
    }

    /// <summary>
    /// 翻译结果类（简化版，用于兼容性）
    /// </summary>
    public class TranslationResult
    {
        public bool Success { get; set; }
        public string OutputPath { get; set; }
        public string ErrorMessage { get; set; }
        public int Progress { get; set; }
        public string StatusMessage { get; set; }
    }
}
