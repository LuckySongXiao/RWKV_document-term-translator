using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Diagnostics;
using System.Net.Http;
using Newtonsoft.Json;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.DependencyInjection;
using DocumentTranslator.Services.Translation;
using DocumentTranslator.Services.Logging;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Windows
{
    public partial class EngineConfigWindow : Window
    {
        private readonly HttpClient _httpClient;
        private readonly string _configPath;
        private readonly TranslationService _translationService;
        private readonly ConfigurationManager _configurationManager;
        private readonly ILogger<EngineConfigWindow> _logger;

        public EngineConfigWindow()
        {
            InitializeComponent();

            // 初始化依赖注入容器
            var services = new ServiceCollection();
            ConfigureServices(services);
            var serviceProvider = services.BuildServiceProvider();

            // 获取服务实例
            _translationService = serviceProvider.GetRequiredService<TranslationService>();
            _configurationManager = serviceProvider.GetRequiredService<ConfigurationManager>();
            _logger = serviceProvider.GetRequiredService<ILogger<EngineConfigWindow>>();

            _httpClient = new HttpClient();
            _configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config");

            LoadAllConfigs();
        }

        /// <summary>
        /// 配置依赖注入服务
        /// </summary>
        private void ConfigureServices(IServiceCollection services)
        {
            // 配置日志
            services.AddLogging(builder =>
            {
                builder.ClearProviders();
                builder.AddProvider(new TranslationLoggerProvider());
                builder.SetMinimumLevel(LogLevel.Information);
            });

            // 注册服务
            services.AddSingleton<ModelConfigurationManager>();
            services.AddSingleton<ConfigurationManager>();
            services.AddSingleton<TranslationService>();
        }

        private void LoadAllConfigs()
        {
            LoadrwkvConfig();
            LoadLlamaCppConfig();
        }

        #region RWKV配置
        private void LoadrwkvConfig()
        {
            try
            {
                var configFile = Path.Combine(_configPath, "rwkv_api.json");
                string apiUrlFromJson = null;
                double? temperature = null;
                string maxTokens = null;
                bool? skipSSL = null;
                bool? useBatchTranslate = null;

                if (File.Exists(configFile))
                {
                    var json = File.ReadAllText(configFile);
                    var config = JsonConvert.DeserializeObject<dynamic>(json);

                    apiUrlFromJson = config?.api_url?.ToString();
                    temperature = config?.temperature;
                    maxTokens = config?.max_tokens?.ToString();
                    skipSSL = config?.skip_ssl;
                    useBatchTranslate = config?.use_batch_translate;
                }

                // 从config.json读取配置作为备用
                var translatorConfig = _configurationManager.GetTranslatorConfig("rwkv");
                rwkvApiUrl.Text = apiUrlFromJson ?? translatorConfig.ApiUrl;
                rwkvTemperature.Value = temperature ?? translatorConfig.Temperature;
                rwkvMaxTokens.Text = maxTokens ?? translatorConfig.MaxTokens.ToString();
                rwkvSkipSSL.IsChecked = skipSSL ?? false;
                rwkvUseBatchTranslate.IsChecked = useBatchTranslate ?? translatorConfig.UseBatchTranslate;

                LoadRWKVApiUrls();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载RWKV配置失败");
                MessageBox.Show($"加载RWKV配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadRWKVApiUrls()
        {
            try
            {
                var selectedUrl = rwkvApiUrl.Text;
                rwkvApiUrl.Items.Clear();

                var defaultUrls = new List<string>
                {
                    "http://localhost:8000/translate/v1/batch-translate",
                    "http://localhost:8000/v1/chat/completions",
                    "http://localhost:8000/v2/chat/completions"
                };

                var configFile = Path.Combine(_configPath, "rwkv_api.json");
                if (File.Exists(configFile))
                {
                    var json = File.ReadAllText(configFile);
                    var config = JsonConvert.DeserializeObject<dynamic>(json);
                    var availableUrls = config?.available_api_urls as Newtonsoft.Json.Linq.JArray;
                    if (availableUrls != null && availableUrls.Count > 0)
                    {
                        foreach (var url in availableUrls)
                        {
                            var urlStr = url?.ToString();
                            if (!string.IsNullOrWhiteSpace(urlStr) && !defaultUrls.Contains(urlStr))
                            {
                                defaultUrls.Add(urlStr);
                            }
                        }
                    }
                }

                foreach (var url in defaultUrls)
                {
                    rwkvApiUrl.Items.Add(url);
                }

                rwkvApiUrl.Text = selectedUrl;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载RWKV API地址列表失败");
            }
        }

        private async void TestrwkvConnection(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            button.IsEnabled = false;
            button.Content = "🔄 测试中...";

            try
            {
                // 检查翻译服务是否已初始化
                if (_translationService == null)
                {
                    MessageBox.Show("❌ 翻译服务未初始化", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 设置当前翻译器类型
                _translationService.CurrentTranslatorType = "rwkv";

                // 检查当前翻译器是否可用
                if (_translationService.CurrentTranslator == null)
                {
                    MessageBox.Show("❌ RWKV翻译器未初始化，请检查配置", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 执行连接测试
                var result = await _translationService.CurrentTranslator.TestConnectionAsync();

                if (result == true)
                {
                    MessageBox.Show("✅ RWKV连接测试成功！", "测试结果", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("❌ RWKV连接测试失败，请检查服务是否运行", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "RWKV连接测试异常");
                MessageBox.Show($"❌ 连接测试异常: {ex.Message}", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                button.IsEnabled = true;
                button.Content = "🧪 测试连接";
            }
        }

        private void SaverwkvConfig(object sender, RoutedEventArgs e)
        {
            try
            {
                var useBatchTranslate = rwkvUseBatchTranslate.IsChecked ?? false;
                var currentUrl = rwkvApiUrl.Text;

                List<string> availableUrls = new List<string>
                {
                    "http://localhost:8000/translate/v1/batch-translate",
                    "http://localhost:8000/v1/chat/completions",
                    "http://localhost:8000/v2/chat/completions"
                };

                var configFile = Path.Combine(_configPath, "rwkv_api.json");
                if (File.Exists(configFile))
                {
                    try
                    {
                        var json = File.ReadAllText(configFile);
                        var existingConfig = JsonConvert.DeserializeObject<dynamic>(json);
                        var existingUrls = existingConfig?.available_api_urls as Newtonsoft.Json.Linq.JArray;
                        if (existingUrls != null && existingUrls.Count > 0)
                        {
                            foreach (var url in existingUrls)
                            {
                                var urlStr = url?.ToString();
                                if (!string.IsNullOrWhiteSpace(urlStr) && !availableUrls.Contains(urlStr))
                                {
                                    availableUrls.Add(urlStr);
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                if (!string.IsNullOrWhiteSpace(currentUrl) && !availableUrls.Contains(currentUrl))
                {
                    availableUrls.Add(currentUrl);
                }

                var config = new
                {
                    api_url = currentUrl,
                    default_model = "default",
                    temperature = rwkvTemperature.Value,
                    max_tokens = int.Parse(rwkvMaxTokens.Text),
                    skip_ssl = rwkvSkipSSL.IsChecked ?? false,
                    use_batch_translate = useBatchTranslate,
                    available_api_urls = availableUrls
                };

                var jsonContent = JsonConvert.SerializeObject(config, Formatting.Indented);

                Directory.CreateDirectory(_configPath);
                File.WriteAllText(configFile, jsonContent);

                // 同时保存到config.json
                _configurationManager.SaveConfig("rwkv_translator", new
                {
                    model = "default",
                    api_url = currentUrl,
                    temperature = rwkvTemperature.Value,
                    max_tokens = int.Parse(rwkvMaxTokens.Text),
                    use_batch_translate = useBatchTranslate
                });

                MessageBox.Show("✅ RWKV配置保存成功！", "保存结果", MessageBoxButton.OK, MessageBoxImage.Information);

                _translationService.ReinitializeTranslator("rwkv");

                _logger.LogInformation($"RWKV配置已保存: API地址={currentUrl}, 批量翻译={useBatchTranslate}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存RWKV配置失败");
                MessageBox.Show($"❌ 保存配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region llama_cpp配置
        private void LoadLlamaCppConfig()
        {
            try
            {
                var llamaConfig = _configurationManager.GetTranslatorConfig("llama_cpp");
                var configDict = _configurationManager.GetConfig<Dictionary<string, object>>("llama_cpp_translator", new Dictionary<string, object>()) ?? new Dictionary<string, object>();

                // API地址
                var apiUrl = configDict.GetValueOrDefault("api_url", "http://127.0.0.1:8080/v1/completions").ToString();
                llamaCppApiUrl.Text = apiUrl;

                // 加载可用的API地址列表
                LoadLlamaCppApiUrls(apiUrl);

                // 推理模式
                var modeStr = configDict.GetValueOrDefault("mode", "completions").ToString();
                llamaCppMode.SelectedIndex = modeStr.Equals("chat", StringComparison.OrdinalIgnoreCase) ? 1 : 0;

                // 对话模板
                var chatTemplate = configDict.GetValueOrDefault("chat_template", "rwkv-world").ToString();
                llamaCppChatTemplate.Text = chatTemplate;

                // 超时时间
                var timeout = configDict.GetValueOrDefault("timeout", 120).ToString();
                llamaCppTimeout.Text = timeout;

                // 翻译规则
                var translationRules = configDict.GetValueOrDefault("translation_rules", null);
                if (translationRules is Dictionary<string, object> rules)
                {
                    llamaCppEosToken.Text = rules.GetValueOrDefault("eos_token", "").ToString();
                    llamaCppEnSeparator.Text = UnescapeNewlines(rules.GetValueOrDefault("en_separator", "\\n\\nEnglish").ToString());
                    llamaCppZhSeparator.Text = UnescapeNewlines(rules.GetValueOrDefault("zh_separator", "\\n\\nChinese").ToString());
                    llamaCppJaSeparator.Text = UnescapeNewlines(rules.GetValueOrDefault("ja_separator", "\\n\\nJapanese").ToString());
                }
                else
                {
                    llamaCppEosToken.Text = "";
                    llamaCppEnSeparator.Text = "\\n\\nEnglish";
                    llamaCppZhSeparator.Text = "\\n\\nChinese";
                    llamaCppJaSeparator.Text = "\\n\\nJapanese";
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载llama_cpp配置失败");
                MessageBox.Show($"加载llama_cpp配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadLlamaCppApiUrls(string currentUrl)
        {
            try
            {
                llamaCppApiUrl.Items.Clear();

                var defaultUrls = new List<string>
                {
                    "http://127.0.0.1:8080/v1/completions",
                    "http://127.0.0.1:8080/v1/chat/completions"
                };

                var configDict = _configurationManager.GetConfig<Dictionary<string, object>>("llama_cpp_translator", null);
                if (configDict != null)
                {
                    var availableUrls = configDict.GetValueOrDefault("available_api_urls", null);
                    if (availableUrls is Newtonsoft.Json.Linq.JArray jArray)
                    {
                        foreach (var url in jArray)
                        {
                            var urlStr = url?.ToString();
                            if (!string.IsNullOrWhiteSpace(urlStr) && !defaultUrls.Contains(urlStr))
                            {
                                defaultUrls.Add(urlStr);
                            }
                        }
                    }
                }

                foreach (var url in defaultUrls)
                {
                    llamaCppApiUrl.Items.Add(url);
                }

                llamaCppApiUrl.Text = currentUrl;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载llama_cpp API地址列表失败");
            }
        }

        private async void TestLlamaCppConnection(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            button.IsEnabled = false;
            button.Content = "🔄 测试中...";

            try
            {
                if (_translationService == null)
                {
                    MessageBox.Show("❌ 翻译服务未初始化", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                _translationService.CurrentTranslatorType = "llama_cpp";

                if (_translationService.CurrentTranslator == null)
                {
                    MessageBox.Show("❌ llama_cpp翻译器未初始化，请检查配置", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var result = await _translationService.CurrentTranslator.TestConnectionAsync();

                if (result == true)
                {
                    MessageBox.Show("✅ llama_cpp连接测试成功！", "测试结果", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("❌ llama_cpp连接测试失败，请检查服务是否运行", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "llama_cpp连接测试异常");
                MessageBox.Show($"❌ 连接测试异常: {ex.Message}", "测试结果", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                button.IsEnabled = true;
                button.Content = "🧪 测试连接";
            }
        }

        private void SaveLlamaCppConfig(object sender, RoutedEventArgs e)
        {
            try
            {
                var currentUrl = llamaCppApiUrl.Text;
                var modeItem = llamaCppMode.SelectedItem as ComboBoxItem;
                var modeStr = modeItem?.Tag?.ToString() ?? "completions";
                var chatTemplate = llamaCppChatTemplate.Text;
                var timeout = int.Parse(llamaCppTimeout.Text);

                // 翻译规则
                var eosToken = llamaCppEosToken.Text;
                var enSeparator = EscapeNewlines(llamaCppEnSeparator.Text);
                var zhSeparator = EscapeNewlines(llamaCppZhSeparator.Text);
                var jaSeparator = EscapeNewlines(llamaCppJaSeparator.Text);

                // 收集可用的API地址
                List<string> availableUrls = new List<string>
                {
                    "http://127.0.0.1:8080/v1/completions",
                    "http://127.0.0.1:8080/v1/chat/completions"
                };

                var configDict = _configurationManager.GetConfig<Dictionary<string, object>>("llama_cpp_translator", null);
                if (configDict != null)
                {
                    var existingUrls = configDict.GetValueOrDefault("available_api_urls", null);
                    if (existingUrls is Newtonsoft.Json.Linq.JArray jArray)
                    {
                        foreach (var url in jArray)
                        {
                            var urlStr = url?.ToString();
                            if (!string.IsNullOrWhiteSpace(urlStr) && !availableUrls.Contains(urlStr))
                            {
                                availableUrls.Add(urlStr);
                            }
                        }
                    }
                }

                if (!string.IsNullOrWhiteSpace(currentUrl) && !availableUrls.Contains(currentUrl))
                {
                    availableUrls.Add(currentUrl);
                }

                var config = new Dictionary<string, object>
                {
                    ["type"] = "llama_cpp",
                    ["api_url"] = currentUrl,
                    ["model"] = "",
                    ["timeout"] = timeout,
                    ["mode"] = modeStr,
                    ["chat_template"] = chatTemplate,
                    ["available_api_urls"] = availableUrls,
                    ["translation_rules"] = new Dictionary<string, object>
                    {
                        ["eos_token"] = eosToken,
                        ["en_separator"] = enSeparator,
                        ["zh_separator"] = zhSeparator,
                        ["ja_separator"] = jaSeparator
                    }
                };

                _configurationManager.SaveConfig("llama_cpp_translator", config);

                MessageBox.Show("✅ llama_cpp配置保存成功！", "保存结果", MessageBoxButton.OK, MessageBoxImage.Information);

                _translationService.ReinitializeTranslator("llama_cpp");

                _logger.LogInformation($"llama_cpp配置已保存: API地址={currentUrl}, 模式={modeStr}, 对话模板={chatTemplate}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存llama_cpp配置失败");
                MessageBox.Show($"❌ 保存配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 将文本中的 \n 转义序列转换为实际换行符
        /// </summary>
        private static string UnescapeNewlines(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return text.Replace("\\n", "\n");
        }

        /// <summary>
        /// 将实际换行符转换为 \n 转义序列（用于显示和存储）
        /// </summary>
        private static string EscapeNewlines(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return text.Replace("\n", "\\n").Replace("\r", "");
        }
        #endregion

        #region 环境变量操作
        private void SaveApiKeyToEnvironment(string keyName, string apiKey)
        {
            try
            {
                if (!string.IsNullOrEmpty(apiKey))
                {
                    // 保存到用户环境变量
                    Environment.SetEnvironmentVariable(keyName, apiKey, EnvironmentVariableTarget.User);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存环境变量失败: {ex.Message}");
            }
        }

        private string LoadApiKeyFromEnvironment(string keyName)
        {
            try
            {
                return Environment.GetEnvironmentVariable(keyName, EnvironmentVariableTarget.User) ?? "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"读取环境变量失败: {ex.Message}");
                return "";
            }
        }

        private void RemoveApiKeyFromEnvironment(string keyName)
        {
            try
            {
                Environment.SetEnvironmentVariable(keyName, null, EnvironmentVariableTarget.User);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"删除环境变量失败: {ex.Message}");
            }
        }
        #endregion

        #region 通用功能
        private void OpenConfigDirectory(object sender, RoutedEventArgs e)
        {
            try
            {
                Directory.CreateDirectory(_configPath);
                Process.Start("explorer.exe", _configPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开配置目录失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ResetAllConfigs(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("确定要重置所有配置吗？这将删除所有已保存的API配置。",
                                       "确认重置", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (Directory.Exists(_configPath))
                    {
                        Directory.Delete(_configPath, true);
                    }

                    // 重新加载默认配置
                    LoadAllConfigs();

                    MessageBox.Show("✅ 所有配置已重置！", "重置完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"重置配置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CloseWindow(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        protected override void OnClosed(EventArgs e)
        {
            _httpClient?.Dispose();
            base.OnClosed(e);
        }
        #endregion
    }
}
