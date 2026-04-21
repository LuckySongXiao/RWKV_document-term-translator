using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Microsoft.Extensions.Logging;
using DocumentTranslator.Helpers;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 配置管理器，负责加载和管理翻译相关的配置
    /// </summary>
    public class ConfigurationManager
    {
        private readonly ILogger<ConfigurationManager> _logger;
        private Dictionary<string, object> _config;
        private Dictionary<string, ApiConfiguration> _apiConfigurations;
        private readonly ModelConfigurationManager _modelConfigurationManager;

        public ConfigurationManager(ILogger<ConfigurationManager> logger, ModelConfigurationManager modelConfigurationManager)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _modelConfigurationManager = modelConfigurationManager ?? throw new ArgumentNullException(nameof(modelConfigurationManager));
            _config = new Dictionary<string, object>();
            _apiConfigurations = new Dictionary<string, ApiConfiguration>();
            
            LoadConfiguration();
            LoadApiConfigurations();
        }

        /// <summary>
        /// 加载主配置文件
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                if (File.Exists(configPath))
                {
                    var configJson = File.ReadAllText(configPath);
                    var config = JsonSerializer.Deserialize<Dictionary<string, object>>(configJson);
                    if (config != null)
                    {
                        _config = config;
                    }
                }
                _logger.LogInformation("主配置文件加载完成");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载主配置文件失败");
            }
        }

        /// <summary>
        /// 加载API配置文件
        /// </summary>
        private void LoadApiConfigurations()
        {
            try
            {
                var apiConfigDir = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config");
                if (!Directory.Exists(apiConfigDir))
                {
                    _logger.LogWarning($"API配置目录不存在: {apiConfigDir}");
                    return;
                }

                LoadApiConfiguration("rwkv", "rwkv_api.json");

                _logger.LogInformation($"API配置加载完成，共加载 {_apiConfigurations.Count} 个配置");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载API配置失败");
            }
        }

        /// <summary>
        /// 加载单个API配置
        /// </summary>
        private void LoadApiConfiguration(string serviceName, string fileName)
        {
            try
            {
                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config", fileName);
                if (File.Exists(configPath))
                {
                    var configJson = File.ReadAllText(configPath);
                    var config = JsonSerializer.Deserialize<ApiConfiguration>(configJson);
                    if (config != null)
                    {
                        _apiConfigurations[serviceName] = config;
                        _logger.LogInformation($"加载 {serviceName} API配置成功");
                    }
                }
                else
                {
                    _logger.LogWarning($"API配置文件不存在: {configPath}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"加载 {serviceName} API配置失败");
            }
        }

        /// <summary>
        /// 获取API密钥
        /// </summary>
        public string GetApiKey(string serviceName)
        {
            if (_apiConfigurations.TryGetValue(serviceName, out var config))
            {
                return config.ApiKey ?? string.Empty;
            }
            return string.Empty;
        }

        /// <summary>
        /// 获取配置项
        /// </summary>
        public T GetConfig<T>(string key, T defaultValue = default)
        {
            try
            {
                if (_config.TryGetValue(key, out var value))
                {
                    if (value is JsonElement jsonElement)
                    {
                        return JsonSerializer.Deserialize<T>(jsonElement.GetRawText());
                    }
                    return (T)Convert.ChangeType(value, typeof(T));
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"获取配置项 '{key}' 失败，使用默认值");
            }
            return defaultValue;
        }

        /// <summary>
        /// 获取翻译器配置
        /// </summary>
        public TranslatorConfiguration GetTranslatorConfig(string translatorType)
        {
            try
            {
                var configKey = $"{translatorType}_translator";
                var config = GetConfig<Dictionary<string, object>>(configKey, new Dictionary<string, object>());

                // 特殊处理RWKV：从rwkv_api.json读取api_url
                string apiUrl = null;
                if (translatorType == "rwkv")
                {
                    try
                    {
                        var rwkvConfigPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config", "rwkv_api.json");
                        if (File.Exists(rwkvConfigPath))
                        {
                            var rwkvJson = File.ReadAllText(rwkvConfigPath);
                            var rwkvConfig = JsonSerializer.Deserialize<Dictionary<string, object>>(rwkvJson);
                            if (rwkvConfig != null && rwkvConfig.TryGetValue("api_url", out var urlObj))
                            {
                                apiUrl = urlObj.ToString();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, "从rwkv_api.json读取api_url失败，使用config.json中的值");
                    }
                }

                return new TranslatorConfiguration
                {
                    Model = GetConfigValue(config, "model", GetDefaultModel(translatorType)),
                    Temperature = GetConfigValue(config, "temperature", 0.2f),
                    Timeout = GetConfigValue(config, "timeout", 60),
                    ApiUrl = apiUrl ?? GetConfigValue(config, "api_url", GetDefaultApiUrl(translatorType)),
                    ModelListTimeout = GetConfigValue(config, "model_list_timeout", 10),
                    TranslateTimeout = GetConfigValue(config, "translate_timeout", 60),
                    ApiVersion = GetConfigValue(config, "api_version", "v1"),
                    UseBatchTranslate = GetConfigValue(config, "use_batch_translate", false),
                    TopP = GetConfigValue(config, "top_p", 0.6f),
                    FrequencyPenalty = GetConfigValue(config, "frequency_penalty", 1.05f),
                    MaxTokens = GetConfigValue(config, "max_tokens", 2048),
                    TopK = GetConfigValue(config, "top_k", 20),
                    UseChimeraMode = GetConfigValue(config, "use_chimera_mode", false)
                };
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"获取 {translatorType} 翻译器配置失败");
                return new TranslatorConfiguration();
            }
        }

        /// <summary>
        /// 获取配置值
        /// </summary>
        private T GetConfigValue<T>(Dictionary<string, object> config, string key, T defaultValue)
        {
            try
            {
                if (config.TryGetValue(key, out var value))
                {
                    if (value is JsonElement jsonElement)
                    {
                        return JsonSerializer.Deserialize<T>(jsonElement.GetRawText());
                    }
                    return (T)Convert.ChangeType(value, typeof(T));
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, $"获取配置值 '{key}' 失败，使用默认值");
            }
            return defaultValue;
        }

        /// <summary>
        /// 获取默认模型
        /// </summary>
        private string GetDefaultModel(string translatorType)
        {
            return translatorType switch
            {
                "rwkv" => "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118",
                _ => "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118"
            };
        }

        /// <summary>
        /// 获取默认API URL
        /// </summary>
        private string GetDefaultApiUrl(string translatorType)
        {
            return translatorType switch
            {
                "rwkv" => "http://127.0.0.1:8000",
                _ => "http://127.0.0.1:8000"
            };
        }

        /// <summary>
        /// 保存配置
        /// </summary>
        public void SaveConfig(string key, object value)
        {
            try
            {
                _config[key] = value;
                
                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                
                var configJson = JsonSerializer.Serialize(_config, options);
                File.WriteAllText(configPath, configJson);
                
                _logger.LogInformation($"配置项 '{key}' 保存成功");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"保存配置项 '{key}' 失败");
            }
        }

        /// <summary>
        /// 保存API配置
        /// </summary>
        public void SaveApiConfig(string serviceName, string apiKey)
        {
            try
            {
                var config = new ApiConfiguration { ApiKey = apiKey };
                _apiConfigurations[serviceName] = config;

                var fileName = $"{serviceName}_api.json";

                var configPath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "API_config", fileName);
                var directory = Path.GetDirectoryName(configPath);

                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var configJson = JsonSerializer.Serialize(config, options);
                File.WriteAllText(configPath, configJson);

                _logger.LogInformation($"{serviceName} API配置保存成功");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"保存 {serviceName} API配置失败");
            }
        }

        /// <summary>
        /// 重新加载配置
        /// </summary>
        public void ReloadConfiguration()
        {
            LoadConfiguration();
            LoadApiConfigurations();
        }

        /// <summary>
        /// 获取所有支持的翻译器类型
        /// </summary>
        public List<string> GetSupportedTranslatorTypes()
        {
            return _modelConfigurationManager.GetSupportedEngines();
        }

        /// <summary>
        /// 获取指定引擎的模型列表
        /// </summary>
        public List<string> GetModelList(string engine)
        {
            return _modelConfigurationManager.GetModelNames(engine);
        }

        /// <summary>
        /// 添加模型配置
        /// </summary>
        public void AddModelConfig(string engine, string modelName, bool isDefault = false, string description = "")
        {
            var config = new ModelConfig
            {
                Engine = engine,
                ModelName = modelName,
                IsDefault = isDefault,
                Description = description
            };
            _modelConfigurationManager.AddModelConfig(config);
        }

        /// <summary>
        /// 删除模型配置
        /// </summary>
        public void RemoveModelConfig(string engine, string modelName)
        {
            _modelConfigurationManager.RemoveModelConfig(engine, modelName);
        }

        /// <summary>
        /// 设置默认模型
        /// </summary>
        public void SetDefaultModel(string engine, string modelName)
        {
            _modelConfigurationManager.SetDefaultModel(engine, modelName);
        }

        /// <summary>
        /// 重新加载模型配置
        /// </summary>
        public void ReloadModelConfiguration()
        {
            _modelConfigurationManager.ReloadConfiguration();
        }

        /// <summary>
        /// 检查翻译器是否已配置
        /// </summary>
        public bool IsTranslatorConfigured(string translatorType)
        {
            return translatorType switch
            {
                "rwkv" => !string.IsNullOrEmpty(GetTranslatorConfig("rwkv").ApiUrl),
                _ => false
            };
        }
    }

    /// <summary>
    /// API配置类
    /// </summary>
    public class ApiConfiguration
    {
        public string ApiKey { get; set; } = string.Empty;
    }

    /// <summary>
    /// 翻译器配置类
    /// </summary>
    public class TranslatorConfiguration
    {
        public string Model { get; set; } = string.Empty;
        public float Temperature { get; set; } = 0.2f;
        public int Timeout { get; set; } = 60;
        public string ApiUrl { get; set; } = string.Empty;
        public int ModelListTimeout { get; set; } = 10;
        public int TranslateTimeout { get; set; } = 60;
        public string ApiVersion { get; set; } = "v1";
        public bool UseBatchTranslate { get; set; } = false;
        public float TopP { get; set; } = 0.6f;
        public float FrequencyPenalty { get; set; } = 1.05f;
        public int MaxTokens { get; set; } = 2048;
        public int TopK { get; set; } = 20;
        public bool UseChimeraMode { get; set; } = false;
        public List<string> AvailableModels { get; set; } = new List<string>();
    }
}
