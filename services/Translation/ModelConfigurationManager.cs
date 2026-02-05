using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.Translation
{
    /// <summary>
    /// 模型配置管理器，负责管理所有AI引擎的模型配置
    /// 配置保存在本地jsonl文件中，支持手动添加
    /// </summary>
    public class ModelConfigurationManager
    {
        private readonly ILogger<ModelConfigurationManager> _logger;
        private readonly string _configFilePath;
        private Dictionary<string, List<ModelConfig>> _modelConfigs;

        public ModelConfigurationManager(ILogger<ModelConfigurationManager> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "models_config.jsonl");
            _modelConfigs = new Dictionary<string, List<ModelConfig>>();
            
            LoadConfiguration();
        }

        /// <summary>
        /// 加载模型配置
        /// </summary>
        private void LoadConfiguration()
        {
            try
            {
                _modelConfigs.Clear();

                if (!File.Exists(_configFilePath))
                {
                    _logger.LogInformation($"模型配置文件不存在，将创建新文件: {_configFilePath}");
                    InitializeDefaultConfiguration();
                    return;
                }

                var lines = File.ReadAllLines(_configFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    try
                    {
                        var modelConfig = JsonSerializer.Deserialize<ModelConfig>(line);
                        if (modelConfig != null && !string.IsNullOrEmpty(modelConfig.Engine))
                        {
                            if (!_modelConfigs.ContainsKey(modelConfig.Engine))
                            {
                                _modelConfigs[modelConfig.Engine] = new List<ModelConfig>();
                            }
                            _modelConfigs[modelConfig.Engine].Add(modelConfig);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex, $"解析模型配置行失败: {line}");
                    }
                }

                _logger.LogInformation($"模型配置加载完成，共加载 {_modelConfigs.Sum(x => x.Value.Count)} 个模型配置");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载模型配置失败");
                _modelConfigs = new Dictionary<string, List<ModelConfig>>();
            }
        }

        /// <summary>
        /// 初始化默认配置
        /// </summary>
        private void InitializeDefaultConfiguration()
        {
            var defaultConfigs = new List<ModelConfig>
            {
                new ModelConfig
                {
                    Engine = "rwkv",
                    ModelName = "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118",
                    IsDefault = true,
                    Description = "RWKV v7 G1c 1.5B - 专用翻译模型"
                },
                new ModelConfig
                {
                    Engine = "rwkv",
                    ModelName = "rwkv7-g1d-2.9b-20260131-ctx8192",
                    IsDefault = false,
                    Description = "RWKV v7 G1 1.5B - 翻译模型"
                },
                new ModelConfig
                {
                    Engine = "rwkv",
                    ModelName = "rwkv7-g1c-13.3b-20251231-ctx8192",
                    IsDefault = false,
                    Description = "RWKV7 G1c 13.3B - 大型翻译模型"
                }
            };

            foreach (var config in defaultConfigs)
            {
                AddModelConfig(config, saveToFile: false);
            }

            SaveConfiguration();
            _logger.LogInformation("默认模型配置已初始化");
        }

        /// <summary>
        /// 获取指定引擎的所有模型配置
        /// </summary>
        public List<ModelConfig> GetModelConfigs(string engine)
        {
            if (_modelConfigs.TryGetValue(engine, out var configs))
            {
                return new List<ModelConfig>(configs);
            }
            return new List<ModelConfig>();
        }

        /// <summary>
        /// 获取指定引擎的模型名称列表
        /// </summary>
        public List<string> GetModelNames(string engine)
        {
            return GetModelConfigs(engine)
                .Select(x => x.ModelName)
                .ToList();
        }

        /// <summary>
        /// 获取指定引擎的默认模型
        /// </summary>
        public string GetDefaultModel(string engine)
        {
            var configs = GetModelConfigs(engine);
            var defaultModel = configs.FirstOrDefault(x => x.IsDefault);
            if (defaultModel != null)
            {
                return defaultModel.ModelName;
            }
            return configs.FirstOrDefault()?.ModelName ?? string.Empty;
        }

        /// <summary>
        /// 添加模型配置
        /// </summary>
        public void AddModelConfig(ModelConfig config, bool saveToFile = true)
        {
            try
            {
                if (config == null || string.IsNullOrEmpty(config.Engine) || string.IsNullOrEmpty(config.ModelName))
                {
                    _logger.LogWarning("模型配置无效，跳过添加");
                    return;
                }

                if (!_modelConfigs.ContainsKey(config.Engine))
                {
                    _modelConfigs[config.Engine] = new List<ModelConfig>();
                }

                var existingConfig = _modelConfigs[config.Engine]
                    .FirstOrDefault(x => x.ModelName == config.ModelName);

                if (existingConfig != null)
                {
                    existingConfig.Description = config.Description;
                    existingConfig.IsDefault = config.IsDefault;
                    existingConfig.UpdatedAt = DateTime.Now;
                    _logger.LogInformation($"更新模型配置: {config.Engine}/{config.ModelName}");
                }
                else
                {
                    config.CreatedAt = DateTime.Now;
                    config.UpdatedAt = DateTime.Now;
                    _modelConfigs[config.Engine].Add(config);
                    _logger.LogInformation($"添加模型配置: {config.Engine}/{config.ModelName}");
                }

                if (saveToFile)
                {
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "添加模型配置失败");
            }
        }

        /// <summary>
        /// 删除模型配置
        /// </summary>
        public void RemoveModelConfig(string engine, string modelName)
        {
            try
            {
                if (_modelConfigs.TryGetValue(engine, out var configs))
                {
                    var config = configs.FirstOrDefault(x => x.ModelName == modelName);
                    if (config != null)
                    {
                        configs.Remove(config);
                        SaveConfiguration();
                        _logger.LogInformation($"删除模型配置: {engine}/{modelName}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除模型配置失败");
            }
        }

        /// <summary>
        /// 设置默认模型
        /// </summary>
        public void SetDefaultModel(string engine, string modelName)
        {
            try
            {
                if (_modelConfigs.TryGetValue(engine, out var configs))
                {
                    foreach (var config in configs)
                    {
                        config.IsDefault = (config.ModelName == modelName);
                    }
                    SaveConfiguration();
                    _logger.LogInformation($"设置默认模型: {engine}/{modelName}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "设置默认模型失败");
            }
        }

        /// <summary>
        /// 保存配置到文件
        /// </summary>
        private void SaveConfiguration()
        {
            try
            {
                var directory = Path.GetDirectoryName(_configFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = false,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var lines = new List<string>();
                foreach (var kvp in _modelConfigs)
                {
                    foreach (var config in kvp.Value)
                    {
                        var json = JsonSerializer.Serialize(config, options);
                        lines.Add(json);
                    }
                }

                File.WriteAllLines(_configFilePath, lines);
                _logger.LogInformation($"模型配置已保存到: {_configFilePath}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存模型配置失败");
            }
        }

        /// <summary>
        /// 重新加载配置
        /// </summary>
        public void ReloadConfiguration()
        {
            LoadConfiguration();
        }

        /// <summary>
        /// 获取所有支持的引擎类型
        /// </summary>
        public List<string> GetSupportedEngines()
        {
            return new List<string> { "rwkv" };
        }
    }

    /// <summary>
    /// 模型配置类
    /// </summary>
    public class ModelConfig
    {
        public string Engine { get; set; } = string.Empty;
        public string ModelName { get; set; } = string.Empty;
        public bool IsDefault { get; set; } = false;
        public string Description { get; set; } = string.Empty;
        public DateTime CreatedAt { get; set; } = DateTime.Now;
        public DateTime UpdatedAt { get; set; } = DateTime.Now;
    }
}
