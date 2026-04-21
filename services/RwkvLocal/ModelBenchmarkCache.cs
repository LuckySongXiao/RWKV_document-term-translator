using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentTranslator.Helpers;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// 模型基准测试缓存服务
    /// 保存和加载模型的最大 batch size 测试结果
    /// </summary>
    public class ModelBenchmarkCache
    {
        private readonly ILogger _logger;
        private readonly string _cacheFilePath;
        private Dictionary<string, ModelBenchmarkData> _cache;

        public ModelBenchmarkCache(ILogger logger)
        {
            _logger = logger;
            _cacheFilePath = PathHelper.SafeCombine(AppDomain.CurrentDomain.BaseDirectory, "model_benchmark_cache.json");
            _cache = new Dictionary<string, ModelBenchmarkData>();
            
            LoadCache();
        }

        /// <summary>
        /// 加载缓存文件
        /// </summary>
        private void LoadCache()
        {
            try
            {
                if (File.Exists(_cacheFilePath))
                {
                    var json = File.ReadAllText(_cacheFilePath);
                    var cache = JsonSerializer.Deserialize<Dictionary<string, ModelBenchmarkData>>(json);
                    if (cache != null)
                    {
                        _cache = cache;
                        _logger.LogInformation($"模型基准测试缓存加载完成，共 {_cache.Count} 个模型");
                    }
                }
                else
                {
                    _logger.LogInformation("模型基准测试缓存文件不存在，将创建新缓存");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载模型基准测试缓存失败");
            }
        }

        /// <summary>
        /// 保存缓存到文件
        /// </summary>
        private void SaveCache()
        {
            try
            {
                var json = JsonSerializer.Serialize(_cache, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_cacheFilePath, json);
                _logger.LogInformation("模型基准测试缓存已保存");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存模型基准测试缓存失败");
            }
        }

        /// <summary>
        /// 获取模型的中 prompt 场景最大 batch size（作为默认并发数）
        /// </summary>
        /// <param name="modelPath">模型文件路径</param>
        /// <param name="gpuIndex">GPU 索引</param>
        /// <returns>最大 batch size，如果不存在或为 0 则返回 null</returns>
        public int? GetMaxBatchSize(string modelPath, int gpuIndex)
        {
            var key = GetModelKey(modelPath, gpuIndex);
            if (_cache.TryGetValue(key, out var data))
            {
                // 检查缓存是否过期（7天）
                if (data.ExpiresAt > DateTime.UtcNow)
                {
                    // 如果中 prompt 场景的 BSN 为 0，说明测试未正确完成，需要重新测试
                    if (data.MediumPromptMaxBatchSize <= 0)
                    {
                        _logger.LogInformation($"缓存的模型 BSN 数据无效（中=0）: {key}，需要重新测试");
                        _cache.Remove(key);
                        return null;
                    }
                    
                    _logger.LogInformation(
                        $"使用缓存的模型 BSN 数据: {key}, " +
                        $"短={data.ShortPromptMaxBatchSize}, " +
                        $"中={data.MediumPromptMaxBatchSize}, " +
                        $"长={data.LongPromptMaxBatchSize}");
                    // 返回中 prompt 场景的 BSN 作为默认并发数
                    return data.MediumPromptMaxBatchSize;
                }
                else
                {
                    _logger.LogInformation($"模型 BSN 缓存已过期: {key}");
                    _cache.Remove(key);
                }
            }
            return null;
        }

        /// <summary>
        /// 保存模型的三种场景的最大 batch size
        /// </summary>
        /// <param name="modelPath">模型文件路径</param>
        /// <param name="gpuIndex">GPU 索引</param>
        /// <param name="shortPromptMaxBatchSize">短 prompt 场景最大 batch size</param>
        /// <param name="mediumPromptMaxBatchSize">中 prompt 场景最大 batch size</param>
        /// <param name="longPromptMaxBatchSize">长 prompt 场景最大 batch size</param>
        public void SaveMaxBatchSize(string modelPath, int gpuIndex, 
            int shortPromptMaxBatchSize, int mediumPromptMaxBatchSize, int longPromptMaxBatchSize)
        {
            var key = GetModelKey(modelPath, gpuIndex);
            _cache[key] = new ModelBenchmarkData
            {
                ModelPath = modelPath,
                GpuIndex = gpuIndex,
                ShortPromptMaxBatchSize = shortPromptMaxBatchSize,
                MediumPromptMaxBatchSize = mediumPromptMaxBatchSize,
                LongPromptMaxBatchSize = longPromptMaxBatchSize,
                TestedAt = DateTime.UtcNow,
                ExpiresAt = DateTime.UtcNow.AddDays(7)
            };
            
            SaveCache();
            _logger.LogInformation(
                $"已保存模型 BSN 数据: {key}, " +
                $"短={shortPromptMaxBatchSize}, " +
                $"中={mediumPromptMaxBatchSize}, " +
                $"长={longPromptMaxBatchSize}");
        }

        /// <summary>
        /// 检查模型是否有有效的缓存数据
        /// </summary>
        public bool HasValidCache(string modelPath, int gpuIndex)
        {
            return GetMaxBatchSize(modelPath, gpuIndex).HasValue;
        }

        /// <summary>
        /// 生成模型缓存键
        /// </summary>
        private string GetModelKey(string modelPath, int gpuIndex)
        {
            var modelName = Path.GetFileNameWithoutExtension(modelPath);
            return $"{modelName}_gpu{gpuIndex}";
        }

        /// <summary>
        /// 清除所有缓存
        /// </summary>
        public void ClearCache()
        {
            _cache.Clear();
            SaveCache();
            _logger.LogInformation("模型基准测试缓存已清除");
        }
    }

    /// <summary>
    /// 模型基准测试数据
    /// </summary>
    public class ModelBenchmarkData
    {
        public string ModelPath { get; set; } = "";
        public int GpuIndex { get; set; }
        
        /// <summary>
        /// 短 prompt 场景的最大 batch size
        /// </summary>
        public int ShortPromptMaxBatchSize { get; set; }
        
        /// <summary>
        /// 中 prompt 场景的最大 batch size
        /// </summary>
        public int MediumPromptMaxBatchSize { get; set; }
        
        /// <summary>
        /// 长 prompt 场景的最大 batch size
        /// </summary>
        public int LongPromptMaxBatchSize { get; set; }
        
        public DateTime TestedAt { get; set; }
        public DateTime ExpiresAt { get; set; }
    }
}
