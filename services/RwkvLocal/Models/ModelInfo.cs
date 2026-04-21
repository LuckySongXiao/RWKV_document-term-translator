using System;
using System.IO;

namespace DocumentTranslator.Services.RwkvLocal.Models
{
    /// <summary>
    /// 模型格式类型
    /// </summary>
    public enum ModelFormat
    {
        /// <summary>
        /// PyTorch格式（.pth）- 需要转换
        /// </summary>
        PyTorch,

        /// <summary>
        /// SafeTensors格式（.safetensors 或 .st）- 可直接使用
        /// </summary>
        SafeTensors,

        /// <summary>
        /// GGUF格式（.gguf）- llama.cpp可直接使用
        /// </summary>
        GGUF,

        /// <summary>
        /// 未知格式
        /// </summary>
        Unknown
    }

    /// <summary>
    /// 模型状态
    /// </summary>
    public enum ModelState
    {
        /// <summary>
        /// 可用（可直接加载）
        /// </summary>
        Available,

        /// <summary>
        /// 需要转换（.pth格式）
        /// </summary>
        NeedsConversion,

        /// <summary>
        /// 转换中
        /// </summary>
        Converting,

        /// <summary>
        /// 转换失败
        /// </summary>
        ConversionFailed,

        /// <summary>
        /// 正在使用
        /// </summary>
        InUse,

        /// <summary>
        /// 错误
        /// </summary>
        Error
    }

    /// <summary>
    /// 模型信息模型
    /// </summary>
    public class ModelInfo
    {
        /// <summary>
        /// 模型文件完整路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 模型文件名（不含路径）
        /// </summary>
        public string FileName => Path.GetFileName(FilePath);

        /// <summary>
        /// 模型名称（不含扩展名）
        /// </summary>
        public string ModelName => Path.GetFileNameWithoutExtension(FilePath);

        /// <summary>
        /// 模型格式
        /// </summary>
        public ModelFormat Format { get; set; } = ModelFormat.Unknown;

        /// <summary>
        /// 模型状态
        /// </summary>
        public ModelState State { get; set; } = ModelState.Available;

        /// <summary>
        /// 文件大小（字节）
        /// </summary>
        public long FileSizeBytes { get; set; }

        /// <summary>
        /// 文件大小（GB）
        /// </summary>
        public double FileSizeGB => FileSizeBytes / (1024.0 * 1024.0 * 1024.0);

        /// <summary>
        /// 预估参数量（如：1.5B, 7B）
        /// </summary>
        public string? EstimatedParameters { get; set; }

        /// <summary>
        /// 上下文长度
        /// </summary>
        public int? ContextLength { get; set; }

        /// <summary>
        /// 预估显存需求（GB）
        /// </summary>
        public double? EstimatedMemoryGB { get; set; }

        /// <summary>
        /// 转换后的文件路径（仅当Format为PyTorch时有效）
        /// </summary>
        public string? ConvertedFilePath { get; set; }

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 显示名称（包含大小和状态）
        /// </summary>
        public string DisplayName
        {
            get
            {
                var status = State switch
                {
                    ModelState.Available => "✅",
                    ModelState.NeedsConversion => "🔄",
                    ModelState.Converting => "⏳",
                    ModelState.ConversionFailed => "❌",
                    ModelState.InUse => "🔧",
                    _ => "❓"
                };

                var sizeInfo = FileSizeGB >= 1 ? $"{FileSizeGB:F1}GB" : $"{FileSizeBytes / (1024.0 * 1024.0):F0}MB";
                var formatInfo = Format == ModelFormat.PyTorch ? "[需转换]" : 
                                 Format == ModelFormat.GGUF ? "[GGUF]" : "";

                return $"{status} {ModelName} ({sizeInfo}) {formatInfo}";
            }
        }

        /// <summary>
        /// 是否可直接使用
        /// </summary>
        public bool IsReady => State == ModelState.Available && (Format == ModelFormat.SafeTensors || Format == ModelFormat.GGUF);

        /// <summary>
        /// 是否需要转换
        /// </summary>
        public bool NeedsConversion => Format == ModelFormat.PyTorch && State == ModelState.NeedsConversion;

        /// <summary>
        /// 获取实际使用的模型路径（转换后或原始路径）
        /// </summary>
        public string GetUsableModelPath()
        {
            if (Format == ModelFormat.PyTorch && !string.IsNullOrEmpty(ConvertedFilePath) && File.Exists(ConvertedFilePath))
            {
                return ConvertedFilePath;
            }
            return FilePath;
        }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
