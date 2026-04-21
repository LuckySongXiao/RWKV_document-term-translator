using System;

namespace DocumentTranslator.Services.RwkvLocal.Models
{
    /// <summary>
    /// GPU信息模型
    /// </summary>
    public class GpuInfo
    {
        /// <summary>
        /// GPU索引（从0开始）
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// GPU名称（如：NVIDIA GeForce RTX 4090）
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 总显存大小（字节）- 包含专属内存和共享内存
        /// </summary>
        public long TotalMemoryBytes { get; set; }

        /// <summary>
        /// 专属显存大小（字节）- 仅 GPU 专用内存
        /// </summary>
        public long DedicatedMemoryBytes { get; set; }

        /// <summary>
        /// 已使用显存大小（字节）
        /// </summary>
        public long UsedMemoryBytes { get; set; }

        /// <summary>
        /// 空闲显存大小（字节）
        /// </summary>
        public long FreeMemoryBytes { get; set; }

        /// <summary>
        /// GPU利用率百分比（0-100）
        /// </summary>
        public float GpuUtilization { get; set; }

        /// <summary>
        /// 显存利用率百分比（0-100）
        /// </summary>
        public float MemoryUtilization { get; set; }

        /// <summary>
        /// 总显存大小（GB）
        /// </summary>
        public double TotalMemoryGB => TotalMemoryBytes / (1024.0 * 1024.0 * 1024.0);

        /// <summary>
        /// 专属显存大小（GB）
        /// </summary>
        public double DedicatedMemoryGB => DedicatedMemoryBytes / (1024.0 * 1024.0 * 1024.0);

        /// <summary>
        /// 已使用显存大小（GB）
        /// </summary>
        public double UsedMemoryGB => UsedMemoryBytes / (1024.0 * 1024.0 * 1024.0);

        /// <summary>
        /// 空闲显存大小（GB）
        /// </summary>
        public double FreeMemoryGB => FreeMemoryBytes / (1024.0 * 1024.0 * 1024.0);

        /// <summary>
        /// 显示名称（包含索引和专属显存大小）
        /// </summary>
        public string DisplayName => $"[{Index}] {Name} ({DedicatedMemoryGB:F1}GB)";

        /// <summary>
        /// 详细状态信息（格式：已用/专属 GB）
        /// </summary>
        public string StatusInfo => $"{Name}: {UsedMemoryGB:F1}/{DedicatedMemoryGB:F1}GB ({MemoryUtilization:F0}%)";

        /// <summary>
        /// UUID（用于唯一标识GPU）
        /// </summary>
        public string? Uuid { get; set; }

        /// <summary>
        /// 驱动版本
        /// </summary>
        public string? DriverVersion { get; set; }

        /// <summary>
        /// CUDA计算能力版本
        /// </summary>
        public string? CudaComputeCapability { get; set; }

        public override string ToString()
        {
            return DisplayName;
        }
    }
}
