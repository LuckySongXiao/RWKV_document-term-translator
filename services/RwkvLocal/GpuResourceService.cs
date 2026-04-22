using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using DocumentTranslator.Services.RwkvLocal.Models;
using Microsoft.Extensions.Logging;

namespace DocumentTranslator.Services.RwkvLocal
{
    /// <summary>
    /// GPU资源检测服务
    /// 使用NVML（NVIDIA Management Library）检测GPU信息
    /// 提供WMI作为备用方案
    /// </summary>
    public class GpuResourceService : IDisposable
    {
        private readonly ILogger _logger;
        private bool _nvmlInitialized;
        private bool _disposed;

        // NVML常量
        private const int NVML_SUCCESS = 0;
        private const int NVML_DEVICE_NAME_BUFFER_SIZE = 96;
        private const int NVML_DEVICE_UUID_BUFFER_SIZE = 80;

        // NVML函数委托
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlInitDelegate();
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlShutdownDelegate();
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetCountDelegate(ref uint count);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetHandleByIndexDelegate(uint index, ref IntPtr handle);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetNameDelegate(IntPtr handle, [Out] char[] name, uint size);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetMemoryInfoDelegate(IntPtr handle, ref NvmlMemory memory);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetUtilizationRatesDelegate(IntPtr handle, ref NvmlUtilization utilization);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetUuidDelegate(IntPtr handle, [Out] char[] uuid, uint size);
        [UnmanagedFunctionPointer(CallingConvention.Cdecl)]
        private delegate int NvmlDeviceGetCudaComputeCapabilityDelegate(IntPtr handle, ref int major, ref int minor);

        [StructLayout(LayoutKind.Sequential)]
        private struct NvmlMemory
        {
            public ulong Total;
            public ulong Free;
            public ulong Used;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NvmlUtilization
        {
            public uint Gpu;
            public uint Memory;
        }

        private static IntPtr _nvmlDll = IntPtr.Zero;
        private static NvmlInitDelegate? _nvmlInit;
        private static NvmlShutdownDelegate? _nvmlShutdown;
        private static NvmlDeviceGetCountDelegate? _nvmlDeviceGetCount;
        private static NvmlDeviceGetHandleByIndexDelegate? _nvmlDeviceGetHandleByIndex;
        private static NvmlDeviceGetNameDelegate? _nvmlDeviceGetName;
        private static NvmlDeviceGetMemoryInfoDelegate? _nvmlDeviceGetMemoryInfo;
        private static NvmlDeviceGetUtilizationRatesDelegate? _nvmlDeviceGetUtilizationRates;
        private static NvmlDeviceGetUuidDelegate? _nvmlDeviceGetUuid;
        private static NvmlDeviceGetCudaComputeCapabilityDelegate? _nvmlDeviceGetCudaComputeCapability;
        private static readonly object _lock = new object();
        private static bool _dllLoaded = false;

        public GpuResourceService(ILogger logger)
        {
            _logger = logger;
            InitializeNvml();
        }

        /// <summary>
        /// 初始化NVML库
        /// </summary>
        private void InitializeNvml()
        {
            try
            {
                lock (_lock)
                {
                    if (_dllLoaded) 
                    {
                        _nvmlInitialized = true;
                        return;
                    }

                    // 尝试加载NVML库
                    _nvmlDll = LoadLibrary("nvml.dll");
                    if (_nvmlDll == IntPtr.Zero)
                    {
                        // 尝试从CUDA安装目录加载
                        var cudaPath = Environment.GetEnvironmentVariable("CUDA_PATH");
                        if (!string.IsNullOrEmpty(cudaPath))
                        {
                            var nvmlPath = System.IO.Path.Combine(cudaPath, "bin", "nvml.dll");
                            if (System.IO.File.Exists(nvmlPath))
                            {
                                _nvmlDll = LoadLibrary(nvmlPath);
                            }
                        }
                    }

                    if (_nvmlDll != IntPtr.Zero)
                    {
                        LoadNvmlFunctions();
                        var result = _nvmlInit?.Invoke() ?? -1;
                        _nvmlInitialized = result == NVML_SUCCESS;
                        _dllLoaded = true;

                        if (_nvmlInitialized)
                        {
                            _logger.LogInformation("NVML初始化成功");
                        }
                        else
                        {
                            _logger.LogWarning($"NVML初始化失败，错误码: {result}");
                        }
                    }
                    else
                    {
                        _logger.LogWarning("无法加载nvml.dll，将使用WMI作为备用方案");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "初始化NVML时发生异常");
            }
        }

        private void LoadNvmlFunctions()
        {
            _nvmlInit = GetDelegate<NvmlInitDelegate>("nvmlInit_v2");
            _nvmlShutdown = GetDelegate<NvmlShutdownDelegate>("nvmlShutdown");
            _nvmlDeviceGetCount = GetDelegate<NvmlDeviceGetCountDelegate>("nvmlDeviceGetCount_v2");
            _nvmlDeviceGetHandleByIndex = GetDelegate<NvmlDeviceGetHandleByIndexDelegate>("nvmlDeviceGetHandleByIndex_v2");
            _nvmlDeviceGetName = GetDelegate<NvmlDeviceGetNameDelegate>("nvmlDeviceGetName");
            _nvmlDeviceGetMemoryInfo = GetDelegate<NvmlDeviceGetMemoryInfoDelegate>("nvmlDeviceGetMemoryInfo");
            _nvmlDeviceGetUtilizationRates = GetDelegate<NvmlDeviceGetUtilizationRatesDelegate>("nvmlDeviceGetUtilizationRates");
            _nvmlDeviceGetUuid = GetDelegate<NvmlDeviceGetUuidDelegate>("nvmlDeviceGetUUID");
            _nvmlDeviceGetCudaComputeCapability = GetDelegate<NvmlDeviceGetCudaComputeCapabilityDelegate>("nvmlDeviceGetCudaComputeCapability");
        }

        private T? GetDelegate<T>(string functionName) where T : Delegate
        {
            var ptr = GetProcAddress(_nvmlDll, functionName);
            return ptr != IntPtr.Zero ? Marshal.GetDelegateForFunctionPointer<T>(ptr) : null;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)]
        private static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        /// <summary>
        /// 获取所有GPU信息（NVIDIA + Intel + AMD）
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        public async Task<List<GpuInfo>> GetAllGpusAsync()
        {
            return await Task.Run(() =>
            {
                var gpus = new List<GpuInfo>();

                // NVIDIA GPU：优先使用NVML，回退到WMI
                if (_nvmlInitialized)
                {
                    gpus = GetGpusViaNvml();
                }
                else
                {
                    gpus = GetGpusViaWmi("NVIDIA");
                }

                // Intel GPU（Arc独显 / Xe核显）
                var intelGpus = GetGpusViaWmi("Intel");
                gpus.AddRange(intelGpus);

                // AMD GPU（Radeon独显 / 核显）
                var amdGpus = GetGpusViaWmi("AMD");
                gpus.AddRange(amdGpus);

                // 重新编号
                for (int i = 0; i < gpus.Count; i++)
                {
                    gpus[i].Index = i;
                }

                return gpus;
            });
        }

        /// <summary>
        /// 通过NVML获取GPU信息
        /// </summary>
        private List<GpuInfo> GetGpusViaNvml()
        {
            var gpus = new List<GpuInfo>();

            try
            {
                uint count = 0;
                var result = _nvmlDeviceGetCount?.Invoke(ref count) ?? -1;
                if (result != NVML_SUCCESS)
                {
                    _logger.LogWarning($"获取GPU数量失败，错误码: {result}");
                    return gpus;
                }

                for (uint i = 0; i < count; i++)
                {
                    var gpu = GetGpuInfoByIndex(i);
                    if (gpu != null)
                    {
                        gpus.Add(gpu);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "通过NVML获取GPU信息时发生异常");
            }

            return gpus;
        }

        /// <summary>
        /// 获取指定索引的GPU信息
        /// </summary>
        private GpuInfo? GetGpuInfoByIndex(uint index)
        {
            try
            {
                IntPtr handle = IntPtr.Zero;
                var result = _nvmlDeviceGetHandleByIndex?.Invoke(index, ref handle) ?? -1;
                if (result != NVML_SUCCESS || handle == IntPtr.Zero)
                {
                    return null;
                }

                var gpu = new GpuInfo { Index = (int)index };

                // 获取名称
                var nameBuffer = new char[NVML_DEVICE_NAME_BUFFER_SIZE];
                result = _nvmlDeviceGetName?.Invoke(handle, nameBuffer, NVML_DEVICE_NAME_BUFFER_SIZE) ?? -1;
                if (result == NVML_SUCCESS)
                {
                    gpu.Name = new string(nameBuffer).TrimEnd('\0');
                }

                // 获取显存信息
                var memory = new NvmlMemory();
                result = _nvmlDeviceGetMemoryInfo?.Invoke(handle, ref memory) ?? -1;
                if (result == NVML_SUCCESS)
                {
                    // NVML 返回的 Total 是专属显存，不包含共享内存
                    gpu.DedicatedMemoryBytes = (long)memory.Total;
                    gpu.TotalMemoryBytes = (long)memory.Total; // 保持兼容
                    gpu.FreeMemoryBytes = (long)memory.Free;
                    // 使用 Total - Free 计算已用显存，比直接使用 memory.Used 更准确
                    // 因为 memory.Used 可能不包含 CUDA 上下文和驱动开销
                    gpu.UsedMemoryBytes = (long)(memory.Total - memory.Free);
                    // 显存使用率基于专属显存计算
                    gpu.MemoryUtilization = memory.Total > 0 ? (float)(memory.Total - memory.Free) / memory.Total * 100 : 0;
                }

                // 获取利用率
                var utilization = new NvmlUtilization();
                result = _nvmlDeviceGetUtilizationRates?.Invoke(handle, ref utilization) ?? -1;
                if (result == NVML_SUCCESS)
                {
                    gpu.GpuUtilization = utilization.Gpu;
                }

                // 获取UUID
                var uuidBuffer = new char[NVML_DEVICE_UUID_BUFFER_SIZE];
                result = _nvmlDeviceGetUuid?.Invoke(handle, uuidBuffer, NVML_DEVICE_UUID_BUFFER_SIZE) ?? -1;
                if (result == NVML_SUCCESS)
                {
                    gpu.Uuid = new string(uuidBuffer).TrimEnd('\0');
                }

                // 获取CUDA计算能力
                int major = 0, minor = 0;
                result = _nvmlDeviceGetCudaComputeCapability?.Invoke(handle, ref major, ref minor) ?? -1;
                if (result == NVML_SUCCESS)
                {
                    gpu.CudaComputeCapability = $"{major}.{minor}";
                }

                return gpu;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"获取GPU {index} 信息时发生异常");
                return null;
            }
        }

        /// <summary>
        /// 通过WMI获取GPU信息（备用方案）
        /// 支持NVIDIA、Intel、AMD三种品牌
        /// </summary>
        [System.Runtime.Versioning.SupportedOSPlatform("windows")]
        private List<GpuInfo> GetGpusViaWmi(string brand = "NVIDIA")
        {
            var gpus = new List<GpuInfo>();

            try
            {
                var scope = new System.Management.ManagementScope(@"\\.\root\cimv2");
                var query = new System.Management.ObjectQuery($"SELECT * FROM Win32_VideoController WHERE Name LIKE '%{brand}%'");
                using var searcher = new System.Management.ManagementObjectSearcher(scope, query);

                foreach (System.Management.ManagementObject obj in searcher.Get())
                {
                    var gpuName = obj["Name"]?.ToString() ?? $"Unknown {brand} GPU";
                    var gpu = new GpuInfo
                    {
                        Name = gpuName
                    };

                    // 尝试获取显存大小（AdapterRAM单位是字节）- WMI 返回的是专属显存
                    if (obj["AdapterRAM"] != null)
                    {
                        gpu.DedicatedMemoryBytes = Convert.ToInt64(obj["AdapterRAM"]);
                        gpu.TotalMemoryBytes = Convert.ToInt64(obj["AdapterRAM"]);
                    }

                    // WMI无法获取实时显存使用情况
                    gpu.FreeMemoryBytes = gpu.TotalMemoryBytes;
                    gpu.UsedMemoryBytes = 0;
                    gpu.MemoryUtilization = 0;
                    gpu.GpuUtilization = 0;

                    gpu.DriverVersion = obj["DriverVersion"]?.ToString();

                    // 根据品牌设置GPU类型标识
                    if (brand == "Intel")
                    {
                        gpu.CudaComputeCapability = "SYCL"; // 标记为SYCL可用
                    }
                    else if (brand == "AMD")
                    {
                        gpu.CudaComputeCapability = "Vulkan"; // 标记为Vulkan可用
                    }

                    gpus.Add(gpu);
                }

                if (gpus.Count > 0)
                {
                    _logger.LogInformation($"通过WMI检测到 {gpus.Count} 个{brand} GPU: {string.Join(", ", gpus.Select(g => g.Name))}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"通过WMI获取{brand} GPU信息时发生异常");
            }

            return gpus;
        }

        /// <summary>
        /// 刷新指定GPU的状态
        /// </summary>
        public async Task<GpuInfo?> RefreshGpuStatusAsync(int index)
        {
            return await Task.Run(() =>
            {
                if (_nvmlInitialized)
                {
                    return GetGpuInfoByIndex((uint)index);
                }
                return null;
            });
        }

        /// <summary>
        /// 检查是否有可用的NVIDIA GPU
        /// </summary>
        public async Task<bool> HasNvidiaGpuAsync()
        {
            var gpus = await GetAllGpusAsync();
            return gpus.Count > 0;
        }

        /// <summary>
        /// 获取默认GPU（第一个可用GPU）
        /// </summary>
        public async Task<GpuInfo?> GetDefaultGpuAsync()
        {
            var gpus = await GetAllGpusAsync();
            return gpus.Count > 0 ? gpus[0] : null;
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                // 不在这里关闭NVML，因为其他服务可能还在使用
            }
        }
    }
}
