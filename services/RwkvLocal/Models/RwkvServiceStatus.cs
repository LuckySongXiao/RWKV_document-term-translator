using System;

namespace DocumentTranslator.Services.RwkvLocal.Models
{
    /// <summary>
    /// RWKV服务状态
    /// </summary>
    public enum ServiceState
    {
        /// <summary>
        /// 已停止
        /// </summary>
        Stopped,

        /// <summary>
        /// 启动中
        /// </summary>
        Starting,

        /// <summary>
        /// 运行中
        /// </summary>
        Running,

        /// <summary>
        /// 停止中
        /// </summary>
        Stopping,

        /// <summary>
        /// 错误
        /// </summary>
        Error
    }

    /// <summary>
    /// RWKV本地推理服务状态模型
    /// </summary>
    public class RwkvServiceStatus
    {
        /// <summary>
        /// 服务状态
        /// </summary>
        public ServiceState State { get; set; } = ServiceState.Stopped;

        /// <summary>
        /// 服务是否正在运行
        /// </summary>
        public bool IsRunning => State == ServiceState.Running;

        /// <summary>
        /// 状态显示文本
        /// </summary>
        public string StateDisplayText => State switch
        {
            ServiceState.Stopped => "⏹️ 已停止",
            ServiceState.Starting => "🟡 启动中...",
            ServiceState.Running => "🟢 运行中",
            ServiceState.Stopping => "🟠 停止中...",
            ServiceState.Error => "🔴 错误",
            _ => "❓ 未知"
        };

        /// <summary>
        /// 状态颜色
        /// </summary>
        public string StateColor => State switch
        {
            ServiceState.Stopped => "#666666",
            ServiceState.Starting => "#FFC107",
            ServiceState.Running => "#28A745",
            ServiceState.Stopping => "#FF9800",
            ServiceState.Error => "#DC3545",
            _ => "#999999"
        };

        /// <summary>
        /// 当前选中的GPU
        /// </summary>
        public GpuInfo? SelectedGpu { get; set; }

        /// <summary>
        /// 当前加载的模型
        /// </summary>
        public ModelInfo? LoadedModel { get; set; }

        /// <summary>
        /// 服务监听端口
        /// </summary>
        public int Port { get; set; } = 8000;

        /// <summary>
        /// 服务API地址
        /// </summary>
        public string ApiUrl => $"http://127.0.0.1:{Port}";

        /// <summary>
        /// 计算得出的最大并发数
        /// </summary>
        public int MaxConcurrency { get; set; }

        /// <summary>
        /// 当前进程ID（-1表示未启动）
        /// </summary>
        public int ProcessId { get; set; } = -1;

        /// <summary>
        /// 服务启动时间
        /// </summary>
        public DateTime? StartTime { get; set; }

        /// <summary>
        /// 运行时长
        /// </summary>
        public TimeSpan? RunningDuration => StartTime.HasValue ? DateTime.Now - StartTime.Value : null;

        /// <summary>
        /// 运行时长显示文本
        /// </summary>
        public string RunningDurationText
        {
            get
            {
                if (!StartTime.HasValue) return "-";
                var duration = DateTime.Now - StartTime.Value;
                if (duration.TotalHours >= 1)
                    return $"{(int)duration.TotalHours}h {duration.Minutes}m";
                else if (duration.TotalMinutes >= 1)
                    return $"{(int)duration.TotalMinutes}m {duration.Seconds}s";
                else
                    return $"{duration.Seconds}s";
            }
        }

        /// <summary>
        /// 最后一次心跳检测时间
        /// </summary>
        public DateTime? LastHeartbeat { get; set; }

        /// <summary>
        /// 心跳检测是否成功
        /// </summary>
        public bool IsHeartbeatOk => LastHeartbeat.HasValue && (DateTime.Now - LastHeartbeat.Value).TotalSeconds < 30;

        /// <summary>
        /// 错误信息
        /// </summary>
        public string? ErrorMessage { get; set; }

        /// <summary>
        /// 当前GPU显存使用情况
        /// </summary>
        public GpuInfo? CurrentGpuStatus { get; set; }

        /// <summary>
        /// 已处理的请求数
        /// </summary>
        public long ProcessedRequests { get; set; }

        /// <summary>
        /// 状态摘要（用于UI显示）
        /// </summary>
        public string Summary
        {
            get
            {
                if (State == ServiceState.Running)
                {
                    var gpuInfo = CurrentGpuStatus?.StatusInfo ?? "GPU信息获取中";
                    return $"{StateDisplayText} | 端口:{Port} | 并发:{MaxConcurrency} | {gpuInfo}";
                }
                else if (State == ServiceState.Error)
                {
                    return $"{StateDisplayText} | {ErrorMessage ?? "未知错误"}";
                }
                else
                {
                    return StateDisplayText;
                }
            }
        }

        /// <summary>
        /// 是否可以启动
        /// </summary>
        public bool CanStart => State == ServiceState.Stopped || State == ServiceState.Error;

        /// <summary>
        /// 是否可以停止
        /// </summary>
        public bool CanStop => State == ServiceState.Running;

        /// <summary>
        /// 重置状态
        /// </summary>
        public void Reset()
        {
            State = ServiceState.Stopped;
            ProcessId = -1;
            StartTime = null;
            LastHeartbeat = null;
            ErrorMessage = null;
            ProcessedRequests = 0;
        }
    }
}
