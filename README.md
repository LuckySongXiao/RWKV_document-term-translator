# C# WPF版本文档术语翻译助手

## 概述

基于RWKV翻译模型和C# WPF技术栈开发的文档术语翻译助手，支持双推理引擎（rwkv_lightning + llama_cpp），提供GPU/CPU双计算模式、BSZ并发上限自动测试、可编辑翻译规则等高级功能。

## 主要特性

### 技术优势
- **原生Windows支持**：基于WPF，完美兼容Windows 10/11
- **双推理引擎**：rwkv_lightning（NVIDIA GPU专属）+ llama_cpp（多后端）
- **GPU/CPU双模式**：自动检测GPU，无GPU时自动切换CPU模式
- **BSZ并发测试**：自动测试最大并发请求数，各推理工具差异化适配
- **可编辑翻译规则**：EOS停止符、语言分隔符均可自定义

### 功能特性
- **多格式支持**：Word (.docx)、Excel (.xlsx)
- **多引擎支持**：rwkv_lightning、llama-cuda、llama-sycl、llama-vulkan、llama-cpu
- **术语库管理**：专业术语一致性翻译
- **实时日志**：翻译过程可视化监控
- **进度显示**：实时翻译进度反馈
- **模型管理**：本地模型扫描、下载、格式转换

## 系统要求

### 必需组件
1. **Windows 10/11** (x64)
2. **.NET 6.0 Runtime** 或更高版本
3. **Visual Studio 2022** (开发环境，可选)

### GPU支持（可选）
- **NVIDIA GPU**：支持CUDA加速（rwkv_lightning / llama-cuda）
- **Intel GPU**：支持SYCL加速（llama-sycl，Arc独显/Xe核显）
- **AMD/通用GPU**：支持Vulkan加速（llama-vulkan）
- 无GPU时可使用CPU模式（llama-cpu）

## 快速开始

### 1. 环境准备
```bash
# 安装.NET SDK
# 下载地址：https://dotnet.microsoft.com/download

# 验证安装
dotnet --version
```

### 2. 下载必要依赖

#### 下载llama_cpp推理引擎
```
前往 https://github.com/ggml-org/llama.cpp/releases 下载最新版本：

- llama-bXXXX-bin-win-cpu-x64.zip
  解压后将所有文件复制到 llama_cpp/cpu/ 目录

- llama-bXXXX-bin-win-cuda-12.4-x64.zip（NVIDIA GPU，二选一）
  解压后将所有文件复制到 llama_cpp/cuda/ 目录

- llama-bXXXX-bin-win-cuda-13.1-x64.zip（NVIDIA GPU，二选一）
  解压后将所有文件复制到 llama_cpp/cuda/ 目录

- llama-bXXXX-bin-win-sycl-x64.zip（Intel GPU）
  解压后将所有文件复制到 llama_cpp/sycl/ 目录

- llama-bXXXX-bin-win-vulkan-x64.zip（AMD/通用GPU）
  解压后将所有文件复制到 llama_cpp/vulkan/ 目录
```

#### 下载rwkv_lightning_libtorch推理引擎
```
前往 https://github.com/Alic-Li/rwkv_lightning_libtorch/releases
下载 rwkv_lightning_libtorch2.10.0+cu132_sm75-120_Windows_amd64.zip
解压后将所有文件复制到 rwkv_lightning_libtorch_win/ 目录
将项目源码放置到 rwkv_lightning_libtorch/ 目录
```

#### 下载RWKV模型
模型请使用RWKV官方提供的Translate系列的.safetensors后缀格式的模型。

具体模型需求可以联系RWKV官方社区技术：
- GitHub: https://github.com/Alic-Li
- HuggingFace: https://huggingface.co/Alic-Li

**模型下载地址**:
- **rwkv_lightning推理引擎** (SafeTensors): https://www.modelscope.cn/models/shoumenchougou/RWKV-7-World-ST/files
- **llama_cpp推理引擎** (GGUF): https://www.modelscope.cn/models/shoumenchougou/RWKV7-G1e-1.5B-GGUF/files

### 3. 配置API
```bash
# 复制配置示例
copy config_immersive_example.json config.json

# 编辑config.json，填入你的API配置信息
```

### 4. 构建项目
```bash
# 运行构建脚本
build_csharp.bat
```

### 5. 运行程序
```bash
# 进入发布目录
cd publish

# 运行程序
DocumentTranslator.exe
```

## 架构设计

### 整体架构
```
┌──────────────────────────────────────────────────────────────┐
│                     C# WPF 应用程序                          │
│                                                              │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────────┐  │
│  │   用户界面    │  │  翻译服务层   │  │   本地推理引擎    │  │
│  │              │  │              │  │                  │  │
│  │ - MainWindow │  │ - Translation│  │ - LlamaCpp       │  │
│  │ - Windows/   │◄►│   Service    │◄►│   ProcessService │  │
│  │ - 配置窗口    │  │ - BaseTrans  │  │ - RwkvProcess   │  │
│  └──────────────┘  │   lator      │  │   Service        │  │
│                    │ - LlamaCpp   │  │ - Concurrency    │  │
│                    │   Translator │  │   Calculator     │  │
│                    │ - RWKVTrans  │  │ - GpuResource    │  │
│                    │   lator      │  │   Service        │  │
│                    └──────────────┘  └──────────────────┘  │
└──────────────────────────────────────────────────────────────┘
```

### 核心组件

#### 1. 用户界面层
- **MainWindow**: 主窗口，翻译任务管理、计算方式选择、推理工具选择
- **TranslationAssistantWindow**: 翻译助手窗口
- **OutputConfigWindow**: 输出配置窗口
- **Windows/**: 各类配置和编辑窗口
  - EngineConfigWindow: 引擎配置（RWKV_lightning设置 + llama_cpp设置）
  - ModelSelectWindow: 模型选择
  - TerminologyEditorWindow: 术语编辑
  - TranslationRulesWindow: 翻译规则

#### 2. 翻译服务层 (services/Translation/)
- **BaseTranslator**: 翻译器基类（含输出过滤、重复检测、异常翻译处理）
- **LlamaCppTranslator**: LlamaCpp翻译引擎（支持Completions/Chat双模式、RWKV官方续写格式、可配置翻译规则）
- **RWKVTranslator**: RWKV翻译引擎实现
- **TranslationService**: 翻译服务协调器
- **DocumentProcessor**: 文档处理器
- **TermExtractor**: 术语提取器

#### 3. 本地推理服务 (services/RwkvLocal/)
- **LlamaCppProcessService**: llama-server进程管理（支持CUDA/SYCL/Vulkan/CPU四种后端、BSZ并发测试）
- **RwkvProcessService**: rwkv_lightning进程管理（含BSZ并发测试）
- **ConcurrencyCalculator**: 并发上限计算服务（rwkv_lightning和llama_cpp各自适配的BSZ测试方案）
- **GpuResourceService**: GPU资源监控服务
- **ModelManagementService**: 模型管理服务（扫描、路径解析、格式转换）
- **ModelBenchmarkCache**: 模型基准测试缓存（7天有效期）

## 使用指南

### 基本操作流程

1. **选择计算方式**
   - GPU模式：显示GPU选择和GPU推理工具
   - CPU模式：仅显示CPU推理工具

2. **选择推理工具**
   - GPU模式可选：rwkv_lightning、benchmark、llama-cuda、llama-sycl、llama-vulkan
   - CPU模式可选：llama-cpu
   - 模型列表自动根据推理工具格式过滤

3. **选择模型并启动**
   - rwkv_lightning：选择.safetensors格式模型
   - llama_cpp：选择.gguf格式模型
   - 可勾选BSZ上限测试，自动测试最大并发数

4. **选择文档并翻译**
   - 点击"选择文件"按钮
   - 支持Word、Excel格式
   - 选择翻译方向和目标语言
   - 点击"开始翻译"，监控实时进度

### 推理工具对照表

| 推理工具 | 计算方式 | 模型格式 | 加速后端 | BSZ测试适配 |
|---------|---------|---------|---------|------------|
| rwkv_lightning | GPU (NVIDIA) | SafeTensors | libtorch/CUDA | 批量翻译API测试 |
| llama-cuda | GPU (NVIDIA) | GGUF | CUDA | Completions API测试，增量≥10 |
| llama-sycl | GPU (Intel) | GGUF | SYCL | Completions API测试，增量≥5 |
| llama-vulkan | GPU (AMD/通用) | GGUF | Vulkan | Completions API测试，增量≥5 |
| llama-cpu | CPU | GGUF | 无 | Completions API测试，增量≤2，上限=min(核心数,100) |

### 高级功能

#### BSZ并发上限测试
- 勾选"启用BSZ上限测试"后，服务启动成功后自动执行并发测试
- 测试三种场景：短prompt(~50 tokens)、中prompt(~500 tokens)、长prompt(~2000 tokens)
- 逐步增加并发请求数，直到GPU/CPU资源使用率达到上限
- GPU模式上限：GPU使用率≥95% 且 显存使用率≥95%
- CPU模式上限：CPU使用率≥95% 且 内存使用率≥90%
- 测试结果缓存7天，避免重复测试
- 各推理工具使用差异化的测试参数（增量、上限、监控方式）

#### 翻译规则配置（llama_cpp）
在"设置 → llama_cpp设置"中可配置：
- **EOS停止符号**：模型输出的结束标记
- **语言分隔符**：翻译prompt中目标语言的分隔格式
  - 默认格式遵循RWKV官方续写翻译规范：`Chinese:\n原文\n\nEnglish:`
  - 英文分隔符默认：`\n\nEnglish`
  - 中文分隔符默认：`\n\nChinese`
  - 日文分隔符默认：`\n\nJapanese`
- **对话模板**：rwkv-world、chatml、llama3等
- **API接口**：Completions模式(/v1/completions) 或 Chat模式(/v1/chat/completions)

#### 术语库管理
- 专业术语一致性翻译
- 支持自定义术语对照表
- 术语预处理优化

#### 输出选项
- 双语对照模式
- 纯翻译结果模式
- PDF格式导出

## 配置说明

### 应用配置
主配置文件为 `config.json`：

```json
{
  "rwkv_translator": {
    "type": "rwkv",
    "api_url": "http://127.0.0.1:8000/translate/v1/batch-translate",
    "model": "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118",
    "timeout": 60.0,
    "use_batch_translate": true,
    "available_api_urls": [
      "http://127.0.0.1:8000/translate/v1/batch-translate",
      "http://127.0.0.1:8000/v1/chat/completions"
    ]
  },
  "llama_cpp_translator": {
    "type": "llama_cpp",
    "api_url": "http://127.0.0.1:8080/v1/completions",
    "model": "",
    "timeout": 120,
    "mode": "completions",
    "chat_template": "rwkv-world",
    "translation_rules": {
      "eos_token": "",
      "en_separator": "\\n\\nEnglish",
      "zh_separator": "\\n\\nChinese",
      "ja_separator": "\\n\\nJapanese"
    }
  },
  "current_translator_type": "rwkv",
  "local_model": {
    "enabled": true,
    "engine_directory": "rwkv_lightning_libtorch_win",
    "models_directory": "rwkv_models",
    "default_port": 8000,
    "max_concurrency": 4
  }
}
```

### 术语库配置
术语库文件位于 `data/terminology.json` 和 `Windows/terminology.json`

## 故障排除

### 常见问题

#### 1. 程序无法启动
**症状**：双击exe文件无反应
**解决方案**：
```bash
# 检查.NET Runtime
dotnet --list-runtimes

# 如果缺少，请安装.NET 6.0 Runtime
```

#### 2. 翻译引擎无法连接
**症状**：翻译时报错"无法连接到翻译引擎"
**解决方案**：
- 确保本地模型服务已启动
- 检查config.json中的api_url配置是否正确
- 验证模型文件是否已正确下载

#### 3. llama-server启动失败
**症状**：llama-server启动超时或报错
**常见原因**：
- 模型文件路径含中文字符（程序会自动处理，但极端情况可能失败）
- GPU显存不足（CUDA out of memory）
- CUDA驱动版本不匹配
- 端口被占用
- GGUF模型文件损坏

#### 4. 翻译输出重复源语言文本
**症状**：Completions模式输出重复的原文而非翻译
**解决方案**：
- 确认翻译规则配置正确（语言分隔符格式）
- 程序已内置prompt echo剥离和stop sequence，如仍有问题请在llama_cpp设置中调整

#### 5. 无GPU时程序不可用
**解决方案**：
- 切换到CPU计算方式，使用llama-cpu推理工具
- 程序检测不到NVIDIA GPU时会自动切换到CPU模式

### 日志分析
程序运行时会生成翻译日志文件：
- `translation_log_*.txt`：翻译过程日志

llama-server的stderr输出中，仅包含error/fail/assert/CUDA/out of memory等关键词的以警告级别记录，其余调试信息以Debug级别记录。

## 开发指南

### 项目结构
```
document-term-translator/
├── DocumentTranslator.csproj           # 项目文件
├── MainWindow.xaml/.cs                 # 主窗口
├── TranslationAssistantWindow.xaml/.cs # 翻译助手窗口
├── OutputConfigWindow.xaml/.cs         # 输出配置窗口
├── App.xaml/.cs                        # 应用程序入口
├── Program.cs                          # 程序入口点
├── config.json                         # 应用配置（需自行配置）
├── config_immersive_example.json       # 配置示例
├── Helpers/
│   └── PathHelper.cs                   # 路径辅助类
├── Windows/
│   ├── EngineConfigWindow.xaml/.cs     # 引擎配置窗口（RWKV_lightning + llama_cpp）
│   ├── ModelSelectWindow.xaml/.cs      # 模型选择窗口
│   ├── TerminologyEditorWindow.xaml/.cs # 术语编辑窗口
│   └── TranslationRulesWindow.xaml/.cs # 翻译规则窗口
├── services/
│   ├── Translation/                    # 翻译服务
│   │   ├── Translators/
│   │   │   ├── BaseTranslator.cs       # 翻译器基类
│   │   │   ├── LlamaCppTranslator.cs   # LlamaCpp翻译器
│   │   │   └── RWKVTranslator.cs       # RWKV翻译器
│   │   ├── TranslationService.cs       # 翻译服务协调器
│   │   ├── DocumentProcessor.cs        # 文档处理器
│   │   └── TermExtractor.cs            # 术语提取器
│   └── RwkvLocal/                      # 本地推理服务
│       ├── LlamaCppProcessService.cs   # llama-server进程管理
│       ├── RwkvProcessService.cs       # rwkv_lightning进程管理
│       ├── ConcurrencyCalculator.cs    # 并发上限计算（BSZ测试）
│       ├── GpuResourceService.cs       # GPU资源监控
│       ├── ModelManagementService.cs   # 模型管理
│       └── ModelBenchmarkCache.cs      # 基准测试缓存
├── data/
│   └── terminology.json                # 术语库
├── llama_cpp/                          # llama.cpp推理引擎（需下载）
│   ├── cuda/                           # CUDA后端
│   ├── sycl/                           # SYCL后端
│   ├── vulkan/                         # Vulkan后端
│   └── cpu/                            # CPU后端
├── rwkv_lightning_libtorch_win/        # rwkv_lightning推理引擎（需下载）
├── rwkv_models/                        # 模型文件目录
└── README_CSHARP.md                    # 说明文档
```

### 重要说明

#### 需要从外部下载的内容
以下文件由于体积过大，不包含在Git仓库中，需要自行下载：

1. **llama_cpp推理引擎**
   - CUDA版本放置在 `llama_cpp/cuda/` 目录
   - SYCL版本放置在 `llama_cpp/sycl/` 目录
   - Vulkan版本放置在 `llama_cpp/vulkan/` 目录
   - CPU版本放置在 `llama_cpp/cpu/` 目录

2. **rwkv_lightning_libtorch推理引擎**
   - 放置在 `rwkv_lightning_libtorch_win/` 目录

3. **RWKV模型文件**
   - SafeTensors格式放置在 `rwkv_models/` 目录
   - GGUF格式放置在 `rwkv_models/` 目录（llama_cpp使用）

4. **配置文件**
   - 复制 `config_immersive_example.json` 为 `config.json`
   - 根据实际情况修改配置参数

#### 不上传到GitHub的内容
- `bin/` 和 `obj/`：编译输出
- `llama_cpp/`：推理引擎二进制文件
- `rwkv_models/`：模型文件
- `rwkv_lightning_libtorch_win/`：rwkv_lightning引擎
- `libtorch-*/`：LibTorch运行时
- `config.json` 和 `API_config/`：配置文件
- `data/terminology.json`：用户术语库
- `*.db`：数据库文件
- `*.log` 和 `translation_log_*.txt`：日志文件

### 扩展开发

#### 添加新的翻译引擎
1. 在 `services/Translation/Translators/` 中创建新的翻译器类
2. 继承 `BaseTranslator` 基类
3. 实现必要的翻译接口
4. 在UI中添加对应的配置选项

#### 添加新的llama_cpp后端
1. 在 `ModelManagementService.GetToolPath` 中添加工具路径映射
2. 在 `MainWindow.xaml.cs` 的 `RefreshInferenceTools` 中添加ComboBox项
3. 在 `ConcurrencyCalculator.GetLlamaCppToolConfig` 中添加差异化测试配置
4. 在 `LlamaCppProcessService.StartAsync` 中添加GPU卸载逻辑

## 性能优化

### 建议配置
- **内存**：建议8GB以上
- **存储**：SSD硬盘，提升文件读写速度
- **GPU**：NVIDIA GPU with CUDA（rwkv_lightning性能最优）

### 优化技巧
1. **BSZ并发测试**：启用后自动确定最优并发数，避免资源浪费或过载
2. **批量处理**：一次处理多个文本段
3. **缓存机制**：复用翻译结果，BSZ测试结果缓存7天
4. **并行处理**：多线程处理大文档

## 版本历史

### v3.2.0 (2026-04)
- 双推理引擎支持（rwkv_lightning + llama_cpp）
- GPU/CPU双计算模式，自动检测切换
- llama_cpp四种后端支持（CUDA/SYCL/Vulkan/CPU）
- BSZ并发上限测试，各推理工具差异化适配
- 可编辑翻译规则（EOS、语言分隔符）
- RWKV官方续写翻译格式支持
- llama-server stderr日志级别优化

### v3.1.0 (2025-01)
- 首个C# WPF版本发布
- 现代化界面设计
- 异步处理优化

## 技术支持

### 联系方式
- **问题反馈**：GitHub Issues
- **技术讨论**：RWKV官方社区或任意讨论群组
- **使用指南**：在线文档

### 贡献指南
欢迎提交Pull Request和Issue，共同完善项目功能。

---

**注意**

本C#版本是对Python tkinter版本的增强替代，提供了更好的Windows兼容性和用户体验。如果您在使用过程中遇到任何问题，请参考故障排除部分或联系技术支持。

模型请使用RWKV官方提供的Translate系列的.safetensors后缀格式的模型。
具体模型需求可以联系RWKV官方社区技术：https://github.com/Alic-Li  https://huggingface.co/Alic-Li
rwkv_lightning推理引擎对应的RWKV基模库的下载地址为：https://www.modelscope.cn/models/shoumenchougou/RWKV-7-World-ST/files
llama_cpp推理引擎对应的GGUF模型库的下载地址为：https://www.modelscope.cn/models/shoumenchougou/RWKV7-G1e-1.5B-GGUF/files
