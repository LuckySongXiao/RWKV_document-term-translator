# C# WPF版本文档翻译助手

## 概述

这是一个基于RWKV翻译模型和C# WPF技术栈开发的文档术语翻译助手，旨在解决海量技术文档翻译方面的问题。该版本提供了更好的Windows兼容性和更流畅的用户体验。

## 主要特性

### 🎯 技术优势
- **原生Windows支持**：基于WPF，完美兼容Windows 10/11
- **现代化界面**：Material Design风格，响应式布局
- **异步处理**：避免UI阻塞，提供流畅的用户体验
- **稳定可靠**：成熟的.NET生态系统，异常处理完善

### 🚀 功能特性
- **多格式支持**：Word (.docx)、Excel (.xlsx)
- **多引擎支持**：RWKV
- **术语库管理**：专业术语一致性翻译
- **实时日志**：翻译过程可视化监控
- **进度显示**：实时翻译进度反馈

## 系统要求

### 必需组件
1. **Windows 10/11** (x64)
2. **.NET 6.0 Runtime** 或更高版本
3. **Visual Studio 2022** (开发环境，可选)

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
```bash
# 使用提供的批处理脚本下载
.\download_cpu.bat          # CPU版本
.\download_libtorch.bat     # LibTorch版本（需要CUDA）
```

或手动下载：
- **CPU版本**: 从llama.cpp官方发布页面下载
- **CUDA版本**: 从llama.cpp官方发布页面下载（需要NVIDIA GPU）

#### 下载RWKV模型
模型请使用RWKV官方提供的Translate系列的.safetensors后缀格式的模型。

具体模型需求可以联系RWKV官方社区技术：
- GitHub: https://github.com/Alic-Li
- HuggingFace: https://huggingface.co/Alic-Li

**模型下载地址**:
- **rwkv_lightning推理引擎**: https://www.modelscope.cn/models/shoumenchougou/RWKV-7-World-ST/files
- **llama_cpp推理引擎(GGUF)**: https://www.modelscope.cn/models/shoumenchougou/RWKV7-G1e-1.5B-GGUF/files

### 3. 配置API
复制并修改配置文件：
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
┌─────────────────────────────────────────────────────────┐
│                    C# WPF 应用程序                      │
│                                                         │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │   用户界面    │  │  翻译服务层   │  │  本地推理引擎  │  │
│  │              │  │              │  │              │  │
│  │ - MainWindow │  │ - Translation│  │ - LlamaCpp   │  │
│  │ - Windows/   │◄►│   Service    │◄►│   Process    │  │
│  │ - 配置窗口    │  │ - BaseTrans  │  │ - RWKV       │  │
│  └──────────────┘  │   lator      │  │   Lightning  │  │
│                    │ - RWKVTrans  │  └──────────────┘  │
│                    │   lator      │                    │
│                    └──────────────┘                    │
└─────────────────────────────────────────────────────────┘
```

### 核心组件

#### 1. 用户界面层
- **MainWindow**: 主窗口，翻译任务管理
- **TranslationAssistantWindow**: 翻译助手窗口
- **OutputConfigWindow**: 输出配置窗口
- **Windows/**: 各类配置和编辑窗口
  - EngineConfigWindow: 引擎配置
  - ModelSelectWindow: 模型选择
  - TerminologyEditorWindow: 术语编辑
  - TranslationRulesWindow: 翻译规则

#### 2. 翻译服务层 (services/Translation/)
- **BaseTranslator**: 翻译器基类
- **RWKVTranslator**: RWKV翻译引擎实现
- **LlamaCppTranslator**: LlamaCpp翻译引擎实现
- **TranslationService**: 翻译服务协调器
- **DocumentProcessor**: 文档处理器
- **TermExtractor**: 术语提取器

#### 3. 本地推理服务 (services/RwkvLocal/)
- **LlamaCppProcessService**: LlamaCpp进程管理
- 本地模型加载和推理

## 使用指南

### 基本操作流程

1. **选择文档**
   - 点击"选择文件"按钮
   - 支持Word、Excel、PDF、PPT格式

2. **配置翻译**
   - 选择翻译方向（中文↔外语）
   - 选择目标语言
   - 配置翻译选项

3. **选择AI引擎**
   - RWKV：本地RWKV模型服务
   - LlamaCpp：本地LlamaCpp推理引擎

4. **开始翻译**
   - 点击"开始翻译"
   - 监控实时进度

### 高级功能

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
    "model": "RWKV_v7_G1c_1.5B_Translate_ctx4096_20260118"
  },
  "local_model": {
    "enabled": true,
    "engine_directory": "rwkv_lightning_libtorch_win",
    "models_directory": "rwkv_models"
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
```bash
# 确保本地模型服务已启动
# 检查config.json中的api_url配置是否正确
# 验证模型文件是否已正确下载
```

#### 3. 翻译引擎连接失败
**症状**：测试连接失败
**解决方案**：
- 检查API配置
- 验证网络连接
- 确认服务可用性

### 日志分析
程序运行时会生成翻译日志文件：
- `translation_log_*.txt`：翻译过程日志

## 开发指南

### 项目结构
```
RWKV_document-term-translator/
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
│   ├── EngineConfigWindow.xaml/.cs     # 引擎配置窗口
│   ├── ModelSelectWindow.xaml/.cs      # 模型选择窗口
│   ├── TerminologyEditorWindow.xaml/.cs # 术语编辑窗口
│   └── TranslationRulesWindow.xaml/.cs # 翻译规则窗口
├── services/
│   ├── Translation/                    # 翻译服务
│   │   ├── BaseTranslator.cs           # 翻译器基类
│   │   ├── RWKVTranslator.cs           # RWKV翻译器
│   │   ├── LlamaCppTranslator.cs       # LlamaCpp翻译器
│   │   ├── TranslationService.cs       # 翻译服务
│   │   ├── DocumentProcessor.cs        # 文档处理器
│   │   ├── TermExtractor.cs            # 术语提取器
│   │   └── ...                         # 其他处理器
│   └── RwkvLocal/                      # 本地推理服务
│       └── LlamaCppProcessService.cs   # LlamaCpp进程管理
├── data/
│   └── terminology.json                # 术语库
├── docs/                               # 文档
├── tests/                              # 测试项目
├── download_cpu.bat                    # 下载CPU版llama_cpp脚本
├── download_libtorch.bat               # 下载LibTorch脚本
└── README_CSHARP.md                    # 说明文档
```

### 重要说明

#### 需要从外部下载的内容
以下文件由于体积过大或包含敏感信息，不包含在Git仓库中，需要自行下载或配置：

1. **llama_cpp推理引擎**
   - 运行 `download_cpu.bat` 下载CPU版本
   - 或运行 `download_libtorch.bat` 下载CUDA版本
   - 下载后放置在 `llama_cpp/` 目录

2. **RWKV模型文件**
   - .safetensors格式的模型文件
   - 下载地址见"快速开始"部分
   - 放置在 `rwkv_models/` 目录

3. **LibTorch运行时**（如使用CUDA版本）
   - 运行 `download_libtorch.bat` 自动下载
   - 或手动下载libtorch-win-shared-with-deps

4. **配置文件**
   - 复制 `config_immersive_example.json` 为 `config.json`
   - 根据实际情况修改配置参数

#### 不上传到GitHub的内容
- `bin/` 和 `obj/`：编译输出
- `llama_cpp/`：推理引擎二进制文件
- `rwkv_models/`：模型文件
- `libtorch-*/`：LibTorch运行时
- `config.json` 和 `API_config/`：配置文件（可能包含敏感信息）
- `data/terminology.json`：用户术语库
- `*.db`：数据库文件
- `*.log` 和 `translation_log_*.txt`：日志文件
- `.trae/` 和 `.codebuddy/`：IDE配置

### 扩展开发

#### 添加新的翻译引擎
1. 在 `services/Translation/Translators/` 中创建新的翻译器类
2. 继承 `BaseTranslator` 基类
3. 实现必要的翻译接口
4. 在UI中添加对应的配置选项

#### 自定义界面主题
1. 修改`App.xaml`中的全局样式
2. 更新`MainWindow.xaml`中的颜色和布局
3. 添加主题切换功能

## 性能优化

### 建议配置
- **内存**：建议8GB以上
- **存储**：SSD硬盘，提升文件读写速度

### 优化技巧
1. **批量处理**：一次处理多个文本段
2. **缓存机制**：复用翻译结果
3. **并行处理**：多线程处理大文档

## 版本历史

### v3.1.0 (2025-01-17)
- 🎉 首个C# WPF版本发布
- ✨ 现代化界面设计
- 🚀 异步处理优化
- 🔧 Python桥接架构

## 技术支持

### 联系方式
- **问题反馈**：GitHub Issues
- **技术讨论**：项目Wiki
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