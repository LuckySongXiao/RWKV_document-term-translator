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
3. **Python 3.7+** (用于翻译引擎)
4. **Visual Studio 2022** (开发环境，可选)

### Python依赖
```bash
pip install -r requirements.txt
```

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
┌─────────────────┐    ┌─────────────────┐    ┌─────────────────┐
│   C# WPF GUI    │    │  Python Bridge  │    │ Python Backend  │
│                 │    │                 │    │                 │
│ - 用户界面      │◄──►│ - 进程通信      │◄──►│ - 翻译引擎      │
│ - 事件处理      │    │ - JSON序列化    │    │ - 文档处理      │
│ - 状态管理      │    │ - 异常处理      │    │ - 术语库管理    │
└─────────────────┘    └─────────────────┘    └─────────────────┘
```

### 核心组件

#### 1. MainWindow (主窗口)
- **职责**：用户界面展示和交互
- **特性**：
  - 响应式布局设计
  - 实时状态更新
  - 异步操作处理
  - 进度条显示

#### 2. PythonBridge (Python桥接)
- **职责**：C#与Python后端通信
- **特性**：
  - 进程间通信
  - JSON数据序列化
  - 异常处理和重试
  - 超时控制

#### 3. Python Scripts (Python脚本)
- **职责**：实际的翻译处理逻辑
- **特性**：
  - 复用现有Python代码
  - 标准化输入输出
  - 详细的错误报告
  - 进度回调支持

## 使用指南

### 基本操作流程

1. **选择文档**
   - 点击"🔍 选择文件"按钮
   - 支持Word、Excel格式

2. **配置翻译**
   - 选择翻译方向（中文↔外语）
   - 选择目标语言
   - 配置翻译选项

3. **选择AI引擎**
   - RWKV：企业RWKV服务

4. **测试连接**
   - 点击"🧪 测试连接"
   - 确保引擎可用

5. **开始翻译**
   - 点击"🚀 开始翻译"
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

### API配置
各AI引擎的配置文件位于 `API_config/` 目录：

```
API_config/
└── rwkv_api.json   # RWKV配置
```

### 术语库配置
术语库文件位于 `data/terminology.json`：

```json
{
  "英语": {
    "人工智能": "Artificial Intelligence",
    "机器学习": "Machine Learning"
  },
  "日语": {
    "人工智能": "人工知能",
    "机器学习": "機械学習"
  }
}
```

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

#### 2. Python脚本执行失败
**症状**：翻译时报错"未找到Python可执行文件"
**解决方案**：
```bash
# 确保Python在PATH中
python --version

# 或者修改PythonBridge.cs中的Python路径
```

#### 3. 翻译引擎连接失败
**症状**：测试连接失败
**解决方案**：
- 检查API密钥配置
- 验证网络连接
- 确认服务可用性

### 日志分析
程序运行时会生成以下日志文件：
- `translation.log`：Python翻译日志
- `application.log`：C#应用程序日志

## 开发指南

### 项目结构
```
RWKV_document-term-translator/
├── DocumentTranslator.csproj      # 项目文件
├── MainWindow.xaml               # 主窗口XAML
├── MainWindow.xaml.cs            # 主窗口代码
├── App.xaml                      # 应用程序XAML
├── App.xaml.cs                   # 应用程序代码
├── TranslationAssistantWindow.xaml   # 翻译助手窗口
├── TranslationAssistantWindow.xaml.cs
├── OutputConfigWindow.xaml       # 输出配置窗口
├── OutputConfigWindow.xaml.cs
├── Helpers/
│   └── PathHelper.cs             # 路径辅助类
├── Windows/
│   ├── EngineConfigWindow.xaml   # 引擎配置窗口
│   ├── EngineConfigWindow.xaml.cs
│   ├── ModelSelectWindow.xaml    # 模型选择窗口
│   ├── ModelSelectWindow.xaml.cs
│   ├── TerminologyEditorWindow.xaml  # 术语编辑器窗口
│   ├── TerminologyEditorWindow.xaml.cs
│   ├── TranslationRulesWindow.xaml   # 翻译规则窗口
│   └── TranslationRulesWindow.xaml.cs
├── services/
│   ├── Translation/              # 翻译服务
│   │   ├── BaseTranslator.cs
│   │   ├── RWKVTranslator.cs
│   │   ├── TranslationService.cs
│   │   └── ...
│   └── RwkvLocal/                # 本地RWKV服务
├── llama_cpp/                    # llama_cpp推理引擎（需下载）
│   ├── cpu/                      # CPU版本
│   └── cuda/                     # CUDA版本
├── config.json                   # 应用配置（需自行配置）
├── config_immersive_example.json # 配置示例
├── download_cpu.bat              # 下载CPU版llama_cpp脚本
├── download_libtorch.bat         # 下载LibTorch脚本
└── README_CSHARP.md              # 说明文档
```

### 重要说明

#### 需要从外部下载的内容
以下文件由于体积过大，不包含在Git仓库中，需要自行下载：

1. **llama_cpp推理引擎**
   - 运行 `download_cpu.bat` 下载CPU版本
   - 或运行 `download_libtorch.bat` 下载CUDA版本

2. **RWKV模型文件**
   - .safetensors格式的模型文件
   - 下载地址见"快速开始"部分

3. **LibTorch运行时**（如使用CUDA版本）
   - 运行 `download_libtorch.bat` 自动下载

### 扩展开发

#### 添加新的翻译引擎
1. 在Python后端添加引擎支持
2. 更新`PythonBridge.cs`中的引擎列表
3. 在UI中添加对应的按钮和配置

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