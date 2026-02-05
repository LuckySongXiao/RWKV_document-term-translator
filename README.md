# RWKV 术语文档翻译助手

一款基于 WPF 的智能文档翻译工具，支持术语库管理与多格式文档翻译，面向 Windows 平台。

## 功能特性

- 支持术语库管理与术语规则编辑
- 支持多种文档格式的翻译处理
- 支持 OCR 文字识别与文档处理
- 提供可配置的翻译输出设置

## 支持格式

- Word
- PDF
- Excel
- PowerPoint
- TXT

## 快速开始

### 环境准备

- Windows
- .NET 6 SDK

### 构建与运行

```bash
dotnet build
dotnet run
```

### 发布

```bash
dotnet publish -c Release -r win-x64
```

## 目录结构

- data：内置术语与示例数据
- docs：项目相关说明
- services：翻译与文档处理核心逻辑
- Windows：WPF 窗口与界面文件

## 开发者

- gitee: 糖醋鹦鹉
- github: LuckySongXiao

## 许可证

本项目使用 [LICENSE](LICENSE) 中所述许可证。
