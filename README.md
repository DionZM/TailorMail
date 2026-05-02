# 千函 TailorMail

**千人千面的个性化邮件群发工具**

[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-0.9.1-green.svg)]

---

## 简介

千函（TailorMail）是一款 WPF 桌面应用，专为需要向多个收件人发送**个性化邮件**的场景设计。不同于普通的群发工具，千函的核心能力是**为每位收件人量身定制邮件内容**——就像裁缝（Tailor）为每位顾客量体裁衣一样。

### 核心特性

- **个性化变量替换** — 邮件模板支持 `{名称}`、`{简称}` 等内置占位符，也可自定义任意变量，发送时自动替换为每个收件人的对应值
- **差异化附件** — 每个收件人可以有专属附件，支持按文件名自动匹配
- **分组收件人管理** — 每个收件人可独立配置 TO/CC/BCC 邮箱，支持从 Excel 批量导入
- **双通道发送** — 支持 Outlook COM 和 SMTP (MailKit) 两种发送方式
- **6 步引导式工作流** — 收件选择 → 变量配置 → 模板撰写 → 附件匹配 → 效果预览 → 批量发送
- **发送进度追踪与失败重试** — 实时显示发送状态，支持一键重试失败项
- **富文本编辑器** — 内置格式工具栏，支持加粗、斜体、下划线、字号、颜色等
- **模板管理** — 保存/加载邮件模板，避免重复编辑
- **Excel 导入导出** — 收件人数据和变量数据均可通过 Excel 批量管理

## 界面预览

应用采用 6 步骤向导式导航：

| 步骤 | 功能 | 说明 |
|------|------|------|
| 1. 收件选择 | 收件人管理 | 分组管理、Excel 导入导出、全选/取消 |
| 2. 变量配置 | 自定义变量 | 添加/删除变量、Excel 导入变量值 |
| 3. 模板撰写 | 邮件编辑 | 富文本编辑器、变量插入、模板管理 |
| 4. 附件匹配 | 附件配置 | 公共附件 + 收件人专属附件、自动匹配 |
| 5. 效果预览 | 预览检查 | 逐个预览变量替换后的邮件效果 |
| 6. 批量发送 | 批量发送 | 进度追踪、失败重试、取消发送 |

## 技术栈

- **框架**: .NET 8 (WPF)
- **架构**: MVVM (CommunityToolkit.Mvvm)
- **UI 库**: WPF UI (Wpf.Ui)
- **邮件发送**: Outlook COM Interop / MailKit
- **Excel 处理**: EPPlus
- **依赖注入**: Microsoft.Extensions.DependencyInjection
- **数据存储**: SQLite (Microsoft.Data.Sqlite)

## 项目结构

```
TailorMail/
├── src/TailorMail/
│   ├── Models/                    # 数据模型
│   │   ├── Recipient.cs           # 收件人与分组模型
│   │   ├── MailTemplate.cs        # 邮件模板模型
│   │   ├── AppSettings.cs         # 应用配置模型
│   │   ├── AttachmentConfig.cs    # 附件配置模型
│   │   └── SendResult.cs          # 发送结果模型
│   ├── Services/                  # 服务层
│   │   ├── IDataService.cs        # 数据服务接口
│   │   ├── SqliteDataService.cs   # SQLite 数据服务实现`n│   │   ├── JsonDataService.cs     # JSON 数据服务（迁移用）
│   │   ├── IEmailSender.cs        # 邮件发送器接口
│   │   ├── OutlookEmailSender.cs  # Outlook COM 发送器
│   │   ├── SmtpEmailSender.cs     # SMTP 发送器
│   │   ├── AttachmentMatchService.cs # 附件自动匹配服务
│   │   └── AppLogger.cs           # 日志工具
│   ├── ViewModels/                # 视图模型
│   │   ├── RecipientsViewModel.cs # 收件人管理 VM
│   │   ├── VariablesViewModel.cs  # 变量管理 VM
│   │   ├── MailComposeViewModel.cs # 邮件撰写 VM
│   │   ├── AttachmentViewModel.cs # 附件管理 VM
│   │   ├── PreviewViewModel.cs    # 预览 VM
│   │   ├── SendViewModel.cs       # 发送 VM
│   │   └── SettingsViewModel.cs   # 设置 VM
│   ├── Views/                     # 视图
│   │   ├── RecipientsPage.xaml    # 收件人管理页面
│   │   ├── VariablesPage.xaml     # 变量管理页面
│   │   ├── MailComposePage.xaml   # 邮件撰写页面
│   │   ├── AttachmentPage.xaml    # 附件管理页面
│   │   ├── PreviewPage.xaml       # 预览页面
│   │   ├── SendPage.xaml          # 发送页面
│   │   └── ...                    # 对话框等
│   ├── Converters/                # WPF 值转换器
│   ├── Helpers/                   # 辅助工具类
│   ├── App.xaml.cs                # 应用入口
│   └── MainWindow.xaml.cs         # 主窗口
└── TailorMail.sln                 # 解决方案文件
```

## 快速开始

### 环境要求

- Windows 10/11
- .NET 8 SDK
- (可选) Microsoft Outlook 桌面版

### 构建与运行

```bash
# 克隆仓库
git clone https://github.com/your-username/TailorMail.git
cd TailorMail

# 还原依赖
dotnet restore

# 构建项目
dotnet build

# 运行应用
dotnet run --project src/TailorMail/TailorMail.csproj
```

### 发布为独立应用

```bash
dotnet publish src/TailorMail/TailorMail.csproj -c Release -r win-x64 --self-contained
```

## 使用指南

### 1. 配置发送方式

点击右上角菜单按钮，选择"设置"，选择邮件发送方式：

- **Outlook**: 需要本地安装 Outlook，无需额外配置
- **SMTP**: 需填写服务器地址、端口、用户名等信息（密码在发送时临时输入，不持久化）

### 2. 添加收件人

- 手动在表格中添加
- 从 Excel 文件导入（格式：名称 | 简称 | 收件人邮箱 | 抄送邮箱 | 密送邮箱 | 备注）
- 从其他分组复制

### 3. 设置变量

- 内置变量：`{名称}``{简称}`
- 自定义变量：添加后在表格中为每个收件人填写值
- 从 Excel 导入变量值（格式：名称 | 变量1 | 变量2 | ...）

### 4. 撰写邮件

- 使用富文本编辑器编辑邮件正文
- 点击"插入变量"按钮插入占位符
- 可保存为模板以便复用

### 5. 配置附件

- **公共附件**: 所有收件人共享的附件
- **收件人专属附件**: 每个收件人独立的附件
- **自动匹配**: 根据收件人名称在指定目录中自动查找匹配的文件

### 6. 预览与发送

- 在预览页面逐个检查邮件效果
- 确认无误后点击"开始发送"
- 发送过程中可查看实时进度，支持取消和失败重试

## 数据存储

应用数据存储在应用程序目录的 `data/` 文件夹中：

| 文件 | 内容 |
|------|------|
| `data/tailormail.db`| SQLite 数据库（收件人、设置、附件配置、邮件模板） |
|
|
|

日志文件存储在 `%LocalAppData%/TailorMail/Logs/` 目录下。

## 许可证

本项目基于 [Apache License 2.0](LICENSE) 开源。
