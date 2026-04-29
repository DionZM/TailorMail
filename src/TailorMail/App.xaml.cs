﻿﻿﻿﻿﻿﻿using System.IO;
using System.Windows;
using System.Windows.Threading;
using TailorMail.Services;
using Microsoft.Extensions.DependencyInjection;

namespace TailorMail;

/// <summary>
/// 应用程序主入口类，负责依赖注入容器配置、全局异常处理和主窗口启动。
/// 使用 <see cref="ServiceCollection"/> 构建 DI 容器，注册所有核心服务。
/// </summary>
public partial class App : System.Windows.Application
{
    /// <summary>
    /// 获取全局 DI 服务提供者，用于在非 DI 上下文中解析服务。
    /// </summary>
    public static IServiceProvider Services { get; private set; } = null!;

    /// <summary>
    /// 获取数据服务的快捷属性，等价于 GetRequiredService&lt;IDataService&gt;()。
    /// </summary>
    public static IDataService DataService => GetRequiredService<IDataService>();

    /// <summary>
    /// 从 DI 容器获取指定类型的服务实例。
    /// </summary>
    /// <typeparam name="T">要获取的服务类型。</typeparam>
    /// <returns>服务实例。</returns>
    public static T GetRequiredService<T>() where T : notnull
    {
        return (T)Services.GetRequiredService(typeof(T));
    }

    /// <summary>
    /// 应用程序启动时的初始化逻辑：
    /// 1. 注册三种全局异常处理器（UI 线程、非 UI 线程、Task 未观察异常）
    /// 2. 配置 DI 容器并注册所有服务
    /// 3. 创建并显示主窗口
    /// </summary>
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        ShutdownMode = ShutdownMode.OnMainWindowClose;

        // UI 线程未处理异常：记录日志并显示错误对话框，防止应用崩溃
        DispatcherUnhandledException += (s, args) =>
        {
            AppLogger.Error("UI线程未处理异常", args.Exception);
            MessageBox.Show($"未处理的异常: {args.Exception.Message}\n\n详细信息已记录到日志", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            args.Handled = true;
        };

        // 非 UI 线程未处理异常：记录日志并显示致命错误对话框
        AppDomain.CurrentDomain.UnhandledException += (s, args) =>
        {
            if (args.ExceptionObject is Exception ex)
            {
                AppLogger.Error("非UI线程未处理异常", ex);
                MessageBox.Show($"致命错误: {ex.Message}\n\n详细信息已记录到日志", "错误",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        };

        // Task 未观察异常：记录日志并标记为已观察，防止进程终止
        TaskScheduler.UnobservedTaskException += (s, args) =>
        {
            AppLogger.Error("Task未观察异常", args.Exception);
            args.SetObserved();
        };

        try
        {
            AppLogger.Info("应用启动");

            // 配置依赖注入容器
            var services = new ServiceCollection();
            services.AddSingleton<IDataService, JsonDataService>();    // JSON 文件数据服务
            services.AddSingleton<AttachmentMatchService>();           // 附件自动匹配服务
            services.AddSingleton<OutlookEmailSender>();               // Outlook 邮件发送器
            services.AddSingleton<SmtpEmailSender>();                  // SMTP 邮件发送器
            Services = services.BuildServiceProvider();
            AppLogger.Info("DI容器初始化完成");

            var mainWindow = new MainWindow();
            mainWindow.Show();
            AppLogger.Info("主窗口显示完成");
        }
        catch (Exception ex)
        {
            AppLogger.Error("应用启动失败", ex);
            MessageBox.Show($"启动失败: {ex.Message}\n\n详细信息已记录到日志", "错误",
                MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown();
        }
    }
}
