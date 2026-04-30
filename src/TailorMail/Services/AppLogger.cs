namespace TailorMail.Services;

/// <summary>
/// 应用程序日志工具类，提供线程安全的文件日志功能。
/// 日志文件按日期存储在 %LocalAppData%/TailorMail/Logs/ 目录下，
/// 文件名格式为 log_yyyyMMdd.txt。
/// 支持三种日志级别：INFO、WARN、ERROR，ERROR 级别可附加异常堆栈信息。
/// </summary>
public static class AppLogger
{
    /// <summary>
    /// 日志文件存储目录的绝对路径。
    /// </summary>
    private static readonly string LogDir = System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "TailorMail", "Logs");

    /// <summary>
    /// 写入锁对象，确保多线程环境下日志写入的线程安全。
    /// </summary>
    private static readonly object _lock = new();

    /// <summary>
    /// 获取当前日期对应的日志文件路径。每天一个日志文件。
    /// </summary>
    private static string LogFile => System.IO.Path.Combine(LogDir, $"log_{DateTime.Now:yyyyMMdd}.txt");

    /// <summary>
    /// 静态构造函数，在类首次使用时创建日志目录。
    /// </summary>
    static AppLogger()
    {
        try
        {
            System.IO.Directory.CreateDirectory(LogDir);
        }
        catch { }
    }

    /// <summary>
    /// 记录 INFO 级别的日志信息。
    /// </summary>
    /// <param name="message">日志消息内容。</param>
    public static void Info(string message)
    {
        Write("INFO", message);
    }

    /// <summary>
    /// 记录 WARN 级别的日志信息。
    /// </summary>
    /// <param name="message">日志消息内容。</param>
    public static void Warning(string message)
    {
        Write("WARN", message);
    }

    /// <summary>
    /// 记录 ERROR 级别的日志信息，可附加异常详情。
    /// 日志格式包含异常类型、消息、堆栈跟踪，以及内部异常信息（如有）。
    /// </summary>
    /// <param name="message">日志消息内容。</param>
    /// <param name="ex">异常对象，可为 null。若提供则附加异常类型、消息和堆栈跟踪。</param>
    public static void Error(string message, Exception? ex = null)
    {
        var sb = new System.Text.StringBuilder();
        sb.Append(message);
        if (ex != null)
        {
            sb.AppendLine();
            sb.Append($"  Exception: {ex.GetType().Name}: {ex.Message}");
            sb.AppendLine();
            sb.Append($"  StackTrace: {ex.StackTrace}");

            // 递归输出内部异常信息
            if (ex.InnerException != null)
            {
                sb.AppendLine();
                sb.Append($"  InnerException: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
                sb.AppendLine();
                sb.Append($"  InnerStackTrace: {ex.InnerException.StackTrace}");
            }
        }
        Write("ERROR", sb.ToString());
    }

    /// <summary>
    /// 将日志消息写入文件。使用锁确保线程安全。
    /// 日志格式：[HH:mm:ss.fff] [LEVEL] message
    /// </summary>
    /// <param name="level">日志级别标识（INFO/WARN/ERROR）。</param>
    /// <param name="message">日志消息内容。</param>
    private static void Write(string level, string message)
    {
        try
        {
            lock (_lock)
            {
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
                System.IO.File.AppendAllText(LogFile, line);
            }
        }
        catch { }
    }
}
