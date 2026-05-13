using System.Threading.Channels;

namespace TailorMail.Services;

/// <summary>
/// M-02: 异步队列日志工具类。使用 Channel 实现生产者-消费者模式，
/// 后台单线程批量写入，避免多线程争用锁和频繁文件 I/O。
/// </summary>
public static class AppLogger
{
    private static readonly string LogDir = System.IO.Path.Combine(
        AppDomain.CurrentDomain.BaseDirectory, "Logs");

    private const int MaxLogFiles = 3;

    // M-02: Unbounded channel for async log queue
    private static readonly Channel<string> _logChannel = Channel.CreateUnbounded<string>(
        new UnboundedChannelOptions { SingleReader = true, SingleWriter = false });

    private static string LogFile => System.IO.Path.Combine(LogDir, $"log_{DateTime.Now:yyyyMMdd}.txt");

    static AppLogger()
    {
        try
        {
            System.IO.Directory.CreateDirectory(LogDir);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"AppLogger: 创建日志目录失败: {ex.Message}");
        }

        // M-02: Start background writer task
        _ = Task.Run(ProcessLogQueueAsync);
    }

    public static void Info(string message)
    {
        Write("INFO", message);
    }

    public static void Warning(string message)
    {
        Write("WARN", message);
    }

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
    /// Write directly to file, bypassing the async queue.
    /// Use only for crash/fatal scenarios where the process is about to terminate.
    /// </summary>
    public static void FlushSync(string message)
    {
        try
        {
            var line = $"[{DateTime.Now:HH:mm:ss.fff}] [FATAL] {message}{Environment.NewLine}";
            System.IO.File.AppendAllText(LogFile, line);
        }
        catch { }
    }

    private static void Write(string level, string message)
    {
        try
        {
            var line = $"[{DateTime.Now:HH:mm:ss.fff}] [{level}] {message}{Environment.NewLine}";
            _logChannel.Writer.TryWrite(line);
        }
        catch
        {
            // Channel closed or other error — silently ignore to avoid recursive logging
        }
    }

    /// <summary>
    /// M-02: Background consumer that batches log lines and writes to file.
    /// </summary>
    private static async Task ProcessLogQueueAsync()
    {
        var batch = new List<string>(16);

        try
        {
            await foreach (var line in _logChannel.Reader.ReadAllAsync())
            {
                batch.Add(line);

                // Drain any immediately available lines into the batch
                while (_logChannel.Reader.TryRead(out var additional))
                    batch.Add(additional);

                try
                {
                    System.IO.File.AppendAllLines(LogFile, batch);
                    CleanupOldLogs();
                }
                catch
                {
                    // File write failed — try to continue processing
                }

                batch.Clear();
            }

            // Channel closed — flush any remaining items
            _flushCompletion.TrySetResult();
        }
        catch (ChannelClosedException) { _flushCompletion.TrySetResult(); }
        catch (OperationCanceledException) { _flushCompletion.TrySetResult(); }
    }

    public static void Flush(int timeoutMs = 2000)
    {
        try
        {
            _logChannel.Writer.TryComplete();
            _flushCompletion.Task.Wait(timeoutMs);
        }
        catch { }
    }

    private static readonly TaskCompletionSource _flushCompletion = new();

    private static void CleanupOldLogs()
    {
        try
        {
            var logFiles = System.IO.Directory.GetFiles(LogDir, "log_*.txt")
                .OrderByDescending(f => f)
                .Skip(MaxLogFiles)
                .ToList();
            foreach (var file in logFiles)
            {
                try { System.IO.File.Delete(file); }
                catch { }
            }
        }
        catch { }
    }
}
