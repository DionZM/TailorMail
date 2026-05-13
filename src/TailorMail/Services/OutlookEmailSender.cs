using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using TailorMail.Models;

namespace TailorMail.Services;

public class OutlookEmailSender : IEmailSender, IDisposable
{
    private static readonly Lazy<Type?> _outlookType = new(() =>
    {
        try { return Type.GetTypeFromProgID("Outlook.Application"); }
        catch { return null; }
    });

    private const BindingFlags InvokeMethod = BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public;
    private const BindingFlags SetProperty = BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public;
    private const BindingFlags GetProperty = BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public;

    public string Name => "Outlook";

    public void PrepareForBulkSend() { }
    public void CleanupBulkSend() { }

    public Task<SendResult> SendAsync(
        string subject,
        string body,
        Recipient recipient,
        List<string> attachments,
        string? smtpPassword = null,
        Models.SmtpSettings? smtpSettings = null)
    {
        var result = new SendResult
        {
            RecipientId = recipient.Id,
            RecipientName = recipient.Name,
            Status = SendStatus.Sending
        };

        var capturedSubject = subject;
        var capturedBody = body;
        var capturedAttachments = attachments.ToArray();
        var capturedToList = recipient.GetToList();
        var capturedCcList = recipient.GetCcList();
        var capturedBccList = recipient.GetBccList();
        var capturedName = recipient.Name;

        Exception? capturedEx = null;

        var staThread = new Thread(() =>
        {
            object? outlookApp = null;
            object? mailItem = null;
            try
            {
                var outlookType = _outlookType.Value
                    ?? throw new InvalidOperationException("未检测到 Outlook，请确认已安装 Microsoft Outlook。");

                FlushSync($"Outlook: 创建实例 for {capturedName}");

                outlookApp = Activator.CreateInstance(outlookType)!;

                FlushSync($"Outlook: CreateItem for {capturedName}");

                mailItem = outlookType.InvokeMember("CreateItem", InvokeMethod, null, outlookApp, new object[] { 0 })!;

                FlushSync($"Outlook: 设置属性 for {capturedName}");

                mailItem.GetType().InvokeMember("Subject", SetProperty, null, mailItem, new object[] { capturedSubject });
                mailItem.GetType().InvokeMember("HTMLBody", SetProperty, null, mailItem, new object[] { capturedBody });

                if (capturedToList.Count > 0)
                    mailItem.GetType().InvokeMember("To", SetProperty, null, mailItem, new object[] { string.Join(";", capturedToList) });
                if (capturedCcList.Count > 0)
                    mailItem.GetType().InvokeMember("CC", SetProperty, null, mailItem, new object[] { string.Join(";", capturedCcList) });
                if (capturedBccList.Count > 0)
                    mailItem.GetType().InvokeMember("BCC", SetProperty, null, mailItem, new object[] { string.Join(";", capturedBccList) });

                FlushSync($"Outlook: 添加附件 for {capturedName}, 数量={capturedAttachments.Length}");

                foreach (var filePath in capturedAttachments)
                {
                    if (System.IO.File.Exists(filePath))
                    {
                        var attachmentsProp = mailItem.GetType().InvokeMember("Attachments", GetProperty, null, mailItem, null)!;
                        attachmentsProp.GetType().InvokeMember("Add", InvokeMethod, null, attachmentsProp, new object[] { filePath, 1 });
                    }
                }

                FlushSync($"Outlook: Send for {capturedName}");

                mailItem.GetType().InvokeMember("Send", InvokeMethod, null, mailItem, null);

                FlushSync($"Outlook: Send 成功 for {capturedName}");

                result.Status = SendStatus.Success;
                result.SendTime = DateTime.Now;
            }
            catch (Exception ex)
            {
                var msg = ex.InnerException?.Message ?? ex.Message;
                FlushSync($"Outlook: 异常 for {capturedName}: {msg}");
                capturedEx = new Exception(msg, ex);
            }
            finally
            {
                if (mailItem != null) try { Marshal.FinalReleaseComObject(mailItem); } catch { }
                if (outlookApp != null) try { Marshal.FinalReleaseComObject(outlookApp); } catch { }
            }
        });

        staThread.SetApartmentState(ApartmentState.STA);
        staThread.IsBackground = false;
        staThread.Start();
        staThread.Join();

        if (capturedEx != null)
        {
            AppLogger.Error($"Outlook发送失败 - {capturedName}", capturedEx);
            result.Status = SendStatus.Failed;
            result.ErrorMessage = capturedEx.Message;
        }

        return Task.FromResult(result);
    }

    private static void FlushSync(string message)
    {
        try
        {
            var line = $"[{DateTime.Now:HH:mm:ss.fff}] [INFO] {message}{Environment.NewLine}";
            var logDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            var logFile = System.IO.Path.Combine(logDir, $"log_{DateTime.Now:yyyyMMdd}.txt");
            System.IO.File.AppendAllText(logFile, line);
        }
        catch { }
    }

    public void Dispose() { }
}
