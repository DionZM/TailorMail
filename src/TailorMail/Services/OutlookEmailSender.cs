using System.Collections.Generic;
using System.Runtime.InteropServices;
using TailorMail.Models;

namespace TailorMail.Services;

public class OutlookEmailSender : IEmailSender, IDisposable
{
    private static readonly Lazy<Type?> _outlookType = new(() =>
    {
        try { return Type.GetTypeFromProgID("Outlook.Application"); }
        catch { return null; }
    });

    // Bulk send: reuse single Outlook Application instance
    private dynamic? _bulkOutlookApp;
    private bool _isBulkMode;

    public string Name => "Outlook";

    /// <summary>
    /// Create and cache a single Outlook Application instance for bulk sending.
    /// </summary>
    public void PrepareForBulkSend()
    {
        // H-09: Release previous COM object if PrepareForBulkSend called twice
        CleanupBulkSend();
        var outlookType = _outlookType.Value
            ?? throw new InvalidOperationException("未检测到 Outlook，请确认已安装 Microsoft Outlook。");
        _bulkOutlookApp = Activator.CreateInstance(outlookType)!;
        _isBulkMode = true;
    }

    /// <summary>
    /// Release the cached Outlook Application instance after bulk sending.
    /// </summary>
    public void CleanupBulkSend()
    {
        if (_bulkOutlookApp != null)
        {
            // R-01: Use FinalReleaseComObject to ensure full release
            try { Marshal.FinalReleaseComObject(_bulkOutlookApp); }
            catch (Exception ex) { AppLogger.Warning($"Outlook COM释放警告: {ex.Message}"); }
            _bulkOutlookApp = null;
        }
        _isBulkMode = false;
    }

    public async Task<SendResult> SendAsync(
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

        try
        {
            await Task.Run(() =>
            {
                dynamic? mailItem = null;
                dynamic? localApp = null;
                try
                {
                    dynamic outlookApp;
                    if (_isBulkMode && _bulkOutlookApp != null)
                    {
                        outlookApp = _bulkOutlookApp!;
                    }
                    else
                    {
                        var outlookType = _outlookType.Value
                            ?? throw new InvalidOperationException("未检测到 Outlook，请确认已安装 Microsoft Outlook。");
                        outlookApp = Activator.CreateInstance(outlookType)!;
                        localApp = outlookApp;
                    }

                    mailItem = outlookApp.CreateItem(0); // 0 = olMailItem

                    mailItem.Subject = subject;
                    mailItem.HTMLBody = body;

                    var toList = recipient.GetToList();
                    var ccList = recipient.GetCcList();
                    var bccList = recipient.GetBccList();

                    if (toList.Count > 0) mailItem.To = string.Join(";", toList);
                    if (ccList.Count > 0) mailItem.CC = string.Join(";", ccList);
                    if (bccList.Count > 0) mailItem.BCC = string.Join(";", bccList);

                    foreach (var filePath in attachments)
                    {
                        if (System.IO.File.Exists(filePath))
                            mailItem.Attachments.Add(filePath, 1);
                    }

                    mailItem.Send();
                }
                finally
                {
                    // R-01: Use FinalReleaseComObject for thorough cleanup
                    if (mailItem != null) try { Marshal.FinalReleaseComObject(mailItem); } catch (Exception ex) { AppLogger.Warning($"Outlook mailItem释放警告: {ex.Message}"); }
                    if (localApp != null) try { Marshal.FinalReleaseComObject(localApp); } catch (Exception ex) { AppLogger.Warning($"Outlook localApp释放警告: {ex.Message}"); }
                }
            });

            result.Status = SendStatus.Success;
            result.SendTime = DateTime.Now;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Outlook发送失败 - {recipient.Name}", ex);
            result.Status = SendStatus.Failed;
            result.ErrorMessage = ex.Message;
        }

        return result;
    }

    /// <summary>
    /// H-10: SendTest now includes attachments.
    /// </summary>
    public bool SendTest(string fromAddress, string toAddress, string subject, string bodyHtml, string[] attachments, out string error)
    {
        try
        {
            var outlookType = _outlookType.Value
                ?? throw new InvalidOperationException("未检测到 Outlook，请确认已安装 Microsoft Outlook。");

            dynamic? outlookApp = null;
            dynamic? mailItem = null;
            try
            {
                outlookApp = Activator.CreateInstance(outlookType)!;
                mailItem = outlookApp.CreateItem(0);
                mailItem.Subject = subject;
                mailItem.HTMLBody = bodyHtml;
                mailItem.To = toAddress;

                foreach (var filePath in attachments)
                {
                    if (!string.IsNullOrEmpty(filePath) && System.IO.File.Exists(filePath))
                        mailItem.Attachments.Add(filePath, 1);
                }

                mailItem.Send();
            }
            finally
            {
                if (mailItem != null) try { Marshal.FinalReleaseComObject(mailItem); } catch (Exception ex) { AppLogger.Warning($"Outlook mailItem释放警告: {ex.Message}"); }
                if (outlookApp != null) try { Marshal.FinalReleaseComObject(outlookApp); } catch (Exception ex) { AppLogger.Warning($"Outlook app释放警告: {ex.Message}"); }
            }

            error = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    public void Dispose()
    {
        CleanupBulkSend();
    }
}
