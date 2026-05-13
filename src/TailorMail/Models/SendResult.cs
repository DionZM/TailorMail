using CommunityToolkit.Mvvm.ComponentModel;

namespace TailorMail.Models;

public partial class SendResult : ObservableObject
{
    [ObservableProperty]
    private string _recipientId = string.Empty;

    [ObservableProperty]
    private string _recipientName = string.Empty;

    [ObservableProperty]
    private SendStatus _status = SendStatus.Pending;

    [ObservableProperty]
    private string _errorMessage = string.Empty;

    [ObservableProperty]
    private DateTime? _sendTime;

    public string StatusText => Status switch
    {
        SendStatus.Pending => "等待中",
        SendStatus.Sending => "发送中...",
        SendStatus.Success => "已发送",
        SendStatus.Failed => "发送失败",
        _ => ""
    };

    private string GetFriendlyError()
    {
        if (string.IsNullOrEmpty(ErrorMessage)) return "发送失败";
        var msg = ErrorMessage;

        if (msg.Contains("AuthenticationFailed") || msg.Contains("535"))
            return "认证失败，请检查用户名和密码";
        if (msg.Contains("timeout") || msg.Contains("Timeout") || msg.Contains("timed out"))
            return "连接超时，请检查网络";
        if (msg.Contains("refused") || msg.Contains("Refused"))
            return "被拒绝，请检查端口/SSL设置";
        if (msg.Contains("not resolve") || msg.Contains("No such host") || msg.Contains("HostNotFound"))
            return "找不到服务器，请检查地址";
        if (msg.Contains("rate") || msg.Contains("limit") || msg.Contains("Too many"))
            return "发送受限，请增加发送间隔";
        if (msg.Contains("quota") || msg.Contains("Quota"))
            return "超出每日发送上限";
        if (msg.Contains("spam") || msg.Contains("Spam"))
            return "被判为垃圾邮件";

        return msg.Length > 100 ? msg[..97] + "..." : msg;
    }

    partial void OnStatusChanged(SendStatus value)
    {
        OnPropertyChanged(nameof(StatusText));
    }
}

public enum SendStatus
{
    Pending,
    Sending,
    Success,
    Failed
}
