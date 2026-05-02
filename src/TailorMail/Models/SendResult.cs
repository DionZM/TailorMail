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
