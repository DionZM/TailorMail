namespace TailorMail.Models;

/// <summary>
/// 邮件发送结果模型，记录单个收件人的邮件发送状态和结果信息。
/// 包含收件人标识、发送状态、错误信息及发送时间，用于在界面上展示发送进度和结果。
/// </summary>
public class SendResult
{
    /// <summary>
    /// 获取或设置关联收件人的唯一标识符。用于与 <see cref="Recipient.Id"/> 建立对应关系。
    /// </summary>
    public string RecipientId { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置关联收件人的名称。用于在发送结果列表中显示收件人信息。
    /// </summary>
    public string RecipientName { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置邮件发送状态。默认值为 <see cref="SendStatus.Pending"/>（等待中）。
    /// </summary>
    public SendStatus Status { get; set; } = SendStatus.Pending;

    /// <summary>
    /// 获取或设置发送失败时的错误信息。
    /// 当 <see cref="Status"/> 为 <see cref="SendStatus.Failed"/> 时，此属性包含具体的错误描述；发送成功时为空字符串。
    /// </summary>
    public string ErrorMessage { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置邮件发送完成的时间。仅在发送成功或失败后赋值；若尚未发送则为 <c>null</c>。
    /// </summary>
    public DateTime? SendTime { get; set; }

    /// <summary>
    /// 获取发送状态的中文显示文本，用于界面展示。
    /// 根据 <see cref="Status"/> 的值返回对应的中文描述：
    /// Pending → "等待中"，Sending → "发送中..."，Success → "已发送"，Failed → "发送失败"。
    /// </summary>
    public string StatusText => Status switch
    {
        SendStatus.Pending => "等待中",
        SendStatus.Sending => "发送中...",
        SendStatus.Success => "已发送",
        SendStatus.Failed => "发送失败",
        _ => ""
    };
}

/// <summary>
/// 邮件发送状态枚举，定义邮件发送过程中的各个阶段。
/// </summary>
public enum SendStatus
{
    /// <summary>等待中：邮件尚未开始发送，正在排队等待。</summary>
    Pending,

    /// <summary>发送中：邮件正在发送过程中。</summary>
    Sending,

    /// <summary>已发送：邮件发送成功。</summary>
    Success,

    /// <summary>发送失败：邮件发送过程中出现错误。错误详情参见 <see cref="SendResult.ErrorMessage"/>。</summary>
    Failed
}
