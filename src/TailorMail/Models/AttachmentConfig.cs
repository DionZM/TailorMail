namespace TailorMail.Models;

/// <summary>
/// 附件配置模型，管理邮件发送时的附件信息。
/// 支持两种附件模式：公共附件（所有收件人共享）和单位专属附件（每个收件人单独配置），
/// 同时支持从指定目录自动匹配附件文件。
/// </summary>
public class AttachmentConfig
{
    /// <summary>
    /// 获取或设置公共附件文件路径列表。
    /// 列表中的附件将附加到每一封发送的邮件中，所有收件人共享。默认为空列表。
    /// </summary>
    public List<string> CommonAttachments { get; set; } = [];

    /// <summary>
    /// 获取或设置单位专属附件配置列表。
    /// 每个元素对应一个收件人的专属附件配置，可实现不同收件人附加不同的文件。默认为空列表。
    /// </summary>
    public List<UnitAttachment> UnitAttachments { get; set; } = [];

    /// <summary>
    /// 获取或设置附件自动匹配目录路径。
    /// 设置后，系统将根据收件人名称在该目录下自动查找匹配的文件作为附件。
    /// 若为 <c>null</c>，则不启用自动匹配功能。
    /// </summary>
    public string? AutoMatchDirectory { get; set; }
}

/// <summary>
/// 单个收件人的附件配置模型，表示某个收件人专属的附件信息。
/// 包含收件人标识、关联的附件文件列表以及自动匹配状态。
/// </summary>
public class UnitAttachment
{
    /// <summary>
    /// 获取或设置关联收件人的唯一标识符。用于与 <see cref="Recipient.Id"/> 建立对应关系。
    /// </summary>
    public string RecipientId { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置关联收件人的名称。用于界面显示及附件自动匹配时的关键字匹配。
    /// </summary>
    public string RecipientName { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置该收件人专属的附件文件路径列表。默认为空列表。
    /// </summary>
    public List<string> Files { get; set; } = [];

    /// <summary>
    /// 获取或设置当前附件是否为自动匹配的结果。
    /// <c>true</c> 表示附件是通过 <see cref="AttachmentConfig.AutoMatchDirectory"/> 自动匹配得到的；
    /// <c>false</c> 表示附件为用户手动添加。
    /// </summary>
    public bool IsAutoMatched { get; set; }

    /// <summary>
    /// 获取一个值，指示该收件人是否存在附件文件。
    /// 当 <see cref="Files"/> 列表中包含至少一个文件路径时返回 <c>true</c>；否则返回 <c>false</c>。
    /// </summary>
    public bool HasFiles => Files.Count > 0;
}
