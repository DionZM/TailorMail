using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 数据服务接口，定义应用程序数据的加载与保存方法。
/// 实现类负责将收件人分组、应用设置、附件配置和邮件模板等数据持久化到存储介质中。
/// </summary>
public interface IDataService
{
    /// <summary>
    /// 加载收件人分组列表。
    /// </summary>
    /// <returns>收件人分组列表；若数据文件不存在则返回包含默认分组的列表。</returns>
    List<RecipientGroup> LoadRecipientGroups();

    /// <summary>
    /// 异步加载收件人分组列表（P-01: 避免阻塞 UI 线程）。
    /// </summary>
    Task<List<RecipientGroup>> LoadRecipientGroupsAsync();

    /// <summary>
    /// 保存收件人分组列表到持久化存储。
    /// </summary>
    /// <param name="groups">要保存的收件人分组列表。</param>
    void SaveRecipientGroups(List<RecipientGroup> groups);

    void SaveRecipientVariables(IEnumerable<Recipient> recipients);

    /// <summary>
    /// 加载所有自定义变量名称列表。
    /// </summary>
    List<string> LoadVariableNames();

    /// <summary>
    /// 添加一个变量名称。
    /// </summary>
    void AddVariableName(string name);

    /// <summary>
    /// 删除一个变量名称，并从所有收件人的变量 JSON 中移除该键。
    /// </summary>
    void DeleteVariableName(string name);

    /// <summary>
    /// 重命名变量，并批量更新所有收件人的变量 JSON 中的键名。
    /// </summary>
    void RenameVariableName(string oldName, string newName);

    /// <summary>
    /// 加载应用程序设置。
    /// </summary>
    AppSettings LoadSettings();

    /// <summary>
    /// 异步加载应用程序设置（P-01: 避免阻塞 UI 线程）。
    /// </summary>
    Task<AppSettings> LoadSettingsAsync();

    /// <summary>
    /// 保存应用程序设置到持久化存储。
    /// </summary>
    void SaveSettings(AppSettings settings);

    /// <summary>
    /// 加载附件配置。
    /// </summary>
    AttachmentConfig LoadAttachmentConfig();

    /// <summary>
    /// 保存附件配置到持久化存储。
    /// </summary>
    void SaveAttachmentConfig(AttachmentConfig config);

    /// <summary>
    /// 加载邮件模板列表。
    /// </summary>
    List<MailTemplate> LoadTemplates();

    /// <summary>
    /// 保存邮件模板列表到持久化存储。
    /// </summary>
    void SaveTemplates(List<MailTemplate> templates);
}
