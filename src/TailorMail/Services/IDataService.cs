﻿﻿using TailorMail.Models;

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
    /// 保存收件人分组列表到持久化存储。
    /// </summary>
    /// <param name="groups">要保存的收件人分组列表。</param>
    void SaveRecipientGroups(List<RecipientGroup> groups);

    /// <summary>
    /// 加载应用程序设置。
    /// </summary>
    /// <returns>应用程序设置对象；若数据文件不存在则返回默认设置。</returns>
    AppSettings LoadSettings();

    /// <summary>
    /// 保存应用程序设置到持久化存储。
    /// </summary>
    /// <param name="settings">要保存的应用程序设置对象。</param>
    void SaveSettings(AppSettings settings);

    /// <summary>
    /// 加载附件配置。
    /// </summary>
    /// <returns>附件配置对象；若数据文件不存在则返回默认配置。</returns>
    AttachmentConfig LoadAttachmentConfig();

    /// <summary>
    /// 保存附件配置到持久化存储。
    /// </summary>
    /// <param name="config">要保存的附件配置对象。</param>
    void SaveAttachmentConfig(AttachmentConfig config);

    /// <summary>
    /// 加载邮件模板列表。
    /// </summary>
    /// <returns>邮件模板列表；若数据文件不存在则返回空列表。</returns>
    List<MailTemplate> LoadTemplates();

    /// <summary>
    /// 保存邮件模板列表到持久化存储。
    /// </summary>
    /// <param name="templates">要保存的邮件模板列表。</param>
    void SaveTemplates(List<MailTemplate> templates);
}
