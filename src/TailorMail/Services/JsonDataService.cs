﻿﻿using System.IO;
using System.Text.Json;
using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 基于 JSON 文件的数据服务实现，实现 <see cref="IDataService"/> 接口。
/// 将所有应用数据（收件人、设置、附件配置、模板）以 JSON 格式存储在应用程序目录下的 data 文件夹中。
/// 使用 <see cref="JavaScriptEncoder.UnsafeRelaxedJsonEscaping"/> 确保中文字符不被转义为 Unicode 序列。
/// </summary>
public class JsonDataService : IDataService
{
    /// <summary>
    /// JSON 序列化选项，启用缩进格式化和中文安全编码。
    /// </summary>
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    /// <summary>
    /// 数据文件存储目录的绝对路径。
    /// </summary>
    private readonly string _dataDir;

    /// <summary>
    /// 初始化 <see cref="JsonDataService"/> 实例。
    /// 在应用程序根目录下创建 data 和 data/templates 子目录。
    /// </summary>
    public JsonDataService()
    {
        _dataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        Directory.CreateDirectory(_dataDir);
        Directory.CreateDirectory(Path.Combine(_dataDir, "templates"));
    }

    /// <inheritdoc/>
    /// <remarks>
    /// 从 data/recipients.json 文件加载。若文件不存在或反序列化失败，返回包含一个"默认分组"的列表。
    /// </remarks>
    public List<RecipientGroup> LoadRecipientGroups()
    {
        var path = Path.Combine(_dataDir, "recipients.json");
        if (!File.Exists(path)) return [new RecipientGroup { Name = "默认分组" }];
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<List<RecipientGroup>>(json, _jsonOptions) ?? [new RecipientGroup { Name = "默认分组" }];
    }

    /// <inheritdoc/>
    public void SaveRecipientGroups(List<RecipientGroup> groups)
    {
        var path = Path.Combine(_dataDir, "recipients.json");
        var json = JsonSerializer.Serialize(groups, _jsonOptions);
        File.WriteAllText(path, json);
    }

    /// <inheritdoc/>
    /// <remarks>
    /// 从 data/settings.json 文件加载。若文件不存在或反序列化失败，返回默认的 <see cref="AppSettings"/> 实例。
    /// </remarks>
    public AppSettings LoadSettings()
    {
        var path = Path.Combine(_dataDir, "settings.json");
        if (!File.Exists(path)) return new AppSettings();
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
    }

    /// <inheritdoc/>
    public void SaveSettings(AppSettings settings)
    {
        var path = Path.Combine(_dataDir, "settings.json");
        var json = JsonSerializer.Serialize(settings, _jsonOptions);
        File.WriteAllText(path, json);
    }

    /// <inheritdoc/>
    /// <remarks>
    /// 从 data/attachments.json 文件加载。若文件不存在或反序列化失败，返回默认的 <see cref="AttachmentConfig"/> 实例。
    /// </remarks>
    public AttachmentConfig LoadAttachmentConfig()
    {
        var path = Path.Combine(_dataDir, "attachments.json");
        if (!File.Exists(path)) return new AttachmentConfig();
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<AttachmentConfig>(json, _jsonOptions) ?? new AttachmentConfig();
    }

    /// <inheritdoc/>
    public void SaveAttachmentConfig(AttachmentConfig config)
    {
        var path = Path.Combine(_dataDir, "attachments.json");
        var json = JsonSerializer.Serialize(config, _jsonOptions);
        File.WriteAllText(path, json);
    }

    /// <inheritdoc/>
    /// <remarks>
    /// 从 data/templates/templates.json 文件加载。若文件不存在或反序列化失败，返回空列表。
    /// </remarks>
    public List<MailTemplate> LoadTemplates()
    {
        var path = Path.Combine(_dataDir, "templates", "templates.json");
        if (!File.Exists(path)) return [];
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<List<MailTemplate>>(json, _jsonOptions) ?? [];
    }

    /// <inheritdoc/>
    public void SaveTemplates(List<MailTemplate> templates)
    {
        var dir = Path.Combine(_dataDir, "templates");
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "templates.json");
        var json = JsonSerializer.Serialize(templates, _jsonOptions);
        File.WriteAllText(path, json);
    }
}
