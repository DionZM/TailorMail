using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;

namespace TailorMail.ViewModels;

/// <summary>
/// 附件管理视图模型，负责公共附件和收件人专属附件的管理与自动匹配。
/// 支持从指定目录按收件人名称自动匹配附件文件，也可手动添加。
/// </summary>
public partial class AttachmentViewModel : ObservableObject
{
    private readonly IDataService _dataService;
    private readonly AttachmentMatchService _matchService;

    /// <summary>
    /// 获取或设置公共附件文件路径列表（所有收件人共享）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<string> _commonAttachments = [];

    /// <summary>
    /// 获取或设置收件人专属附件配置列表（每个收件人可独立配置）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<RecipientAttachment> _RecipientAttachments = [];

    /// <summary>
    /// 获取或设置当前选中的收件人附件配置项。
    /// </summary>
    [ObservableProperty]
    private RecipientAttachment? _selectedRecipientAttachment;

    /// <summary>
    /// 获取或设置自动匹配的预览文本，显示匹配结果摘要。
    /// </summary>
    [ObservableProperty]
    private string _matchPreviewText = string.Empty;

    /// <summary>
    /// 获取或设置已匹配到专属附件的收件人数量。
    /// </summary>
    [ObservableProperty]
    private int _matchedCount;

    /// <summary>
    /// 获取或设置未匹配到专属附件的收件人数量。
    /// </summary>
    [ObservableProperty]
    private int _unmatchedCount;

    /// <summary>
    /// 获取或设置附件自动匹配的目录路径。
    /// </summary>
    [ObservableProperty]
    private string _matchDirectory = string.Empty;

    public AttachmentViewModel(IDataService dataService, AttachmentMatchService matchService)
    {
        _dataService = dataService;
        _matchService = matchService;
        LoadData();
    }

    /// <summary>
    /// 从数据服务加载附件配置，并确保所有选中的收件人都在列表中。
    /// </summary>
    public void LoadData()
    {
        var config = _dataService.LoadAttachmentConfig();
        CommonAttachments = new ObservableCollection<string>(config.CommonAttachments);
        RecipientAttachments = new ObservableCollection<RecipientAttachment>(config.RecipientAttachments);
        if (config.AutoMatchDirectory != null)
            MatchDirectory = config.AutoMatchDirectory;

        EnsureAllRecipientsInList();
    }

    /// <summary>
    /// 确保所有选中的收件人在收件人附件列表中有对应条目，
    /// 同时移除未选中收件人的条目，保持列表与选中状态同步。
    /// </summary>
    private void EnsureAllRecipientsInList()
    {
        var groups = _dataService.LoadRecipientGroups();
        var selectedRecipients = groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected).ToList();
        var existingIds = RecipientAttachments.Select(ua => ua.RecipientId).ToHashSet();

        // 添加缺失的选中收件人
        foreach (var r in selectedRecipients)
        {
            if (!existingIds.Contains(r.Id))
            {
                RecipientAttachments.Add(new RecipientAttachment
                {
                    RecipientId = r.Id,
                    RecipientName = r.Name,
                    Files = [],
                    IsAutoMatched = false
                });
            }
        }

        // 移除未选中的收件人
        var selectedIds = selectedRecipients.Select(r => r.Id).ToHashSet();
        var toRemove = RecipientAttachments.Where(ua => !selectedIds.Contains(ua.RecipientId)).ToList();
        foreach (var ua in toRemove)
            RecipientAttachments.Remove(ua);
    }

    /// <summary>
    /// 添加公共附件文件。添加后自动触发从公共附件目录的匹配。
    /// </summary>
    [RelayCommand]
    private void AddCommonAttachment()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Multiselect = true,
            Title = "选择公共附件"
        };
        if (dialog.ShowDialog() != true) return;

        foreach (var file in dialog.FileNames)
        {
            if (!CommonAttachments.Contains(file))
                CommonAttachments.Add(file);
        }

        AutoMatchFromCommonAttachments();
        SaveConfig();
    }

    /// <summary>
    /// 移除指定的公共附件。
    /// </summary>
    /// <param name="filePath">要移除的文件路径。</param>
    [RelayCommand]
    private void RemoveCommonAttachment(string filePath)
    {
        CommonAttachments.Remove(filePath);
        SaveConfig();
    }

    /// <summary>
    /// 移除指定收件人的指定专属附件。
    /// </summary>
    /// <param name="filePath">要移除的文件路径。</param>
    [RelayCommand]
    private void RemoveRecipientAttachment(string filePath)
    {
        foreach (var ua in RecipientAttachments)
        {
            if (ua.Files.Contains(filePath))
            {
                ua.Files.Remove(filePath);
                break;
            }
        }
        SaveConfig();
    }

    /// <summary>
    /// 执行自动匹配。优先使用指定的匹配目录，否则从公共附件所在目录推断。
    /// </summary>
    [RelayCommand]
    private void AutoMatch()
    {
        if (!string.IsNullOrWhiteSpace(MatchDirectory) && System.IO.Directory.Exists(MatchDirectory))
        {
            AutoMatchFromDirectory(MatchDirectory);
            return;
        }
        AutoMatchFromCommonAttachments();
    }

    /// <summary>
    /// 从公共附件所在目录推断匹配目录（选择出现频率最高的目录），
    /// 然后在该目录下执行自动匹配。
    /// </summary>
    private void AutoMatchFromCommonAttachments()
    {
        if (CommonAttachments.Count == 0) return;

        // 推断公共附件最常出现的目录作为匹配目录
        var dir = CommonAttachments.Select(System.IO.Path.GetDirectoryName)
            .Where(d => d != null)
            .GroupBy(d => d)
            .OrderByDescending(g => g.Count())
            .FirstOrDefault()?.Key;

        if (dir != null && System.IO.Directory.Exists(dir))
            AutoMatchFromDirectory(dir);
    }

    /// <summary>
    /// 在指定目录下根据收件人名称自动匹配附件文件。
    /// 匹配结果会更新到收件人附件列表，并生成预览文本。
    /// </summary>
    /// <param name="dir">待搜索的目录路径。</param>
    private void AutoMatchFromDirectory(string dir)
    {
        var groups = _dataService.LoadRecipientGroups();
        var selectedRecipients = groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected).ToList();
        var matched = _matchService.MatchFilesByRecipient(dir, selectedRecipients, CommonAttachments);

        // 生成匹配预览文本
        var preview = new System.Text.StringBuilder();
        preview.AppendLine($"匹配目录: {dir}");
        preview.AppendLine();

        var matchedIds = new HashSet<string>();
        foreach (var kvp in matched)
        {
            var recipient = selectedRecipients.FirstOrDefault(r => r.Id == kvp.Key);
            if (recipient == null) continue;
            matchedIds.Add(recipient.Id);

            preview.AppendLine($"✓ {recipient.Name}");
            foreach (var file in kvp.Value)
                preview.AppendLine($"    → {System.IO.Path.GetFileName(file)}");

            // 将匹配结果合并到收件人附件列表
            var existing = RecipientAttachments.FirstOrDefault(ua => ua.RecipientId == kvp.Key);
            if (existing != null)
            {
                foreach (var file in kvp.Value)
                {
                    if (!existing.Files.Contains(file))
                        existing.Files.Add(file);
                }
            }
            else
            {
                RecipientAttachments.Add(new RecipientAttachment
                {
                    RecipientId = kvp.Key,
                    RecipientName = recipient.Name,
                    Files = kvp.Value,
                    IsAutoMatched = true
                });
            }
        }

        // 记录未匹配的收件人
        var unmatched = selectedRecipients.Where(r => !matchedIds.Contains(r.Id)).ToList();
        foreach (var r in unmatched)
        {
            var existing = RecipientAttachments.FirstOrDefault(ua => ua.RecipientId == r.Id);
            if (existing == null)
            {
                RecipientAttachments.Add(new RecipientAttachment
                {
                    RecipientId = r.Id,
                    RecipientName = r.Name,
                    Files = [],
                    IsAutoMatched = false
                });
            }
            preview.AppendLine($"✗ {r.Name}");
        }

        if (unmatched.Count > 0)
        {
            preview.AppendLine();
            preview.AppendLine($"共 {unmatched.Count} 位收件人未匹配到专有附件，可在列表中手动添加");
        }

        MatchPreviewText = preview.ToString();
        MatchedCount = matchedIds.Count;
        UnmatchedCount = unmatched.Count;
        SaveConfig();
    }

    /// <summary>
    /// 保存当前附件配置到数据服务。仅保存有文件的收件人附件条目。
    /// </summary>
    public void SaveConfig()
    {
        var config = new AttachmentConfig
        {
            CommonAttachments = CommonAttachments.ToList(),
            RecipientAttachments = RecipientAttachments.Where(ua => ua.Files.Count > 0).ToList(),
            AutoMatchDirectory = MatchDirectory
        };
        _dataService.SaveAttachmentConfig(config);
    }
}
