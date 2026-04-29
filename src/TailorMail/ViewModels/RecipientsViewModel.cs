﻿﻿using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;
using OfficeOpenXml;

namespace TailorMail.ViewModels;

/// <summary>
/// 收件人管理视图模型，负责收件人分组的增删改查、收件人的导入导出、
/// 选择状态维护以及与数据服务的交互。
/// 支持从 Excel 文件和其他分组导入收件人，采用"按名称匹配、仅更新非空字段"的合并策略。
/// </summary>
public partial class RecipientsViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    /// <summary>
    /// 记录切换分组前是否全选，用于在切换分组后恢复选择状态。
    /// </summary>
    private bool _wasAllSelected;

    /// <summary>
    /// 获取或设置收件人分组列表。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<RecipientGroup> _groups = [];

    /// <summary>
    /// 获取或设置当前选中的分组。切换分组时会自动保存旧分组数据并加载新分组的收件人。
    /// </summary>
    [ObservableProperty]
    private RecipientGroup? _selectedGroup;

    /// <summary>
    /// 获取或设置当前分组下的收件人列表（用于界面绑定）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<Recipient> _currentRecipients = [];

    /// <summary>
    /// 获取或设置新分组的名称输入。
    /// </summary>
    [ObservableProperty]
    private string _newGroupName = string.Empty;

    /// <summary>
    /// 获取或设置所有分组中被选中的收件人总数。
    /// </summary>
    [ObservableProperty]
    private int _selectedCount;

    /// <summary>
    /// 获取或设置所有分组中的收件人总数。
    /// </summary>
    [ObservableProperty]
    private int _totalCount;

    /// <summary>
    /// 当选中的分组发生变化时触发，用于通知界面刷新。
    /// </summary>
    public event Action? SelectedGroupChanged;

    public RecipientsViewModel(IDataService dataService)
    {
        _dataService = dataService;
        LoadGroups();
    }

    /// <summary>
    /// 从数据服务加载所有分组，并默认选中第一个分组。
    /// </summary>
    public void LoadGroups()
    {
        var groups = _dataService.LoadRecipientGroups();
        Groups = new ObservableCollection<RecipientGroup>(groups);
        if (Groups.Count > 0)
            SelectedGroup = Groups[0];
    }

    /// <summary>
    /// 分组切换前的处理：记录当前分组是否全选，并保存当前数据。
    /// </summary>
    partial void OnSelectedGroupChanging(RecipientGroup? value)
    {
        if (SelectedGroup != null)
        {
            _wasAllSelected = CurrentRecipients.Count > 0 && CurrentRecipients.All(r => r.IsSelected);
        }
        SaveAll();
    }

    /// <summary>
    /// 分组切换后的处理：加载新分组的收件人，恢复选择状态，更新计数。
    /// </summary>
    partial void OnSelectedGroupChanged(RecipientGroup? value)
    {
        if (value != null)
        {
            CurrentRecipients = new ObservableCollection<Recipient>(value.Recipients);
            // 根据之前分组的全选状态决定新分组的选择状态
            if (_wasAllSelected)
            {
                foreach (var r in CurrentRecipients) r.IsSelected = true;
            }
            else
            {
                foreach (var r in CurrentRecipients) r.IsSelected = false;
            }
        }
        else
        {
            CurrentRecipients = [];
        }
        UpdateCounts();
        SelectedGroupChanged?.Invoke();
    }

    /// <summary>
    /// 添加新分组。若分组名已存在则提示用户。
    /// </summary>
    [RelayCommand]
    private void AddGroup()
    {
        if (string.IsNullOrWhiteSpace(NewGroupName)) return;
        if (Groups.Any(g => g.Name == NewGroupName))
        {
            System.Windows.MessageBox.Show("分组名已存在", "提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }
        Groups.Add(new RecipientGroup { Name = NewGroupName });
        NewGroupName = string.Empty;
        SaveAll();
    }

    /// <summary>
    /// 删除当前选中的分组（至少保留一个分组）。删除前弹出确认对话框。
    /// </summary>
    [RelayCommand]
    private void DeleteGroup()
    {
        if (SelectedGroup == null) return;
        if (Groups.Count <= 1) return;
        if (System.Windows.MessageBox.Show($"确定删除分组「{SelectedGroup.Name}」及其所有发送对象？", "确认删除",
            System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question) != System.Windows.MessageBoxResult.Yes) return;
        Groups.Remove(SelectedGroup);
        SelectedGroup = Groups.FirstOrDefault();
        SaveAll();
    }

    /// <summary>
    /// 删除指定的收件人。
    /// </summary>
    /// <param name="r">要删除的收件人对象。</param>
    [RelayCommand]
    private void DeleteRecipient(Recipient r)
    {
        if (SelectedGroup == null) return;
        SelectedGroup.Recipients.Remove(r);
        CurrentRecipients.Remove(r);
        SaveAll();
        UpdateCounts();
    }

    /// <summary>
    /// 将指定收件人在列表中上移一位。
    /// </summary>
    /// <param name="r">要移动的收件人对象。</param>
    public void MoveUp(Recipient r)
    {
        if (SelectedGroup == null) return;
        var idx = CurrentRecipients.IndexOf(r);
        if (idx <= 0) return;
        CurrentRecipients.Move(idx, idx - 1);
        SyncRecipientsToGroup();
        SaveAll();
    }

    /// <summary>
    /// 将指定收件人在列表中下移一位。
    /// </summary>
    /// <param name="r">要移动的收件人对象。</param>
    public void MoveDown(Recipient r)
    {
        if (SelectedGroup == null) return;
        var idx = CurrentRecipients.IndexOf(r);
        if (idx < 0 || idx >= CurrentRecipients.Count - 1) return;
        CurrentRecipients.Move(idx, idx + 1);
        SyncRecipientsToGroup();
        SaveAll();
    }

    /// <summary>
    /// 将界面上的收件人列表同步回分组对象，过滤掉所有字段均为空的行。
    /// </summary>
    private void SyncRecipientsToGroup()
    {
        if (SelectedGroup == null) return;
        SelectedGroup.Recipients = CurrentRecipients
            .Where(r => !string.IsNullOrEmpty(r.Name) ||
                        !string.IsNullOrEmpty(r.ShortName) ||
                        !string.IsNullOrEmpty(r.ToEmails) ||
                        !string.IsNullOrEmpty(r.CcEmails) ||
                        !string.IsNullOrEmpty(r.BccEmails) ||
                        !string.IsNullOrEmpty(r.Remark))
            .ToList();
    }

    /// <summary>
    /// 移除名称为空的收件人行。
    /// </summary>
    public void RemoveEmptyRows()
    {
        var emptyRows = CurrentRecipients.Where(r => string.IsNullOrEmpty(r.Name)).ToList();
        foreach (var r in emptyRows)
        {
            CurrentRecipients.Remove(r);
            if (SelectedGroup != null)
                SelectedGroup.Recipients.Remove(r);
        }
        SaveAll();
    }

    /// <summary>
    /// 合并收件人列表到当前分组。采用"按名称匹配"策略：
    /// 若名称已存在则仅更新非空字段，否则新增收件人。
    /// </summary>
    /// <param name="incoming">待合并的收件人列表。</param>
    private void MergeRecipients(List<Recipient> incoming)
    {
        if (SelectedGroup == null) return;
        int updated = 0, added = 0;
        foreach (var src in incoming)
        {
            var existing = CurrentRecipients.FirstOrDefault(r => r.Name == src.Name);
            if (existing != null)
            {
                // 仅更新非空字段，避免覆盖已有数据
                if (!string.IsNullOrEmpty(src.ShortName)) existing.ShortName = src.ShortName;
                if (!string.IsNullOrEmpty(src.ToEmails)) existing.ToEmails = src.ToEmails;
                if (!string.IsNullOrEmpty(src.CcEmails)) existing.CcEmails = src.CcEmails;
                if (!string.IsNullOrEmpty(src.BccEmails)) existing.BccEmails = src.BccEmails;
                if (!string.IsNullOrEmpty(src.Remark)) existing.Remark = src.Remark;
                updated++;
            }
            else
            {
                var clone = new Recipient
                {
                    Name = src.Name, ShortName = src.ShortName,
                    ToEmails = src.ToEmails, CcEmails = src.CcEmails,
                    BccEmails = src.BccEmails, Remark = src.Remark,
                    Variables = new Dictionary<string, string>(src.Variables)
                };
                SelectedGroup.Recipients.Add(clone);
                CurrentRecipients.Add(clone);
                added++;
            }
        }
        SaveAll();
        UpdateCounts();
        System.Windows.MessageBox.Show($"导入完成：新增 {added} 项，更新 {updated} 项", "导入完成");
    }

    /// <summary>
    /// 从 Excel 文件导入收件人。Excel 格式：第1列名称，第2列简称，第3列收件人邮箱，第4列抄送邮箱，第5列密送邮箱，第6列备注。
    /// </summary>
    [RelayCommand]
    private void ImportFromExcel()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel文件|*.xlsx;*.xls|CSV文件|*.csv|所有文件|*.*",
            Title = "导入发送对象"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            var imported = ImportFromFile(dialog.FileName);
            if (SelectedGroup == null) return;
            MergeRecipients(imported);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"导入失败: {ex.Message}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 将当前分组的收件人导出为 Excel 文件。
    /// </summary>
    [RelayCommand]
    private void ExportToExcel()
    {
        if (SelectedGroup == null || CurrentRecipients.Count == 0) return;
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel文件|*.xlsx",
            FileName = $"{SelectedGroup.Name}_发送对象.xlsx"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("TailorMail");
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("发送对象");
            ws.Cells[1, 1].Value = "名称";
            ws.Cells[1, 2].Value = "简称";
            ws.Cells[1, 3].Value = "收件人邮箱";
            ws.Cells[1, 4].Value = "抄送邮箱";
            ws.Cells[1, 5].Value = "密送邮箱";
            ws.Cells[1, 6].Value = "备注";
            using (var range = ws.Cells[1, 1, 1, 6])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }
            for (int i = 0; i < CurrentRecipients.Count; i++)
            {
                var r = CurrentRecipients[i];
                ws.Cells[i + 2, 1].Value = r.Name;
                ws.Cells[i + 2, 2].Value = r.ShortName;
                ws.Cells[i + 2, 3].Value = r.ToEmails;
                ws.Cells[i + 2, 4].Value = r.CcEmails;
                ws.Cells[i + 2, 5].Value = r.BccEmails;
                ws.Cells[i + 2, 6].Value = r.Remark;
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            package.SaveAs(new System.IO.FileInfo(dialog.FileName));
            System.Windows.MessageBox.Show($"成功导出 {CurrentRecipients.Count} 个发送对象", "导出完成");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"导出失败: {ex.Message}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 从其他分组中复制收件人到当前分组。弹出分组选择对话框供用户选择。
    /// </summary>
    [RelayCommand]
    private void ImportFromOtherGroup()
    {
        if (SelectedGroup == null) return;
        var others = Groups.Where(g => g.Id != SelectedGroup.Id).ToList();
        if (others.Count == 0)
        {
            System.Windows.MessageBox.Show("没有其他分组可复制", "提示");
            return;
        }
        var dlg = new Views.GroupSelectDialog(others)
        {
            Owner = System.Windows.Application.Current.MainWindow
        };
        if (dlg.ShowDialog() == true && dlg.SelectedRecipients.Count > 0)
        {
            MergeRecipients(dlg.SelectedRecipients.ToList());
        }
    }

    /// <summary>
    /// 全选所有收件人。
    /// </summary>
    public void SelectAll()
    {
        foreach (var r in CurrentRecipients) r.IsSelected = true;
        UpdateCounts();
    }

    /// <summary>
    /// 取消全选所有收件人。
    /// </summary>
    public void DeselectAll()
    {
        foreach (var r in CurrentRecipients) r.IsSelected = false;
        UpdateCounts();
    }

    /// <summary>
    /// 更新选中数量和总数量的统计。
    /// </summary>
    public void UpdateCounts()
    {
        var all = Groups.SelectMany(g => g.Recipients).ToList();
        SelectedCount = all.Count(r => r.IsSelected);
        TotalCount = all.Count;
    }

    /// <summary>
    /// 从 Excel 文件中读取收件人数据。
    /// Excel 格式：第1列名称，第2列简称，第3列收件人邮箱，第4列抄送邮箱，第5列密送邮箱，第6列备注。
    /// 跳过名称为空的行。
    /// </summary>
    /// <param name="filePath">Excel 文件路径。</param>
    /// <returns>读取到的收件人列表。</returns>
    private List<Recipient> ImportFromFile(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("TailorMail");
        using var package = new ExcelPackage(new System.IO.FileInfo(filePath));
        var ws = package.Workbook.Worksheets[0];
        var recipients = new List<Recipient>();
        var rowCount = ws.Dimension?.Rows ?? 0;
        for (int row = 2; row <= rowCount; row++)
        {
            var name = ws.Cells[row, 1].Text?.Trim();
            if (string.IsNullOrEmpty(name)) continue;
            recipients.Add(new Recipient
            {
                Name = name,
                ShortName = ws.Cells[row, 2].Text?.Trim() ?? "",
                ToEmails = ws.Cells[row, 3].Text?.Trim() ?? "",
                CcEmails = ws.Cells[row, 4].Text?.Trim() ?? "",
                BccEmails = ws.Cells[row, 5].Text?.Trim() ?? "",
                Remark = ws.Cells[row, 6].Text?.Trim() ?? ""
            });
        }
        return recipients;
    }

    /// <summary>
    /// 同步收件人数据并保存到数据服务。
    /// 保存前先从文件重新加载最新数据，将当前分组的收件人变更合并进去，
    /// 仅更新收件人的基本字段（名称、简称、邮箱、备注、选择状态），
    /// 保留文件中的变量数据，避免覆盖其他页面（如变量配置页）已保存的修改。
    /// </summary>
    public void SaveAll()
    {
        SyncRecipientsToGroup();

        var latestGroups = _dataService.LoadRecipientGroups();

        foreach (var group in Groups)
        {
            var latestGroup = latestGroups.FirstOrDefault(g => g.Id == group.Id);
            if (latestGroup != null)
            {
                latestGroup.Name = group.Name;

                var updatedRecipients = new List<Recipient>();
                foreach (var r in group.Recipients)
                {
                    var existingInFile = latestGroup.Recipients.FirstOrDefault(lr => lr.Id == r.Id);
                    if (existingInFile != null)
                    {
                        existingInFile.Name = r.Name;
                        existingInFile.ShortName = r.ShortName;
                        existingInFile.ToEmails = r.ToEmails;
                        existingInFile.CcEmails = r.CcEmails;
                        existingInFile.BccEmails = r.BccEmails;
                        existingInFile.Remark = r.Remark;
                        existingInFile.IsSelected = r.IsSelected;
                        updatedRecipients.Add(existingInFile);
                    }
                    else
                    {
                        updatedRecipients.Add(r);
                    }
                }
                latestGroup.Recipients = updatedRecipients;
            }
            else
            {
                latestGroups.Add(group);
            }
        }

        _dataService.SaveRecipientGroups(latestGroups);
    }
}
