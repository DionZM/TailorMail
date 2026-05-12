using System.Collections.ObjectModel;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;
using OfficeOpenXml;

namespace TailorMail.ViewModels;

public partial class RecipientsViewModel : ObservableObject
{
    private readonly IDataService _dataService;
    private DispatcherTimer? _saveTimer;
    private bool _hasUnsavedChanges;

    [ObservableProperty]
    private ObservableCollection<RecipientGroup> _groups = [];

    [ObservableProperty]
    private RecipientGroup? _selectedGroup;

    [ObservableProperty]
    private ObservableCollection<Recipient> _currentRecipients = [];

    [ObservableProperty]
    private string _newGroupName = string.Empty;

    [ObservableProperty]
    private int _selectedCount;

    [ObservableProperty]
    private int _totalCount;

    public event Action? SelectedGroupChanged;

    public RecipientsViewModel(IDataService dataService)
    {
        _dataService = dataService;
        LoadGroups();
    }

    public void LoadGroups()
    {
        var groups = _dataService.LoadRecipientGroups();
        Groups = new ObservableCollection<RecipientGroup>(groups);
        if (Groups.Count > 0)
            SelectedGroup = Groups[0];
    }

    partial void OnSelectedGroupChanging(RecipientGroup? value)
    {
        FlushSave();
    }

    partial void OnSelectedGroupChanged(RecipientGroup? value)
    {
        if (value != null)
        {
            CurrentRecipients = new ObservableCollection<Recipient>(value.Recipients);
        }
        else
        {
            CurrentRecipients = [];
        }
        UpdateCounts();
        SelectedGroupChanged?.Invoke();
    }

    [RelayCommand]
    private void AddGroup()
    {
        var inputDlg = new Views.InputDialog
        {
            Title = "添加分组",
            Prompt = "请输入分组名称："
        };
        
        if (inputDlg.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputDlg.InputText))
        {
            var newName = inputDlg.InputText.Trim();
            if (Groups.Any(g => g.Name == newName))
            {
                App.ShowWarning("分组名已存在");
                return;
            }
            Groups.Add(new RecipientGroup { Name = newName });
            SaveAll();
        }
    }

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

    [RelayCommand]
    private void DeleteRecipient(Recipient r)
    {
        if (SelectedGroup == null) return;
        SelectedGroup.Recipients.Remove(r);
        CurrentRecipients.Remove(r);
        ScheduleSave();
        UpdateCounts();
    }

    private bool HasEmptyNameRow()
    {
        return CurrentRecipients.Any(r => string.IsNullOrEmpty(r.Name));
    }

    [RelayCommand]
    private void AddRecipient()
    {
        if (SelectedGroup == null) return;
        if (HasEmptyNameRow()) return;
        
        var newRecipient = new Recipient { Name = "", IsSelected = false };
        CurrentRecipients.Add(newRecipient);
        SelectedGroup.Recipients.Add(newRecipient);
        ScheduleSave();
    }

    [RelayCommand]
    private void DeleteSelectedRecipients()
    {
        if (SelectedGroup == null) return;
        
        var selected = CurrentRecipients.Where(r => r.IsSelected).ToList();
        if (selected.Count == 0)
        {
            App.ShowWarning("请先选择要删除的收件人");
            return;
        }
        
        if (System.Windows.MessageBox.Show($"确定删除选中的 {selected.Count} 个收件人？", "确认删除",
            System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question) != System.Windows.MessageBoxResult.Yes) return;
        
        foreach (var r in selected)
        {
            CurrentRecipients.Remove(r);
            SelectedGroup.Recipients.Remove(r);
        }
        ScheduleSave();
        UpdateCounts();
    }

    public void MoveUp(Recipient r)
    {
        if (SelectedGroup == null) return;
        var idx = CurrentRecipients.IndexOf(r);
        if (idx <= 0) return;
        CurrentRecipients.Move(idx, idx - 1);
        SyncRecipientsToGroup();
        ScheduleSave();
    }

    public void MoveDown(Recipient r)
    {
        if (SelectedGroup == null) return;
        var idx = CurrentRecipients.IndexOf(r);
        if (idx < 0 || idx >= CurrentRecipients.Count - 1) return;
        CurrentRecipients.Move(idx, idx + 1);
        SyncRecipientsToGroup();
        ScheduleSave();
    }

    public void SyncRecipientsToGroup()
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

    public void RemoveEmptyRows()
    {
        var emptyRows = CurrentRecipients.Where(r => string.IsNullOrEmpty(r.Name)).ToList();
        foreach (var r in emptyRows)
        {
            CurrentRecipients.Remove(r);
            if (SelectedGroup != null)
                SelectedGroup.Recipients.Remove(r);
        }
        ScheduleSave();
    }

    private void MergeRecipients(List<Recipient> incoming)
    {
        if (SelectedGroup == null) return;
        int updated = 0, added = 0;
        var nameIndex = new Dictionary<string, Recipient>();
        foreach (var r in CurrentRecipients)
        {
            if (!string.IsNullOrEmpty(r.Name))
                nameIndex[r.Name] = r;
        }

        foreach (var src in incoming)
        {
            if (nameIndex.TryGetValue(src.Name, out var existing))
            {
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
                nameIndex[src.Name] = clone;
                added++;
            }
        }
        SaveAll();
        UpdateCounts();
        App.ShowSuccess($"导入完成：新增 {added} 项，更新 {updated} 项");
    }

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
            App.ShowError($"导入失败: {ex.Message}");
        }
    }

    [RelayCommand]
    private void ExportToExcel()
    {
        if (SelectedGroup == null)
        {
            App.ShowWarning("请先选择一个分组");
            return;
        }
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel文件|*.xlsx",
            FileName = $"{SelectedGroup.Name}_发送对象.xlsx"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            // EPPlus license already set in App.xaml.cs (PERF-27)
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
            if (CurrentRecipients.Count == 0)
                App.ShowSuccess("已导出空白模板（仅含标题行），可填写后导入使用");
            else
                App.ShowSuccess($"成功导出 {CurrentRecipients.Count} 个发送对象");
        }
        catch (Exception ex)
        {
            App.ShowError($"导出失败: {ex.Message}");
        }
    }

    [RelayCommand]
    private void ImportFromOtherGroup()
    {
        if (SelectedGroup == null) return;
        var others = Groups.Where(g => g.Id != SelectedGroup.Id).ToList();
        if (others.Count == 0)
        {
            App.ShowNotification("没有其他分组可复制");
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

    public void SelectAll()
    {
        foreach (var r in CurrentRecipients) r.IsSelected = true;
        UpdateCounts();
    }

    public void DeselectAll()
    {
        foreach (var r in CurrentRecipients) r.IsSelected = false;
        UpdateCounts();
    }

    public void UpdateCounts()
    {
        int selected = 0;
        int total = CurrentRecipients.Count;
        for (int i = 0; i < CurrentRecipients.Count; i++)
        {
            if (CurrentRecipients[i].IsSelected) selected++;
        }
        SelectedCount = selected;
        TotalCount = total;
    }

    private List<Recipient> ImportFromFile(string filePath)
    {
        ExcelPackage.License.SetNonCommercialPersonal("TailorMail"); // L-12: Redundant but harmless — already set in App.xaml.cs
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

    public void ScheduleSave()
    {
        _hasUnsavedChanges = true;
        if (_saveTimer == null)
        {
            _saveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _saveTimer.Tick += (s, e) =>
            {
                _saveTimer.Stop();
                if (_hasUnsavedChanges)
                {
                    _hasUnsavedChanges = false;
                    SaveAll();
                }
            };
        }
        _saveTimer.Stop();
        _saveTimer.Start();
    }

    public void FlushSave()
    {
        if (_saveTimer != null)
        {
            _saveTimer.Stop();
        }
        if (_hasUnsavedChanges)
        {
            _hasUnsavedChanges = false;
            SaveAll();
        }
    }

    public void SaveAll()
    {
        SyncRecipientsToGroup();
        _dataService.SaveRecipientGroups(Groups.ToList());
    }
}
