using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;
using OfficeOpenXml;

namespace TailorMail.ViewModels;

/// <summary>
/// 变量管理视图模型，负责自定义变量的增删、导入导出以及邮件模板占位符替换。
/// 内置占位符包括 {名称}、{简称}，
/// 用户可添加自定义变量并在邮件模板中使用 {变量名} 格式引用。
/// </summary>
public partial class VariablesViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    /// <summary>
    /// 获取或设置当前选中的收件人列表（仅包含 IsSelected 为 true 的收件人）。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<Recipient> _selectedRecipients = [];

    /// <summary>
    /// 获取或设置所有自定义变量名称的列表。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<string> _variableNames = [];

    /// <summary>
    /// 获取或设置新变量的名称输入。
    /// </summary>
    [ObservableProperty]
    private string _newVariableName = string.Empty;

    public VariablesViewModel(IDataService dataService)
    {
        _dataService = dataService;
        LoadData();
    }

    /// <summary>
    /// 加载选中的收件人列表和变量名称。
    /// </summary>
    public void LoadData()
    {
        var groups = _dataService.LoadRecipientGroups();
        SelectedRecipients = new ObservableCollection<Recipient>(
            groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected));
        RefreshVariableNames();
    }

    /// <summary>
    /// 从数据服务刷新变量名称列表，收集所有收件人中出现的变量键名。
    /// </summary>
    private void RefreshVariableNames()
    {
        var groups = _dataService.LoadRecipientGroups();
        var names = new HashSet<string>();
        foreach (var r in groups.SelectMany(g => g.Recipients))
            foreach (var key in r.Variables.Keys)
                names.Add(key);
        VariableNames = new ObservableCollection<string>(names.OrderBy(n => n));
    }

    /// <summary>
    /// 添加新的自定义变量。会为所有收件人初始化该变量的空值。
    /// </summary>
    public void AddVariableAndSave()
    {
        if (string.IsNullOrWhiteSpace(NewVariableName)) return;
        if (VariableNames.Contains(NewVariableName)) return;

        var varName = NewVariableName;
        VariableNames.Add(varName);

        // 为所有收件人添加该变量的空值条目
        var groups = _dataService.LoadRecipientGroups();
        foreach (var r in groups.SelectMany(g => g.Recipients))
        {
            if (!r.Variables.ContainsKey(varName))
                r.Variables[varName] = string.Empty;
        }
        _dataService.SaveRecipientGroups(groups);

        // 同步更新界面上的收件人对象
        foreach (var r in SelectedRecipients)
        {
            if (!r.Variables.ContainsKey(varName))
                r.Variables[varName] = string.Empty;
        }

        NewVariableName = string.Empty;
    }

    /// <summary>
    /// 删除指定的自定义变量。会从所有收件人中移除该变量。
    /// </summary>
    /// <param name="name">要删除的变量名称。</param>
    public void DeleteVariableAndSave(string name)
    {
        if (!VariableNames.Contains(name)) return;
        VariableNames.Remove(name);

        var groups = _dataService.LoadRecipientGroups();
        foreach (var r in groups.SelectMany(g => g.Recipients))
            r.Variables.Remove(name);
        _dataService.SaveRecipientGroups(groups);

        foreach (var r in SelectedRecipients)
            r.Variables.Remove(name);
    }

    /// <summary>
    /// 从 Excel 文件导入变量数据。
    /// Excel 格式：第1列为收件人名称，第2列起每列为一个变量（表头为变量名）。
    /// 导入时仅在值非空时写入，避免覆盖已有数据。
    /// </summary>
    [RelayCommand]
    private void ImportVariables()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
            Title = "导入变量"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("TailorMail");
            using var package = new ExcelPackage(new System.IO.FileInfo(dialog.FileName));
            var ws = package.Workbook.Worksheets[0];
            var rowCount = ws.Dimension?.Rows ?? 0;
            var colCount = ws.Dimension?.Columns ?? 0;

            // 读取表头（第2列起为变量名）
            var headers = new List<string>();
            for (int col = 2; col <= colCount; col++)
                headers.Add(ws.Cells[1, col].Text?.Trim() ?? "");

            // 将新变量名添加到变量列表
            foreach (var h in headers.Where(h => !string.IsNullOrEmpty(h) && !VariableNames.Contains(h)))
                VariableNames.Add(h);

            // 按行读取数据，按名称匹配收件人并写入变量值
            for (int row = 2; row <= rowCount; row++)
            {
                var name = ws.Cells[row, 1].Text?.Trim() ?? "";
                var recipient = SelectedRecipients.FirstOrDefault(r => r.Name == name);
                if (recipient == null) continue;
                for (int col = 2; col <= colCount; col++)
                {
                    var varName = headers[col - 2];
                    if (string.IsNullOrEmpty(varName)) continue;
                    var val = ws.Cells[row, col].Text?.Trim() ?? "";
                    // 仅在值非空时写入，避免覆盖已有数据
                    if (!string.IsNullOrEmpty(val))
                        recipient.Variables[varName] = val;
                }
            }
            SaveAll();
            System.Windows.MessageBox.Show("导入完成", "提示");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"导入失败: {ex.Message}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 将变量数据导出为 Excel 文件。
    /// Excel 格式：第1列为收件人名称，第2列起每列为一个变量的值。
    /// </summary>
    [RelayCommand]
    private void ExportVariables()
    {
        if (SelectedRecipients.Count == 0 || VariableNames.Count == 0) return;
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "Excel文件|*.xlsx",
            FileName = "变量数据.xlsx"
        };
        if (dialog.ShowDialog() != true) return;
        try
        {
            ExcelPackage.License.SetNonCommercialPersonal("TailorMail");
            using var package = new ExcelPackage();
            var ws = package.Workbook.Worksheets.Add("变量");
            ws.Cells[1, 1].Value = "名称";
            for (int i = 0; i < VariableNames.Count; i++)
                ws.Cells[1, i + 2].Value = VariableNames[i];
            using (var range = ws.Cells[1, 1, 1, VariableNames.Count + 1])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
            }
            for (int i = 0; i < SelectedRecipients.Count; i++)
            {
                var r = SelectedRecipients[i];
                ws.Cells[i + 2, 1].Value = r.Name;
                for (int j = 0; j < VariableNames.Count; j++)
                    ws.Cells[i + 2, j + 2].Value = r.Variables.GetValueOrDefault(VariableNames[j], "");
            }
            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            package.SaveAs(new System.IO.FileInfo(dialog.FileName));
            System.Windows.MessageBox.Show("导出完成", "提示");
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show($"导出失败: {ex.Message}", "错误",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 对邮件正文进行变量替换。先替换内置占位符（{名称}、{简称}），
    /// 再替换自定义变量占位符（{变量名}）。
    /// </summary>
    /// <param name="body">包含占位符的原始文本。</param>
    /// <param name="recipient">收件人对象，提供变量值。</param>
    /// <returns>替换后的文本。</returns>
    public string ProcessBody(string body, Recipient recipient)
    {
        var result = body.Replace("{名称}", recipient.Name)
                         .Replace("{简称}", recipient.ShortName);
        foreach (var kvp in recipient.Variables)
            result = result.Replace($"{{{kvp.Key}}}", kvp.Value);
        return result;
    }

    /// <summary>
    /// 获取所有可用的变量占位符列表，包括内置占位符和自定义变量占位符。
    /// </summary>
    /// <returns>占位符字符串列表，如 "{名称}"、"{简称}"、"{自定义变量}" 等。</returns>
    public List<string> GetAllVariablePlaceholders()
    {
        var list = new List<string> { "{名称}", "{简称}" };
        foreach (var name in VariableNames)
            list.Add($"{{{name}}}");
        return list;
    }

    /// <summary>
    /// 将选中收件人的变量数据保存到数据服务。
    /// </summary>
    public void SaveAll()
    {
        var groups = _dataService.LoadRecipientGroups();
        foreach (var recipient in SelectedRecipients)
        {
            var group = groups.FirstOrDefault(g => g.Recipients.Any(r => r.Id == recipient.Id));
            if (group == null) continue;
            var existing = group.Recipients.FirstOrDefault(r => r.Id == recipient.Id);
            if (existing != null)
            {
                existing.Variables = new Dictionary<string, string>(recipient.Variables);
            }
        }
        _dataService.SaveRecipientGroups(groups);
    }
}
