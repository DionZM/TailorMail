using System.Collections.ObjectModel;
using System.Text;
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
    private List<RecipientGroup> _groups = [];

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
        _groups = _dataService.LoadRecipientGroups();
        SelectedRecipients = new ObservableCollection<Recipient>(
            _groups.SelectMany(g => g.Recipients).Where(r => r.IsSelected));
        RefreshVariableNames();
    }

    /// <summary>
    /// 从 VariableNames 表加载变量名称列表。
    /// </summary>
    private void RefreshVariableNames()
    {
        var names = _dataService.LoadVariableNames();
        VariableNames = new ObservableCollection<string>(names);
    }

    /// <summary>
    /// 添加新的自定义变量。仅在 VariableNames 表中插入一行，不修改收件人数据。
    /// </summary>
    public void AddVariableAndSave()
    {
        if (string.IsNullOrWhiteSpace(NewVariableName)) return;
        if (VariableNames.Contains(NewVariableName)) return;

        var varName = NewVariableName;
        _dataService.AddVariableName(varName);
        VariableNames.Add(varName);

        // 为界面上已加载的收件人添加该变量的空值（代码层处理，无需写库）
        foreach (var r in SelectedRecipients)
        {
            if (!r.Variables.ContainsKey(varName))
                r.Variables[varName] = string.Empty;
        }

        NewVariableName = string.Empty;
    }

    /// <summary>
    /// 重命名变量。通过数据服务批量更新数据库中的键名和模板引用。
    /// </summary>
    public void RenameVariableAndSave(string oldName, string newName)
    {
        if (!VariableNames.Contains(oldName) || VariableNames.Contains(newName)) return;

        _dataService.RenameVariableName(oldName, newName);

        VariableNames.Remove(oldName);
        VariableNames.Add(newName);

        // 同步更新内存中的收件人变量
        foreach (var r in SelectedRecipients)
        {
            if (r.Variables.TryGetValue(oldName, out var value))
            {
                r.Variables.Remove(oldName);
                r.Variables[newName] = value;
            }
        }

        // Update template references
        var settings = _dataService.LoadSettings();
        var oldPlaceholder = $"{{{oldName}}}";
        var newPlaceholder = $"{{{newName}}}";
        if (settings.LastSubject?.Contains(oldPlaceholder) == true)
            settings.LastSubject = settings.LastSubject.Replace(oldPlaceholder, newPlaceholder);
        if (settings.LastBody?.Contains(oldPlaceholder) == true)
            settings.LastBody = settings.LastBody.Replace(oldPlaceholder, newPlaceholder);
        if (settings.LastBodyXaml?.Contains(oldPlaceholder) == true)
            settings.LastBodyXaml = settings.LastBodyXaml.Replace(oldPlaceholder, newPlaceholder);
        _dataService.SaveSettings(settings);
    }

    /// <summary>
    /// 删除指定的自定义变量。通过数据服务批量从数据库中移除键名。
    /// </summary>
    /// <param name="name">要删除的变量名称。</param>
    public void DeleteVariableAndSave(string name)
    {
        if (!VariableNames.Contains(name)) return;

        _dataService.DeleteVariableName(name);
        VariableNames.Remove(name);

        // 同步清理内存中收件人的变量
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
            // EPPlus license already set in App.xaml.cs (PERF-27)
            using var package = new ExcelPackage(new System.IO.FileInfo(dialog.FileName));
            var ws = package.Workbook.Worksheets[0];
            var rowCount = ws.Dimension?.Rows ?? 0;
            var colCount = ws.Dimension?.Columns ?? 0;

            // 读取表头（第2列起为变量名）
            var headers = new List<string>();
            for (int col = 2; col <= colCount; col++)
                headers.Add(ws.Cells[1, col].Text?.Trim() ?? "");

            // 将新变量名注册到 VariableNames 表
            foreach (var h in headers.Where(h => !string.IsNullOrEmpty(h) && !VariableNames.Contains(h)))
            {
                _dataService.AddVariableName(h);
                VariableNames.Add(h);
            }

            // 按行读取数据，按名称匹配收件人并写入变量值
            // M-05: Pre-build dictionary for O(1) name lookup
            var recipientByName = SelectedRecipients.ToDictionary(r => r.Name);

            for (int row = 2; row <= rowCount; row++)
            {
                var name = ws.Cells[row, 1].Text?.Trim() ?? "";
                if (!recipientByName.TryGetValue(name, out var recipient)) continue;
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
            App.ShowSuccess("导入完成");
        }
        catch (System.IO.IOException)
        {
            App.ShowError("导入失败: 文件被占用或无法访问，请关闭文件后重试");
        }
        catch (Exception ex)
        {
            App.ShowError($"导入失败: {ex.Message}");
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
            // EPPlus license already set in App.xaml.cs (PERF-27)
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
            App.ShowSuccess("导出完成");
        }
        catch (Exception ex)
        {
            App.ShowError($"导出失败: {ex.Message}");
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
    /// High-performance single-pass variable substitution using StringBuilder.
    /// Scans the template once, replacing all {key} placeholders via dictionary lookup.
    /// Avoids the N intermediate string allocations of chained .Replace() calls.
    /// </summary>
    public static string ProcessBodyFast(string body, Recipient recipient)
    {
        // Build the replacement dictionary
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["名称"] = recipient.Name ?? "",
            ["简称"] = recipient.ShortName ?? ""
        };
        foreach (var kvp in recipient.Variables)
            replacements[kvp.Key] = kvp.Value ?? "";

        var sb = new StringBuilder(body.Length);
        for (int i = 0; i < body.Length;)
        {
            if (body[i] == '{')
            {
                int closeIndex = body.IndexOf('}', i + 1);
                if (closeIndex > i + 1)
                {
                    var key = body.AsSpan(i + 1, closeIndex - i - 1);
                    // Try exact key match
                    if (replacements.TryGetValue(key.ToString(), out var value))
                    {
                        sb.Append(value);
                        i = closeIndex + 1;
                        continue;
                    }
                }
                sb.Append('{');
                i++;
            }
            else
            {
                sb.Append(body[i]);
                i++;
            }
        }
        return sb.ToString();
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
    /// 使用单行 UPDATE 只写入被修改的收件人。
    /// </summary>
    public void SaveAll()
    {
        _dataService.SaveRecipientVariables(SelectedRecipients);
    }
}
