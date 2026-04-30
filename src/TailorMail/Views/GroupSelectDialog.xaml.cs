using System.Windows;
using TailorMail.Models;
using Wpf.Ui.Controls;

namespace TailorMail.Views;

/// <summary>
/// 分组选择对话框，用于从其他分组中选择收件人进行复制。
/// </summary>
public class GroupCheckItem
{
    public string GroupId { get; set; } = "";
    public string GroupName { get; set; } = "";
    public bool IsChecked { get; set; }
    public List<RecipientCheckItem> Recipients { get; set; } = [];
}

/// <summary>
/// 分组选择对话框，用于从其他分组中选择收件人进行复制。
/// </summary>
public class RecipientCheckItem
{
    public Recipient Original { get; set; } = null!;
    public string RecipientName { get; set; } = "";
    public bool IsChecked { get; set; }
}

/// <summary>
/// 分组选择对话框，用于从其他分组中选择收件人进行复制。
/// </summary>
public partial class GroupSelectDialog : FluentWindow
{
    private readonly List<GroupCheckItem> _items;

    public List<Recipient> SelectedRecipients { get; private set; } = [];

    public GroupSelectDialog(List<RecipientGroup> groups)
    {
        InitializeComponent();
        _items = groups.Select(g => new GroupCheckItem
        {
            GroupId = g.Id,
            GroupName = g.Name,
            IsChecked = false,
            Recipients = g.Recipients.Select(r => new RecipientCheckItem
            {
                Original = r,
                RecipientName = r.Name,
                IsChecked = false
            }).ToList()
        }).ToList();
        GroupTree.ItemsSource = _items;
    }

    private void OnGroupCheckClick(object sender, RoutedEventArgs e)
    {
        if (sender is System.Windows.Controls.CheckBox cb && cb.DataContext is GroupCheckItem item)
        {
            foreach (var r in item.Recipients)
                r.IsChecked = item.IsChecked;
            RefreshTree();
        }
    }

    private void OnRecipientCheckClick(object sender, RoutedEventArgs e)
    {
        RefreshTree();
    }

    private void RefreshTree()
    {
        var source = GroupTree.ItemsSource;
        GroupTree.ItemsSource = null;
        GroupTree.ItemsSource = source;
    }

    private void OnImport(object sender, RoutedEventArgs e)
    {
        SelectedRecipients = _items
            .SelectMany(g => g.Recipients)
            .Where(r => r.IsChecked)
            .Select(r => r.Original)
            .ToList();

        if (SelectedRecipients.Count == 0)
        {
            System.Windows.MessageBox.Show("请至少选择一个发送对象", "提示");
            return;
        }
        DialogResult = true;
        Close();
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
