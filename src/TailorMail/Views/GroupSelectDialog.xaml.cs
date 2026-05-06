using System.Windows;
using System.Windows.Controls;
using TailorMail.Models;
using Wpf.Ui.Controls;

namespace TailorMail.Views;

public class GroupCheckItem
{
    public string GroupId { get; set; } = "";
    public string GroupName { get; set; } = "";
    public bool? IsChecked { get; set; }
    public List<RecipientCheckItem> Recipients { get; set; } = [];
}

public class RecipientCheckItem
{
    public Recipient Original { get; set; } = null!;
    public string RecipientName { get; set; } = "";
    public bool IsChecked { get; set; }
}

public partial class GroupSelectDialog : FluentWindow
{
    private readonly List<GroupCheckItem> _items;
    private bool _syncing;

    public List<Recipient> SelectedRecipients { get; private set; } = [];

    public GroupSelectDialog(List<RecipientGroup> groups)
    {
        InitializeComponent();
        Owner = Application.Current.MainWindow;
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
        if (_syncing) return;
        _syncing = true;
        try
        {
            if (sender is CheckBox cb && cb.DataContext is GroupCheckItem item)
            {
                var isChecked = cb.IsChecked == true;
                item.IsChecked = isChecked;
                foreach (var r in item.Recipients)
                    r.IsChecked = isChecked;
                UpdateGroupCheckBoxVisual(cb, isChecked ? true : false);
            }
        }
        finally
        {
            _syncing = false;
        }
    }

    private void OnRecipientCheckClick(object sender, RoutedEventArgs e)
    {
        if (_syncing) return;
        _syncing = true;
        try
        {
            if (sender is CheckBox && sender is FrameworkElement fe)
            {
                var treeItem = FindVisualParent<System.Windows.Controls.TreeViewItem>(fe);
                if (treeItem == null) return;

                var groupData = treeItem.DataContext as GroupCheckItem;
                if (groupData == null) return;

                var allChecked = groupData.Recipients.All(r => r.IsChecked);
                var anyChecked = groupData.Recipients.Any(r => r.IsChecked);

                if (allChecked)
                    groupData.IsChecked = true;
                else if (anyChecked)
                    groupData.IsChecked = null;
                else
                    groupData.IsChecked = false;

                var groupCheckBox = FindChildCheckBox(treeItem);
                if (groupCheckBox != null)
                {
                    groupCheckBox.IsThreeState = !allChecked && anyChecked;
                    groupCheckBox.IsChecked = groupData.IsChecked;
                }
            }
        }
        finally
        {
            _syncing = false;
        }
    }

    private void UpdateGroupCheckBoxVisual(CheckBox cb, bool isChecked)
    {
        cb.IsThreeState = false;
        cb.IsChecked = isChecked;
    }

    private static CheckBox? FindChildCheckBox(DependencyObject parent)
    {
        for (int i = 0; i < System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = System.Windows.Media.VisualTreeHelper.GetChild(parent, i);
            if (child is CheckBox cb) return cb;
            var result = FindChildCheckBox(child);
            if (result != null) return result;
        }
        return null;
    }

    private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
    {
        var parent = System.Windows.Media.VisualTreeHelper.GetParent(child);
        while (parent != null)
        {
            if (parent is T result) return result;
            parent = System.Windows.Media.VisualTreeHelper.GetParent(parent);
        }
        return null;
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