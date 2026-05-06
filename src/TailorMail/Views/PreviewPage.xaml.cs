using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class PreviewPage : UserControl, IRefreshable
{
    private readonly PreviewViewModel _viewModel;
    private string? _lastSelectedRecipientId;

    public PreviewPage()
    {
        InitializeComponent();
        _viewModel = new PreviewViewModel(App.DataService);
        DataContext = _viewModel;
        _viewModel.SelectedRecipientChanged += UpdateBrowser;
    }

    public void RefreshData()
    {
        _viewModel.LoadData();
        ListGroups.ItemsSource = _viewModel.SelectedRecipients;
        if (_viewModel.SelectedRecipients.Count > 0)
        {
            // Try to restore previous selection
            var restoreIndex = -1;
            if (_lastSelectedRecipientId != null)
            {
                for (int i = 0; i < _viewModel.SelectedRecipients.Count; i++)
                {
                    if (_viewModel.SelectedRecipients[i].Id == _lastSelectedRecipientId)
                    {
                        restoreIndex = i;
                        break;
                    }
                }
            }

            ListGroups.SelectedIndex = -1;
            Dispatcher.BeginInvoke(() =>
            {
                ListGroups.SelectedIndex = restoreIndex >= 0 ? restoreIndex : 0;
            });
        }
    }

    private void OnGroupSelected(object sender, SelectionChangedEventArgs e)
    {
        if (ListGroups.SelectedItem is Recipient r)
        {
            _lastSelectedRecipientId = r.Id;
            _viewModel.SelectRecipient(r);
            UpdateAttachmentList();
            UpdateBrowser();
        }
    }

    private void UpdateAttachmentList()
    {
        PanelAttachments.Items.Clear();
        var config = App.DataService.LoadAttachmentConfig();
        var commonFiles = config.CommonAttachments;
        var unitAtt = config.RecipientAttachments.FirstOrDefault(ua =>
            _viewModel.SelectedRecipient != null && ua.RecipientId == _viewModel.SelectedRecipient.Id);
        var specialFiles = unitAtt?.Files ?? [];

        long totalSize = 0;
        var allFiles = new HashSet<string>();
        foreach (var file in commonFiles)
        {
            PanelAttachments.Items.Add(CreateAttachmentItem("【公共】", file));
            allFiles.Add(file);
        }
        foreach (var file in specialFiles)
        {
            PanelAttachments.Items.Add(CreateAttachmentItem("【专有】", file));
            allFiles.Add(file);
        }

        foreach (var file in allFiles)
        {
            try
            {
                if (System.IO.File.Exists(file))
                    totalSize += new System.IO.FileInfo(file).Length;
            }
            catch { }
        }

        TxtAttachmentTotal.Text = allFiles.Count > 0
            ? $"共 {allFiles.Count} 个文件，合计 {FormatFileSize(totalSize)}"
            : "无附件";
    }

    private StackPanel CreateAttachmentItem(string prefix, string filePath)
    {
        var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 2) };
        var prefixBlock = new TextBlock
        {
            Text = prefix + " ",
            Foreground = (System.Windows.Media.Brush)FindResource("TextSecondaryBrush"),
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center
        };

        if (!System.IO.File.Exists(filePath))
        {
            var notFoundBlock = new TextBlock
            {
                Text = System.IO.Path.GetFileName(filePath) + " (文件不存在)",
                FontSize = 12,
                Foreground = (System.Windows.Media.Brush)FindResource("DangerBrush"),
                VerticalAlignment = VerticalAlignment.Center,
                TextTrimming = TextTrimming.CharacterEllipsis,
                ToolTip = filePath
            };
            panel.Children.Add(prefixBlock);
            panel.Children.Add(notFoundBlock);
            return panel;
        }

        var link = new Hyperlink(new Run(System.IO.Path.GetFileName(filePath)))
        {
            NavigateUri = new Uri(filePath),
            Foreground = (System.Windows.Media.Brush)FindResource("AccentBrush")
        };
        link.RequestNavigate += OnAttachmentLinkClick;
        var linkBlock = new TextBlock(link)
        {
            FontSize = 12,
            VerticalAlignment = VerticalAlignment.Center,
            ToolTip = filePath
        };
        panel.Children.Add(prefixBlock);
        panel.Children.Add(linkBlock);

        try
        {
            var info = new System.IO.FileInfo(filePath);
            if (info.Exists)
            {
                var sizeText = FormatFileSize(info.Length);
                var sizeBlock = new TextBlock
                {
                    Text = sizeText,
                    FontSize = 11,
                    Foreground = (System.Windows.Media.Brush)FindResource("TextTertiaryBrush"),
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(6, 0, 0, 0)
                };
                panel.Children.Add(sizeBlock);
            }
        }
        catch { }

        return panel;
    }

    private void OnAttachmentLinkClick(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.OriginalString)
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            App.ShowError($"无法打开文件: {ex.Message}");
        }
    }

    private void UpdateBrowser()
    {
        var doc = _viewModel.GetPreviewDocument();
        PreviewViewer.Document = doc;
    }

    private void OnSearchTextChanged(object sender, TextChangedEventArgs e)
    {
        var searchText = ((System.Windows.Controls.TextBox)sender).Text.Trim().ToLower();
        var view = CollectionViewSource.GetDefaultView(ListGroups.ItemsSource);
        if (view == null) return;
        view.Filter = item =>
        {
            if (string.IsNullOrEmpty(searchText)) return true;
            if (item is Recipient r)
            {
                return (r.Name?.ToLower().Contains(searchText) == true) ||
                       (r.ToEmails?.ToLower().Contains(searchText) == true);
            }
            return false;
        };
    }

    private static string FormatFileSize(long bytes)
    {
        string[] suffixes = { "B", "KB", "MB", "GB" };
        var order = 0;
        double size = bytes;
        while (size >= 1024 && order < suffixes.Length - 1)
        {
            order++;
            size /= 1024;
        }
        return order == 0 ? $"{bytes} {suffixes[order]}" : $"{size:0.#} {suffixes[order]}";
    }

    public void FocusSearch()
    {
        SearchBox?.Focus();
    }
}
