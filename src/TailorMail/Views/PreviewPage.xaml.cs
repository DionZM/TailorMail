using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using TailorMail.Helpers;
using TailorMail.Models;
using TailorMail.Services;
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
        PreviewKeyDown += OnPreviewKeyDown;
        IsVisibleChanged += OnVisibilityChanged;
    }

    private void OnVisibilityChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if ((bool)e.NewValue)
            UpdateRecipientCounter();
    }

    private void OnPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        if (e.Key == Key.Up && Keyboard.Modifiers == ModifierKeys.None)
        {
            NavigateToPreviousRecipient();
            e.Handled = true;
        }
        else if (e.Key == Key.Down && Keyboard.Modifiers == ModifierKeys.None)
        {
            NavigateToNextRecipient();
            e.Handled = true;
        }
    }

    private void NavigateToPreviousRecipient()
    {
        if (ListGroups.Items.Count == 0) return;
        var currentIdx = ListGroups.SelectedIndex;
        if (currentIdx > 0)
        {
            ListGroups.SelectedIndex = currentIdx - 1;
        }
        else if (currentIdx < 0 && ListGroups.Items.Count > 0)
        {
            ListGroups.SelectedIndex = ListGroups.Items.Count - 1;
        }
    }

    private void NavigateToNextRecipient()
    {
        if (ListGroups.Items.Count == 0) return;
        var currentIdx = ListGroups.SelectedIndex;
        if (currentIdx < ListGroups.Items.Count - 1)
        {
            ListGroups.SelectedIndex = currentIdx + 1;
        }
        else if (currentIdx < 0 && ListGroups.Items.Count > 0)
        {
            ListGroups.SelectedIndex = 0;
        }
    }

    private void BtnPrevRecipient_Click(object sender, RoutedEventArgs e) => NavigateToPreviousRecipient();
    private void BtnNextRecipient_Click(object sender, RoutedEventArgs e) => NavigateToNextRecipient();

    private void UpdateRecipientCounter()
    {
        var total = ListGroups.Items.Count;
        var current = ListGroups.SelectedIndex + 1;
        TxtRecipientCounter.Text = total > 0 ? $"{current}/{total}" : "";
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
                Helpers.AnimationHelper.StaggerFadeIn(ListGroups);
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
            UpdateRecipientCounter();
        }
    }

    private void UpdateAttachmentList()
    {
        PanelAttachments.Items.Clear();
        var config = _viewModel.GetCachedAttachmentConfig();
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

        var ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
        var iconSymbol = ext switch
        {
            ".pdf" or ".doc" or ".docx" => Wpf.Ui.Controls.SymbolRegular.Document24,
            ".xls" or ".xlsx" or ".csv" => Wpf.Ui.Controls.SymbolRegular.Table24,
            ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".svg" or ".webp" => Wpf.Ui.Controls.SymbolRegular.Image24,
            ".zip" or ".rar" or ".7z" => Wpf.Ui.Controls.SymbolRegular.FolderZip24,
            ".ppt" or ".pptx" => Wpf.Ui.Controls.SymbolRegular.SlideText24,
            _ => Wpf.Ui.Controls.SymbolRegular.Document24
        };
        var icon = new Wpf.Ui.Controls.SymbolIcon
        {
            Symbol = iconSymbol,
            FontSize = 14,
            Foreground = (System.Windows.Media.Brush)FindResource("TextSecondaryBrush"),
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 4, 0)
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
            panel.Children.Add(icon);
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
        panel.Children.Add(icon);
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

    // UI-39: Zoom handler
    private void OnZoomChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        if (PreviewViewer == null || TxtZoom == null) return; // Fires during InitializeComponent
        var scale = e.NewValue / 100.0;
        TxtZoom.Text = $"{(int)e.NewValue}%";
        PreviewViewer.LayoutTransform = new System.Windows.Media.ScaleTransform(scale, scale);
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
