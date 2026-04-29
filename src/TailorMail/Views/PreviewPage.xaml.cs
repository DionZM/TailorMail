using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using TailorMail.Models;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 邮件预览页面，提供收件人选择、邮件主题/正文/附件的预览功能。
/// </summary>
public partial class PreviewPage : UserControl, IRefreshable
{
    private readonly PreviewViewModel _viewModel;

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
            ListGroups.SelectedIndex = -1;
            Dispatcher.BeginInvoke(() =>
            {
                ListGroups.SelectedIndex = 0;
            });
        }
    }

    private void OnGroupSelected(object sender, SelectionChangedEventArgs e)
    {
        if (ListGroups.SelectedItem is Recipient r)
        {
            _viewModel.SelectRecipient(r);
            TxtSubject.Text = _viewModel.PreviewSubject;
            TxtTo.Text = _viewModel.PreviewTo;
            TxtCc.Text = _viewModel.PreviewCc;
            TxtBcc.Text = _viewModel.PreviewBcc;

            var hasCc = !string.IsNullOrEmpty(_viewModel.PreviewCc);
            var hasBcc = !string.IsNullOrEmpty(_viewModel.PreviewBcc);
            CcBccPanel.Visibility = (hasCc || hasBcc) ? Visibility.Visible : Visibility.Collapsed;
            BccLabel.Visibility = hasBcc ? Visibility.Visible : Visibility.Collapsed;
            TxtBcc.Visibility = hasBcc ? Visibility.Visible : Visibility.Collapsed;

            UpdateAttachmentList();
            UpdateBrowser();
        }
    }

    private void UpdateAttachmentList()
    {
        PanelAttachments.Items.Clear();
        var config = App.DataService.LoadAttachmentConfig();
        var commonFiles = config.CommonAttachments;
        var unitAtt = config.UnitAttachments.FirstOrDefault(ua =>
            _viewModel.SelectedRecipient != null && ua.RecipientId == _viewModel.SelectedRecipient.Id);
        var specialFiles = unitAtt?.Files ?? [];

        foreach (var file in commonFiles)
        {
            PanelAttachments.Items.Add(CreateAttachmentItem("【公共】", file));
        }
        foreach (var file in specialFiles)
        {
            PanelAttachments.Items.Add(CreateAttachmentItem("【专有】", file));
        }
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
        return panel;
    }

    private void OnAttachmentLinkClick(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        try
        {
            var filePath = Uri.UnescapeDataString(e.Uri.AbsolutePath);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath)
            {
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"无法打开文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UpdateBrowser()
    {
        var html = _viewModel.GetPreviewHtml();
        if (!string.IsNullOrEmpty(html))
            PreviewBrowser.NavigateToString(html);
    }

    private void OnPrevGroup(object sender, RoutedEventArgs e)
    {
        var idx = ListGroups.SelectedIndex;
        if (idx > 0) ListGroups.SelectedIndex = idx - 1;
    }

    private void OnNextGroup(object sender, RoutedEventArgs e)
    {
        var idx = ListGroups.SelectedIndex;
        if (idx < ListGroups.Items.Count - 1) ListGroups.SelectedIndex = idx + 1;
    }
}
