using System.Windows;
using System.Windows.Controls;
using TailorMail.ViewModels;

namespace TailorMail.Views;

/// <summary>
/// 变量选择对话框，用于选择要删除的自定义变量。
/// </summary>
public partial class VariableSelectDialog : Window
{
    public List<string> SelectedVariables { get; private set; } = [];

    public VariableSelectDialog(List<string> variableNames)
    {
        InitializeComponent();
        Title = "选择要删除的变量";
        Width = 320;
        Height = 400;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        ResizeMode = ResizeMode.NoResize;

        var panel = new StackPanel { Margin = new Thickness(16) };

        var hint = new System.Windows.Controls.TextBlock
        {
            Text = "请勾选要删除的变量：",
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 12)
        };
        panel.Children.Add(hint);

        var scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, MaxHeight = 250 };
        var itemsPanel = new StackPanel();
        var checkBoxes = new List<System.Windows.Controls.CheckBox>();

        foreach (var name in variableNames)
        {
            var cb = new System.Windows.Controls.CheckBox
            {
                Content = name,
                Margin = new Thickness(0, 4, 0, 4),
                Tag = name
            };
            checkBoxes.Add(cb);
            itemsPanel.Children.Add(cb);
        }
        scrollViewer.Content = itemsPanel;
        panel.Children.Add(scrollViewer);

        var btnPanel = new StackPanel
        {
            Orientation = System.Windows.Controls.Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };

        var btnOk = new System.Windows.Controls.Button
        {
            Content = "删除",
            Padding = new Thickness(20, 6, 20, 6),
            Margin = new Thickness(0, 0, 8, 0),
            FontWeight = FontWeights.SemiBold
        };
        btnOk.Click += (_, _) =>
        {
            SelectedVariables = checkBoxes.Where(cb => cb.IsChecked == true).Select(cb => (string)cb.Tag!).ToList();
            DialogResult = true;
        };

        var btnCancel = new System.Windows.Controls.Button
        {
            Content = "取消",
            Padding = new Thickness(20, 6, 20, 6)
        };
        btnCancel.Click += (_, _) => { DialogResult = false; };

        btnPanel.Children.Add(btnOk);
        btnPanel.Children.Add(btnCancel);
        panel.Children.Add(btnPanel);

        Content = panel;
    }
}
