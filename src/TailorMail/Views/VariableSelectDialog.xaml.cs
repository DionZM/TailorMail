using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Wpf.Ui.Controls;

namespace TailorMail.Views;

public partial class VariableSelectDialog : FluentWindow
{
    public List<string> SelectedVariables { get; private set; } = [];
    public bool AllowSingle { get; set; }
    public string ConfirmText { get; set; } = "删除";
    public string HeaderText { get => TitleText.Text; set => TitleText.Text = value; }

    private readonly List<CheckBox> _checkBoxes = [];

    public VariableSelectDialog(List<string> variableNames)
    {
        InitializeComponent();
        Owner = Application.Current.MainWindow;

        foreach (var name in variableNames)
        {
            var cb = new CheckBox
            {
                Content = name,
                Margin = new Thickness(0, 4, 0, 4),
                FontFamily = (FontFamily)FindResource("UIFontFamily"),
                Tag = name
            };
            _checkBoxes.Add(cb);
            ItemsPanel.Children.Add(cb);
        }

        BtnOk.Content = ConfirmText;
    }

    private void OnOk(object sender, RoutedEventArgs e)
    {
        var checkedItems = _checkBoxes.Where(cb => cb.IsChecked == true).Select(cb => (string)cb.Tag!).ToList();

        if (AllowSingle)
        {
            if (checkedItems.Count != 1)
            {
                System.Windows.MessageBox.Show("请选择一个变量", "提示",
                    System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                return;
            }
        }
        else if (checkedItems.Count == 0)
        {
            System.Windows.MessageBox.Show("请至少选择一个变量", "提示",
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }

        SelectedVariables = checkedItems;
        DialogResult = true;
    }

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
