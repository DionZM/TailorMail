using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace TailorMail.Views;

public partial class InputDialog
{
    public string Prompt { get; set; } = "请输入：";
    
    public string InputText { get; set; } = string.Empty;
    
    public InputDialog()
    {
        InitializeComponent();
        DataContext = this;
        Owner = Application.Current.MainWindow;
        Loaded += (_, _) => InputTextBox.Focus();
    }

    private void OnTextChanged(object sender, TextChangedEventArgs e)
    {
        InputText = InputTextBox.Text;
        BtnOk.IsEnabled = !string.IsNullOrWhiteSpace(InputText);
    }

    private void OnInputKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && BtnOk.IsEnabled)
        {
            e.Handled = true;
            BtnOk_Click(sender, e);
        }
    }

    private void BtnOk_Click(object sender, RoutedEventArgs e)
    {
        InputText = InputTextBox.Text;
        DialogResult = true;
        Close();
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}