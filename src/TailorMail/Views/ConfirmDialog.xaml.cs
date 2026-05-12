using System.Windows;

namespace TailorMail.Views;

public partial class ConfirmDialog
{
    public string Message
    {
        get => MessageText.Text;
        set => MessageText.Text = value;
    }

    public ConfirmDialog()
    {
        InitializeComponent();
        Owner = Application.Current.MainWindow;
        Loaded += (_, _) => BtnYes.Focus();
    }

    private void BtnYes_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private void BtnNo_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }
}
