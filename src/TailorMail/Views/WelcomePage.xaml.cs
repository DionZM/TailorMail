using System.Windows;
using System.Windows.Controls;

namespace TailorMail.Views;

public partial class WelcomePage : UserControl
{
    public event Action? StartClicked;

    public WelcomePage()
    {
        InitializeComponent();
    }

    private void BtnStart_Click(object sender, RoutedEventArgs e)
    {
        StartClicked?.Invoke();
    }
}