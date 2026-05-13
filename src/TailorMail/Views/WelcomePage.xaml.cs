using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace TailorMail.Views;

public partial class WelcomePage : UserControl
{
    public event Action? StartClicked;

    public WelcomePage()
    {
        InitializeComponent();
        Loaded += OnLoaded;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        // Version
        try
        {
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            if (version != null)
                TxtVersion.Text = $"v{version.Major}.{version.Minor}.{version.Build}";
        }
        catch { }
    }

    private void BtnStart_Click(object sender, RoutedEventArgs e)
    {
        StartClicked?.Invoke();
    }

    private void OnLinkClick(object sender, RequestNavigateEventArgs e)
    {
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(e.Uri.AbsoluteUri)
        {
            UseShellExecute = true
        });
        e.Handled = true;
    }
}