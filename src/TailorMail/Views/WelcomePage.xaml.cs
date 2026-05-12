using System.Windows;
using System.Windows.Controls;

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
        // UI-57: Load statistics
        try
        {
            var groups = App.DataService.LoadRecipientGroups();
            var groupCount = groups.Count;
            var recipientCount = groups.SelectMany(g => g.Recipients).Count(r => r.IsSelected);

            if (groupCount > 0 || recipientCount > 0)
            {
                TxtGroupCount.Text = groupCount.ToString();
                TxtRecipientCount.Text = recipientCount.ToString();
                StatsPanel.Visibility = Visibility.Visible;
            }
        }
        catch { }

        // UI-58: Show version
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
}