using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace TailorMail.Views;

public partial class ConfirmDialog
{
    public string Message
    {
        get => MessageText.Text;
        set => MessageText.Text = value;
    }

    public bool IsDangerConfirmation { get; set; }

    public ConfirmDialog()
    {
        InitializeComponent();
        Owner = Application.Current.MainWindow;
        Loaded += OnLoaded;
        KeyDown += OnKeyDown;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (IsDangerConfirmation)
        {
            BtnYes.Appearance = Wpf.Ui.Controls.ControlAppearance.Danger;
        }
        BtnYes.Focus();
        PlayOpenAnimation();
    }

    private void OnKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            CloseWithResult(false);
        }
    }

    private void BtnYes_Click(object sender, RoutedEventArgs e)
    {
        CloseWithResult(true);
    }

    private void BtnNo_Click(object sender, RoutedEventArgs e)
    {
        CloseWithResult(false);
    }

    private void PlayOpenAnimation()
    {
        DialogScale.ScaleX = 0.95;
        DialogScale.ScaleY = 0.95;
        Opacity = 0;

        var scaleAnim = new DoubleAnimation(0.95, 1.0, TimeSpan.FromMilliseconds(200))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        var opacityAnim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };

        DialogScale.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
        DialogScale.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
        BeginAnimation(OpacityProperty, opacityAnim);
    }

    private void CloseWithResult(bool result)
    {
        var scaleAnim = new DoubleAnimation(1.0, 0.95, TimeSpan.FromMilliseconds(120))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
        };
        var opacityAnim = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(120))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn }
        };
        opacityAnim.Completed += (_, _) =>
        {
            DialogResult = result;
            Close();
        };
        DialogScale.BeginAnimation(ScaleTransform.ScaleXProperty, scaleAnim);
        DialogScale.BeginAnimation(ScaleTransform.ScaleYProperty, scaleAnim);
        BeginAnimation(OpacityProperty, opacityAnim);
    }
}
