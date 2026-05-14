using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;

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
        Loaded += OnLoaded;
        KeyDown += OnDialogKeyDown;
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        InputTextBox.Focus();
        PlayOpenAnimation();
    }

    private void OnDialogKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape)
        {
            e.Handled = true;
            CloseWithResult(false);
        }
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
        CloseWithResult(true);
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
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
