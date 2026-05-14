using System.Windows;
using System.Windows.Media.Animation;

namespace TailorMail.Helpers;

public static class AnimationHelper
{
    public static void FadeIn(FrameworkElement element, double durationMs = 200)
    {
        element.Opacity = 0;
        var anim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(durationMs))
        {
            EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
        };
        element.BeginAnimation(UIElement.OpacityProperty, anim);
    }

    public static void StaggerFadeIn(System.Windows.Controls.ListBox listBox, int delayMs = 40, int maxItems = 10)
    {
        var count = Math.Min(listBox.Items.Count, maxItems);
        for (int i = 0; i < count; i++)
        {
            if (listBox.ItemContainerGenerator.ContainerFromIndex(i) is System.Windows.Controls.ListBoxItem item)
            {
                item.Opacity = 0;
                var anim = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200))
                {
                    BeginTime = TimeSpan.FromMilliseconds(i * delayMs),
                    EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut }
                };
                item.BeginAnimation(UIElement.OpacityProperty, anim);
            }
        }
    }
}
