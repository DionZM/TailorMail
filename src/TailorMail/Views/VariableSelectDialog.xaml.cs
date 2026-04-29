﻿using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TailorMail.ViewModels;

namespace TailorMail.Views;

public partial class VariableSelectDialog : Wpf.Ui.Controls.FluentWindow
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

        var panel = new StackPanel { Margin = new Thickness(20) };

        var hint = new TextBlock
        {
            Text = "请勾选要删除的变量：",
            Style = (Style)FindResource("SectionTitleStyle"),
            Margin = new Thickness(0, 0, 0, 12)
        };
        panel.Children.Add(hint);

        var scrollViewer = new ScrollViewer { VerticalScrollBarVisibility = ScrollBarVisibility.Auto, MaxHeight = 250 };
        var itemsPanel = new StackPanel();
        var checkBoxes = new List<CheckBox>();

        foreach (var name in variableNames)
        {
            var cb = new CheckBox
            {
                Content = name,
                Margin = new Thickness(0, 4, 0, 4),
                FontFamily = (FontFamily)FindResource("AppFontFamily"),
                Tag = name
            };
            checkBoxes.Add(cb);
            itemsPanel.Children.Add(cb);
        }
        scrollViewer.Content = itemsPanel;
        panel.Children.Add(scrollViewer);

        var btnPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 16, 0, 0)
        };

        var btnOk = new Wpf.Ui.Controls.Button
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

        var btnCancel = new Wpf.Ui.Controls.Button
        {
            Content = "取消",
            Padding = new Thickness(20, 6, 20, 6),
            Appearance = Wpf.Ui.Controls.ControlAppearance.Secondary
        };
        btnCancel.Click += (_, _) => { DialogResult = false; };

        btnPanel.Children.Add(btnOk);
        btnPanel.Children.Add(btnCancel);
        panel.Children.Add(btnPanel);

        Content = panel;
    }
}
