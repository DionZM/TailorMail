using System.Globalization;
using System.Windows;
using System.Windows.Data;
using TailorMail.Models;

namespace TailorMail.Converters;

/// <summary>
/// 枚举值与布尔值的互转转换器，用于 RadioButton 绑定枚举类型。
/// ConverterParameter 指定枚举值的字符串表示。
/// </summary>
public class EnumToBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is Enum enumValue && parameter is string paramStr)
        {
            var paramValue = Enum.Parse(enumValue.GetType(), paramStr);
            return enumValue.Equals(paramValue);
        }
        return false;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is true && parameter is string paramStr && targetType.IsEnum)
        {
            return Enum.Parse(targetType, paramStr);
        }
        return System.Windows.Data.Binding.DoNothing;
    }
}

/// <summary>
/// 布尔值取反转换器。true → false，false → true。
/// </summary>
public class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && !b;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && !b;
    }
}

/// <summary>
/// 空值/空字符串转 Collapsed 转换器。值为空时隐藏元素，非空时显示。
/// </summary>
public class NullToCollapsedConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return string.IsNullOrEmpty(value?.ToString()) ? Visibility.Collapsed : Visibility.Visible;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 空值/空字符串转 Visible 转换器。值为空时显示元素，非空时隐藏（与 NullToCollapsedConverter 相反）。
/// </summary>
public class NullToVisibleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return !string.IsNullOrEmpty(value?.ToString()) ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 整数转可见性转换器。整数大于 0 时显示元素，否则隐藏。
/// 支持 "invert" 参数反转逻辑。
/// </summary>
public class IntToVisibleConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var visible = value is int i && i > 0;
        if (parameter is string s && s.Equals("invert", StringComparison.OrdinalIgnoreCase))
            visible = !visible;
        return visible ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 布尔值取反后转可见性转换器。true → Collapsed，false → Visible。
/// </summary>
public class InverseBoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is bool b && !b ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 文件路径转文件名转换器。从完整路径中提取文件名部分用于显示。
/// </summary>
public class FileNameConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is string path && !string.IsNullOrEmpty(path))
            return System.IO.Path.GetFileName(path);
        return value;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 文件路径转文件大小字符串转换器。从完整路径获取文件大小并转换为人类可读格式。
/// </summary>
public class FilePathSizeConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is not string path || string.IsNullOrEmpty(path))
            return "";
        try
        {
            var info = new System.IO.FileInfo(path);
            if (!info.Exists) return "";
            string[] suffixes = { "B", "KB", "MB", "GB" };
            var order = 0;
            double size = info.Length;
            while (size >= 1024 && order < suffixes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            return order == 0 ? $"{info.Length} {suffixes[order]}" : $"{size:0.#} {suffixes[order]}";
        }
        catch
        {
            return "";
        }
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}

/// <summary>
/// 文件大小转人类可读字符串转换器。将字节数（long）转换为如 "2.5 MB" 的格式。
/// </summary>
public class FileSizeConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is long bytes)
            return FormatSize(bytes);
        if (value is int intBytes && intBytes >= 0)
            return FormatSize(intBytes);
        return "";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }

    private static string FormatSize(long bytes)
    {
        string[] suffixes = { "B", "KB", "MB", "GB" };
        var order = 0;
        double size = bytes;
        while (size >= 1024 && order < suffixes.Length - 1)
        {
            order++;
            size /= 1024;
        }
        return order == 0 ? $"{bytes} {suffixes[order]}" : $"{size:0.#} {suffixes[order]}";
    }
}

/// <summary>
/// 发送失败状态转可见性转换器。当 <see cref="SendStatus"/> 为 Failed 时显示元素。
/// </summary>
public class FailedToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is SendStatus status && status == SendStatus.Failed
            ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
