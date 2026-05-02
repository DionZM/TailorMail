using System.Windows;
using System.Windows.Media;

namespace TailorMail;

public static class DesignTokens
{
    public const string AccentColor = "#1D6FB5";
    public const string AccentHoverColor = "#185A94";
    public const string AccentLightColor = "#E0F0FF";
    public const string SurfaceColor = "#F2F5F9";
    public const string SurfaceElevatedColor = "#FFFFFF";
    public const string BorderSubtleColor = "#DDE2E9";
    public const string TextPrimaryColor = "#1A2332";
    public const string TextSecondaryColor = "#5A6A7E";
    public const string TextTertiaryColor = "#626F82";
    public const string SuccessColor = "#047857";
    public const string WarningColor = "#A16207";
    public const string DangerColor = "#C42B1C";
    public const string RowHoverColor = "#EDF2F8";
    public const string RowSelectedColor = "#D6E9FB";

    public const string DisplayFontFamily = "Georgia, Segoe UI Variable, Segoe UI, Microsoft YaHei UI";
    public const string BodyFontFamily = "Segoe UI Variable, Segoe UI, Microsoft YaHei UI";
    public const string UIFontFamily = "Segoe UI Variable, Segoe UI, Microsoft YaHei UI";
    public const string MonoFontFamily = "Cascadia Code, Consolas";

    public const double SpacingXS = 4;
    public const double SpacingSM = 8;
    public const double SpacingMD = 16;
    public const double SpacingLG = 24;
    public const double SpacingXL = 32;
    public const double Spacing2XL = 48;
    public const double Spacing3XL = 64;
    public const double Spacing4XL = 96;

    public const double FontSizeXS = 10;
    public const double FontSizeSM = 13;
    public const double FontSizeBase = 16;
    public const double FontSizeMD = 20;
    public const double FontSizeLG = 25;
    public const double FontSizeXL = 31;
    public const double FontSize2XL = 39;
    public const double FontSize3XL = 49;

    public static readonly CornerRadius CornerRadiusXS = new(2);
    public static readonly CornerRadius CornerRadiusSM = new(4);
    public static readonly CornerRadius CornerRadiusMD = new(8);
    public static readonly CornerRadius CornerRadiusLG = new(12);
    public static readonly CornerRadius CornerRadiusXL = new(16);

    public const double DurationFast = 150;
    public const double DurationNormal = 300;
    public const double DurationSlow = 500;

    public const double AnimationSlideOffset = 12;
    public const double AnimationDuration = 250;
}
