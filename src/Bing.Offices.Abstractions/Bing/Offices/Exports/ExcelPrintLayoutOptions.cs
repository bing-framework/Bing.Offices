using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 工作表打印布局选项。
/// </summary>
public sealed class ExcelPrintLayoutOptions
{
    /// <summary>
    /// 获取或初始化纸张大小。
    /// </summary>
    public ExcelPrintPaperSize PaperSize { get; init; } = ExcelPrintPaperSize.A4;
    /// <summary>
    /// 获取或初始化页面方向。
    /// </summary>
    public ExcelPrintOrientation Orientation { get; init; } = ExcelPrintOrientation.Portrait;
    /// <summary>
    /// 获取或初始化上边距，单位为英寸。
    /// </summary>
    public double TopMargin { get; init; } = 0.75;
    /// <summary>
    /// 获取或初始化下边距，单位为英寸。
    /// </summary>
    public double BottomMargin { get; init; } = 0.75;
    /// <summary>
    /// 获取或初始化左边距，单位为英寸。
    /// </summary>
    public double LeftMargin { get; init; } = 0.7;
    /// <summary>
    /// 获取或初始化右边距，单位为英寸。
    /// </summary>
    public double RightMargin { get; init; } = 0.7;
    /// <summary>
    /// 获取或初始化缩放百分比。
    /// </summary>
    public short? ScalePercent { get; init; }
    /// <summary>
    /// 获取或初始化适应页宽。
    /// </summary>
    public short? FitToWidth { get; init; }
    /// <summary>
    /// 获取或初始化适应页高。
    /// </summary>
    public short? FitToHeight { get; init; }
    /// <summary>
    /// 获取或初始化打印区域，A1 表达式。
    /// </summary>
    public string PrintArea { get; init; }
    /// <summary>
    /// 获取或初始化重复打印标题行，A1 表达式。
    /// </summary>
    public string RepeatRows { get; init; }
    /// <summary>
    /// 获取或初始化重复打印标题列，A1 表达式。
    /// </summary>
    public string RepeatColumns { get; init; }
    /// <summary>
    /// 获取或初始化页眉文本。
    /// </summary>
    public string Header { get; init; }
    /// <summary>
    /// 获取或初始化页脚文本。
    /// </summary>
    public string Footer { get; init; }

    /// <summary>
    /// 验证打印布局。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (!Enum.IsDefined(typeof(ExcelPrintPaperSize), PaperSize)
            || !Enum.IsDefined(typeof(ExcelPrintOrientation), Orientation))
            throw new ArgumentOutOfRangeException(parameterName ?? nameof(PaperSize));
        ValidateMargin(TopMargin, nameof(TopMargin));
        ValidateMargin(BottomMargin, nameof(BottomMargin));
        ValidateMargin(LeftMargin, nameof(LeftMargin));
        ValidateMargin(RightMargin, nameof(RightMargin));
        if (ScalePercent.HasValue && (ScalePercent < 1 || ScalePercent > 400))
            throw new ArgumentOutOfRangeException(nameof(ScalePercent));
        if (FitToWidth.HasValue && FitToWidth < 1 || FitToHeight.HasValue && FitToHeight < 1)
            throw new ArgumentOutOfRangeException(nameof(FitToWidth));
    }

    /// <summary>
    /// 验证打印页边距的有效范围。
    /// </summary>
    /// <param name="value">页边距，单位为英寸，允许范围为 0 至 10。</param>
    /// <param name="name">无效值对应的参数名称。</param>
    private static void ValidateMargin(double value, string name)
    {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > 10)
            throw new ArgumentOutOfRangeException(name);
    }
}
