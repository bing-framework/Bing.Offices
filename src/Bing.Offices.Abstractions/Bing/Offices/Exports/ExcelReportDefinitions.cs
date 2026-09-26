using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

/// <summary>
/// 条件格式类型。
/// </summary>
public enum ExcelConditionalFormatType
{
    /// <summary>
    /// 单元格值比较。
    /// </summary>
    CellValue,
    /// <summary>
    /// 简单公式。
    /// </summary>
    Formula,
    /// <summary>
    /// 色阶。
    /// </summary>
    ColorScale,
    /// <summary>
    /// 数据条。
    /// </summary>
    DataBar,
    /// <summary>
    /// 图标集。
    /// </summary>
    IconSet
}

/// <summary>
/// 条件格式比较运算符。
/// </summary>
public enum ExcelConditionalComparisonOperator
{
    /// <summary>
    /// 等于。
    /// </summary>
    Equal,
    /// <summary>
    /// 不等于。
    /// </summary>
    NotEqual,
    /// <summary>
    /// 大于。
    /// </summary>
    GreaterThan,
    /// <summary>
    /// 大于或等于。
    /// </summary>
    GreaterThanOrEqual,
    /// <summary>
    /// 小于。
    /// </summary>
    LessThan,
    /// <summary>
    /// 小于或等于。
    /// </summary>
    LessThanOrEqual,
    /// <summary>
    /// 位于两个边界之间。
    /// </summary>
    Between,
    /// <summary>
    /// 不位于两个边界之间。
    /// </summary>
    NotBetween
}

/// <summary>
/// 工作表区域定义。
/// </summary>
public sealed class ExcelRangeDefinition
{
    /// <summary>
    /// 获取或初始化零基起始行，包含。
    /// </summary>
    public int StartRow { get; init; }
    /// <summary>
    /// 获取或初始化零基起始列，包含。
    /// </summary>
    public int StartColumn { get; init; }
    /// <summary>
    /// 获取或初始化零基结束行，包含。
    /// </summary>
    public int EndRow { get; init; }
    /// <summary>
    /// 获取或初始化零基结束列，包含。
    /// </summary>
    public int EndColumn { get; init; }

    /// <summary>
    /// 验证区域边界。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称。</param>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public void Validate(string parameterName)
    {
        if (StartRow < 0 || StartColumn < 0 || EndRow < StartRow || EndColumn < StartColumn)
            throw new ArgumentException("区域必须是非空的正向区域。", parameterName);
    }
}

/// <summary>
/// Excel 表格定义。
/// </summary>
public sealed class ExcelTableDefinition
{
    /// <summary>
    /// 获取或初始化表格稳定名称。
    /// </summary>
    public string Name { get; init; }
    /// <summary>
    /// 获取或初始化表格区域。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 获取或初始化是否包含表头。
    /// </summary>
    public bool HasHeaders { get; init; } = true;
    /// <summary>
    /// 获取或初始化是否显示汇总行。
    /// </summary>
    public bool ShowTotals { get; init; }
    /// <summary>
    /// 获取或初始化可选的内置表格样式名称。
    /// </summary>
    public string StyleName { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("表格必须指定名称。", parameterName ?? nameof(Name));
        Range?.Validate(parameterName ?? nameof(Range));
        if (Range == null)
            throw new ArgumentException("表格必须指定区域。", parameterName ?? nameof(Range));
    }
}

/// <summary>
/// 自动筛选区域定义。
/// </summary>
public sealed class ExcelAutoFilterDefinition
{
    /// <summary>
    /// 获取或初始化筛选区域。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (Range == null)
            throw new ArgumentException("自动筛选必须指定区域。", parameterName ?? nameof(Range));
        Range.Validate(parameterName ?? nameof(Range));
    }
}

/// <summary>
/// 冻结窗格定义。
/// </summary>
public sealed class ExcelFreezePaneDefinition
{
    /// <summary>
    /// 获取或初始化冻结的行数。
    /// </summary>
    public int Rows { get; init; }
    /// <summary>
    /// 获取或初始化冻结的列数。
    /// </summary>
    public int Columns { get; init; }
    /// <summary>
    /// 获取或初始化可选的顶部可视行索引。
    /// </summary>
    public int? TopRow { get; init; }
    /// <summary>
    /// 获取或初始化可选的左侧可视列索引。
    /// </summary>
    public int? LeftColumn { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (Rows < 0 || Columns < 0 || (Rows == 0 && Columns == 0))
            throw new ArgumentException("冻结窗格至少需要冻结一行或一列。", parameterName ?? nameof(Rows));
        if (TopRow.HasValue && TopRow.Value < 0 || LeftColumn.HasValue && LeftColumn.Value < 0)
            throw new ArgumentException("冻结窗格可视起点不能为负数。", parameterName ?? nameof(TopRow));
    }
}

/// <summary>
/// 条件格式定义。
/// </summary>
public sealed class ExcelConditionalFormatDefinition
{
    /// <summary>
    /// 获取或初始化条件格式类型。
    /// </summary>
    public ExcelConditionalFormatType Type { get; init; }
    /// <summary>
    /// 获取或初始化应用区域。
    /// </summary>
    public ExcelRangeDefinition Range { get; init; }
    /// <summary>
    /// 获取或初始化比较运算符。
    /// </summary>
    public ExcelConditionalComparisonOperator Operator { get; init; }
    /// <summary>
    /// 获取或初始化第一个比较值或公式。
    /// </summary>
    public string Formula1 { get; init; }
    /// <summary>
    /// 获取或初始化Between/NotBetween 的第二个比较值或公式。
    /// </summary>
    public string Formula2 { get; init; }
    /// <summary>
    /// 获取或初始化格式前景色（ARGB 或 CSS 十六进制）。
    /// </summary>
    public string ForegroundColor { get; init; }
    /// <summary>
    /// 获取或初始化格式背景色（ARGB 或 CSS 十六进制）。
    /// </summary>
    public string BackgroundColor { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        var name = parameterName ?? nameof(Range);
        if (Range == null)
            throw new ArgumentException("条件格式必须指定区域。", name);
        Range.Validate(name);
        if (!Enum.IsDefined(typeof(ExcelConditionalFormatType), Type))
            throw new ArgumentOutOfRangeException(nameof(Type));
        if ((Type == ExcelConditionalFormatType.CellValue || Type == ExcelConditionalFormatType.Formula)
            && string.IsNullOrWhiteSpace(Formula1))
            throw new ArgumentException("值比较和公式条件格式必须指定 Formula1。", nameof(Formula1));
        if (Type == ExcelConditionalFormatType.CellValue
            && (Operator == ExcelConditionalComparisonOperator.Between
                || Operator == ExcelConditionalComparisonOperator.NotBetween)
            && string.IsNullOrWhiteSpace(Formula2))
            throw new ArgumentException("Between 和 NotBetween 必须指定 Formula2。", nameof(Formula2));
    }
}

/// <summary>
/// 名称范围定义。
/// </summary>
public sealed class ExcelNamedRangeDefinition
{
    /// <summary>
    /// 获取或初始化名称范围名称。
    /// </summary>
    public string Name { get; init; }
    /// <summary>
    /// 获取或初始化可选的 Sheet 作用域；为空表示 Workbook 作用域。
    /// </summary>
    public string SheetName { get; init; }
    /// <summary>
    /// 获取或初始化有界区域的 A1 引用。
    /// </summary>
    public string Address { get; init; }

    /// <summary>
    /// 验证定义。
    /// </summary>
    /// <param name="parameterName">验证失败时使用的参数名称；未指定时使用对应属性名称。</param>
    public void Validate(string parameterName = null)
    {
        if (string.IsNullOrWhiteSpace(Name))
            throw new ArgumentException("名称范围必须指定名称。", parameterName ?? nameof(Name));
        if (string.IsNullOrWhiteSpace(Address)
            || Address.IndexOf('[') >= 0
            || Address.IndexOf(']') >= 0
            || !Regex.IsMatch(Address, @"^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$",
                RegexOptions.CultureInvariant))
            throw new ArgumentException("名称范围必须指定有界区域引用。", parameterName ?? nameof(Address));
    }
}

/// <summary>
/// 打印方向。
/// </summary>
public enum ExcelPrintOrientation
{
    /// <summary>
    /// 纵向。
    /// </summary>
    Portrait,
    /// <summary>
    /// 横向。
    /// </summary>
    Landscape
}

/// <summary>
/// 常用打印纸张。
/// </summary>
public enum ExcelPrintPaperSize
{
    /// <summary>
    /// Letter。
    /// </summary>
    Letter,
    /// <summary>
    /// A4。
    /// </summary>
    A4,
    /// <summary>
    /// A3。
    /// </summary>
    A3,
    /// <summary>
    /// Legal。
    /// </summary>
    Legal
}

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
