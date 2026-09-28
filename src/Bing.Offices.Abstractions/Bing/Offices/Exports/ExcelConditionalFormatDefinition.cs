using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

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
