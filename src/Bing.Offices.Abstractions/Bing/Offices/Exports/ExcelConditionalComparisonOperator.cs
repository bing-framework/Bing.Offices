using System.ComponentModel;
using System.Text.RegularExpressions;

namespace Bing.Offices.Exports;

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
