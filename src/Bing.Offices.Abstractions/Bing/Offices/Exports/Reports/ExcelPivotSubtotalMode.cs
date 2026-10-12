namespace Bing.Offices.Exports;

/// <summary>
/// 指定透视表轴字段的小计方式。
/// </summary>
public enum ExcelPivotSubtotalMode
{
    /// <summary>
    /// 不显示该字段的小计。
    /// </summary>
    None = 0,

    /// <summary>
    /// 使用该字段的默认自动小计。
    /// </summary>
    Automatic = 1
}
