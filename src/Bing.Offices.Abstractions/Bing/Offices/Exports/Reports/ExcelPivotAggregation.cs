namespace Bing.Offices.Exports;

/// <summary>
/// 指定透视表值字段的聚合方式。
/// </summary>
public enum ExcelPivotAggregation
{
    /// <summary>
    /// 对数值求和。
    /// </summary>
    Sum = 0,

    /// <summary>
    /// 统计非空单元格。
    /// </summary>
    Count = 1,

    /// <summary>
    /// 计算非空数值的平均值。
    /// </summary>
    Average = 2,

    /// <summary>
    /// 取非空数值中的最小值。
    /// </summary>
    Minimum = 3,

    /// <summary>
    /// 取非空数值中的最大值。
    /// </summary>
    Maximum = 4
}
