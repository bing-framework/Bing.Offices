namespace Bing.Offices.Exports;

/// <summary>
/// 定义透视表中的一个值字段。
/// </summary>
/// <remarks>
/// 同一来源字段可声明多个度量，但每个显示标题必须唯一。
/// </remarks>
public sealed class ExcelPivotValueDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotValueDefinition" /> 类型的实例。
    /// </summary>
    public ExcelPivotValueDefinition()
    {
    }

    /// <summary>
    /// 获取或初始化来源区域中的表头名称。
    /// </summary>
    public string Field { get; init; }

    /// <summary>
    /// 获取或初始化透视表中显示的度量标题。
    /// </summary>
    public string Caption { get; init; }

    /// <summary>
    /// 获取或初始化度量的聚合方式。
    /// </summary>
    public ExcelPivotAggregation Aggregation { get; init; } = ExcelPivotAggregation.Sum;

    /// <summary>
    /// 获取或初始化 Excel 数字格式；未设置时使用引擎默认格式。
    /// </summary>
    public string NumberFormat { get; init; }
}
