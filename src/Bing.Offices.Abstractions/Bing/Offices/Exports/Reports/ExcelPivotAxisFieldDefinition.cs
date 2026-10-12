namespace Bing.Offices.Exports;

/// <summary>
/// 定义透视表的一个行轴或列轴字段。
/// </summary>
/// <remarks>
/// 字段顺序决定轴层级；排序作用于标签，小计方式由目标工作簿引擎解释。
/// </remarks>
public sealed class ExcelPivotAxisFieldDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotAxisFieldDefinition" /> 类型的实例。
    /// </summary>
    public ExcelPivotAxisFieldDefinition()
    {
    }

    /// <summary>
    /// 获取或初始化来源区域中的表头名称。
    /// </summary>
    public string Field { get; init; }

    /// <summary>
    /// 获取或初始化轴标签的排序方式。
    /// </summary>
    public ExcelPivotSortOrder SortOrder { get; init; } = ExcelPivotSortOrder.Default;

    /// <summary>
    /// 获取或初始化该轴字段的小计方式。
    /// </summary>
    public ExcelPivotSubtotalMode SubtotalMode { get; init; } = ExcelPivotSubtotalMode.Automatic;
}
