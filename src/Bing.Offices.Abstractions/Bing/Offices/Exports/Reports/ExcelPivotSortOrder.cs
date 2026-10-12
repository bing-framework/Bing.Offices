namespace Bing.Offices.Exports;

/// <summary>
/// 指定透视表轴标签的排序方向。
/// </summary>
public enum ExcelPivotSortOrder
{
    /// <summary>
    /// 使用目标工作簿引擎的默认排序。
    /// </summary>
    Default = 0,

    /// <summary>
    /// 按标签升序排列。
    /// </summary>
    Ascending = 1,

    /// <summary>
    /// 按标签降序排列。
    /// </summary>
    Descending = 2
}
