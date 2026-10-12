using System;
using System.Collections.Generic;

namespace Bing.Offices.Exports;

/// <summary>
/// 定义透视表的一个报表筛选字段及其选项。
/// </summary>
/// <remarks>
/// 筛选只影响透视报表的选择状态，不会从来源区域或透视缓存中删除记录。
/// </remarks>
public sealed class ExcelPivotFilterDefinition
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotFilterDefinition" /> 类型的实例。
    /// </summary>
    public ExcelPivotFilterDefinition()
    {
    }

    /// <summary>
    /// 获取或初始化来源区域中的表头名称。
    /// </summary>
    public string Field { get; init; }

    /// <summary>
    /// 获取或初始化按首次出现顺序排列的筛选字面值；空集合表示全选。
    /// </summary>
    /// <remarks>
    /// 支持字符串、布尔值、有限数值、日期时间和 null 空项。
    /// </remarks>
    public IReadOnlyList<object> Values { get; init; } = Array.Empty<object>();
}
