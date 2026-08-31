using System.Collections.Generic;
using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的不可变 Workbook 映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingWorkbookPlan
{
    /// <summary>获取按请求 Sheet 顺序排列的映射计划。</summary>
    IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}

/// <summary>
/// Provider 使用的不可变 Sheet 映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingSheetPlan
{
    /// <summary>获取该计划适用的工作表名称。</summary>
    string Name { get; }

    /// <summary>获取工作表使用的不可变列映射计划。</summary>
    IExcelMappingPlan Mapping { get; }
}
