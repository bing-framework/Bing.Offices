using System.Collections.Generic;
using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的不可变 Workbook 映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingWorkbookPlan
{
    /// <summary>获取 Sheet 映射计划。</summary>
    IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}

/// <summary>
/// Provider 使用的不可变 Sheet 映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingSheetPlan
{
    /// <summary>获取 Sheet 名称。</summary>
    string Name { get; }

    /// <summary>获取 Sheet 的列映射计划。</summary>
    IExcelMappingPlan Mapping { get; }
}
