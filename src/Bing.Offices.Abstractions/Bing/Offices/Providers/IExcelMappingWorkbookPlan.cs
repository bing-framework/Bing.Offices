using System.ComponentModel;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的不可变 Workbook 映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingWorkbookPlan
{
    /// <summary>
    /// 获取按请求 Sheet 顺序排列的映射计划。
    /// </summary>
    IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}
