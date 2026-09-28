using System.Collections.ObjectModel;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

/// <summary>
/// 多个工作表映射计划的不可变集合。
/// </summary>
internal sealed class ExcelMappingWorkbookPlan : IExcelMappingWorkbookPlan
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelMappingWorkbookPlan" /> 类型的实例。
    /// </summary>
    /// <param name="sheets">按请求顺序排列的工作表计划。</param>
    internal ExcelMappingWorkbookPlan(IReadOnlyList<IExcelMappingSheetPlan> sheets)
    {
        Sheets = new ReadOnlyCollection<IExcelMappingSheetPlan>(sheets.ToArray());
    }

    /// <inheritdoc />
    public IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}
