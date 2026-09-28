using System.Collections.ObjectModel;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

/// <summary>
/// 单个工作表名称及映射计划的不可变描述。
/// </summary>
internal sealed class ExcelMappingSheetPlan : IExcelMappingSheetPlan
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelMappingSheetPlan" /> 类型的实例。
    /// </summary>
    /// <param name="name">工作表名称。</param>
    /// <param name="mapping">工作表使用的列映射计划。</param>
    internal ExcelMappingSheetPlan(string name, IExcelMappingPlan mapping)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(name));
        Name = name;
        Mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public IExcelMappingPlan Mapping { get; }
}
