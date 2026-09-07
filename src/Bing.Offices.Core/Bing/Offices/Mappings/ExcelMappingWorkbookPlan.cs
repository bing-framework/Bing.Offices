using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Collections.ObjectModel;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Validations;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

internal sealed class ExcelMappingWorkbookPlan : IExcelMappingWorkbookPlan
{
    /// <summary>从工作表映射计划创建不可变 Workbook 计划。</summary>
    /// <param name="sheets">按请求顺序排列的工作表计划。</param>
    internal ExcelMappingWorkbookPlan(IReadOnlyList<IExcelMappingSheetPlan> sheets)
    {
        Sheets = new ReadOnlyCollection<IExcelMappingSheetPlan>(sheets.ToArray());
    }

    /// <inheritdoc />
    public IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}
internal sealed class ExcelMappingSheetPlan : IExcelMappingSheetPlan
{
    /// <summary>为指定工作表名称和列映射创建计划。</summary>
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
