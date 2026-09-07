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

internal sealed class ExcelMappingPlan : IExcelMappingPlan
{
    /// <summary>从已编译的列、动态列、样式和布局创建不可变映射计划。</summary>
    /// <param name="columns">固定列映射。</param>
    /// <param name="dynamicColumns">动态列映射。</param>
    /// <param name="style">样式配置。</param>
    /// <param name="layout">布局配置。</param>
    /// <param name="profileName">来源 Profile 名称。</param>
    /// <param name="modelAlias">来源模型别名。</param>
    internal ExcelMappingPlan(IReadOnlyList<IExcelMappingColumn> columns,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns, IExcelMappingStyle style,
        IExcelMappingLayout layout, string profileName, string modelAlias)
    {
        Columns = new ReadOnlyCollection<IExcelMappingColumn>(columns.ToArray());
        DynamicColumns = new ReadOnlyCollection<IExcelDynamicMappingColumn>(dynamicColumns.ToArray());
        Style = style;
        Layout = layout;
        ProfileName = profileName;
        ModelAlias = modelAlias;
    }

    /// <inheritdoc />
    public string ProfileName { get; }
    /// <inheritdoc />
    public string ModelAlias { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelMappingColumn> Columns { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelDynamicMappingColumn> DynamicColumns { get; }
    /// <inheritdoc />
    public IExcelMappingStyle Style { get; }
    /// <inheritdoc />
    public IExcelMappingLayout Layout { get; }
}
