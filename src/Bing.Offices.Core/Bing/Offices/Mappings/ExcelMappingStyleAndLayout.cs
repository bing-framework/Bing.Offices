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

internal sealed class ExcelMappingStyle : IExcelMappingStyle
{
    /// <summary>从规范化样式配置创建不可变映射样式。</summary>
    /// <param name="style">样式配置；为 null 时创建空样式。</param>
    internal ExcelMappingStyle(ExcelMappingStyleConfiguration style)
    {
        HeaderStyleKey = style?.HeaderStyleKey;
        BodyStyleKey = style?.BodyStyleKey;
        NumberFormat = style?.NumberFormat;
    }
    /// <inheritdoc />
    public string HeaderStyleKey { get; }
    /// <inheritdoc />
    public string BodyStyleKey { get; }
    /// <inheritdoc />
    public string NumberFormat { get; }
}
internal sealed class ExcelMappingLayout : IExcelMappingLayout
{
    /// <summary>从规范化布局配置创建不可变映射布局。</summary>
    /// <param name="layout">布局配置；为 null 时创建空布局。</param>
    internal ExcelMappingLayout(ExcelMappingLayoutConfiguration layout)
    {
        ColumnIndex = layout?.ColumnIndex;
        PlacementKey = layout?.PlacementKey;
    }
    /// <inheritdoc />
    public int? ColumnIndex { get; }
    /// <inheritdoc />
    public string PlacementKey { get; }
}
