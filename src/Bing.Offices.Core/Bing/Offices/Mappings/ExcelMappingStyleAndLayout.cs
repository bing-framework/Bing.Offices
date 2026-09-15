using Bing.Offices.Configurations;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

/// <summary>编译后的映射样式配置。</summary>
internal sealed class ExcelMappingStyle : IExcelMappingStyle
{
    /// <summary>初始化一个 <see cref="ExcelMappingStyle" /> 类型的实例。</summary>
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
/// <summary>编译后的映射布局配置。</summary>
internal sealed class ExcelMappingLayout : IExcelMappingLayout
{
    /// <summary>初始化一个 <see cref="ExcelMappingLayout" /> 类型的实例。</summary>
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
