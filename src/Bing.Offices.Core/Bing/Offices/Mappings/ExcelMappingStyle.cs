using Bing.Offices.Configurations;
using Bing.Offices.Providers;

#nullable enable annotations

namespace Bing.Offices.Mappings;

/// <summary>
/// 编译后的映射样式配置。
/// </summary>
internal sealed class ExcelMappingStyle : IExcelMappingStyle
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelMappingStyle" /> 类型的实例。
    /// </summary>
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
