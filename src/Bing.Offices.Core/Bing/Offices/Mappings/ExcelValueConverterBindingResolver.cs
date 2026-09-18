using System.Collections.Concurrent;
using System.Runtime.CompilerServices;
using Bing.Offices.Conversions;

namespace Bing.Offices.Mappings;

/// <summary>
/// 统一解析并缓存值转换器能力。
/// </summary>
internal static class ExcelValueConverterBindingResolver
{
    /// <summary>
    /// 按转换器实例缓存其对属性类型的支持能力。
    /// </summary>
    private static readonly ConditionalWeakTable<IExcelValueConverter,
        ConcurrentDictionary<Type, bool>> Capabilities = new();

    /// <summary>
    /// 解析指定名称和类型的转换器。
    /// </summary>
    /// <param name="converters">可供解析的值转换器集合。</param>
    /// <param name="converterName">可选的命名转换器名称。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <returns>按名称或能力匹配的转换器集合。</returns>
    public static IReadOnlyList<IExcelValueConverter> Resolve(IEnumerable<IExcelValueConverter> converters,
        string converterName, Type propertyType)
    {
        if (propertyType == null)
            throw new ArgumentNullException(nameof(propertyType));
        var items = (converters ?? Array.Empty<IExcelValueConverter>()).ToArray();
        if (string.IsNullOrWhiteSpace(converterName))
            return items.Where(converter => CanConvert(converter, propertyType)).ToArray();
        var named = items.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, converterName, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (named.Length != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {converterName}");
        if (!CanConvert(named[0], propertyType))
            throw new InvalidOperationException($"值转换器 {converterName} 不支持属性类型: {propertyType.FullName}");
        return named;
    }

    /// <summary>
    /// 判断并缓存转换器对指定属性类型的支持能力。
    /// </summary>
    /// <param name="converter">待检查的值转换器。</param>
    /// <param name="propertyType">目标属性类型。</param>
    /// <returns>转换器支持该属性类型时为 true。</returns>
    private static bool CanConvert(IExcelValueConverter converter, Type propertyType)
    {
        var cache = Capabilities.GetValue(converter, _ => new ConcurrentDictionary<Type, bool>());
        return cache.GetOrAdd(propertyType, type => converter.CanConvert(type));
    }
}
