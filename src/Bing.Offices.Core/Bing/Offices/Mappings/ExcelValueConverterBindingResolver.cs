using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using Bing.Offices.Conversions;

namespace Bing.Offices.Mappings;

/// <summary>
/// 统一解析并缓存值转换器能力。
/// </summary>
internal static class ExcelValueConverterBindingResolver
{
    private static readonly ConditionalWeakTable<IExcelValueConverter,
        ConcurrentDictionary<Type, bool>> Capabilities = new();

    /// <summary>
    /// 解析指定名称和类型的转换器。
    /// </summary>
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

    private static bool CanConvert(IExcelValueConverter converter, Type propertyType)
    {
        var cache = Capabilities.GetValue(converter, _ => new ConcurrentDictionary<Type, bool>());
        return cache.GetOrAdd(propertyType, type => converter.CanConvert(type));
    }
}
