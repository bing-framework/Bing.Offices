using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

/// <summary>
/// 复制并合并请求级动态列定义的内部辅助类。
/// </summary>
internal static class ExcelDynamicColumnCloner
{
    /// <summary>
    /// 复制动态列定义集合。
    /// </summary>
    /// <param name="columns">待复制的动态列定义集合。</param>
    /// <returns>动态列定义的独立数组；输入为 <see langword="null" /> 时返回空数组。</returns>
    internal static IReadOnlyList<ExcelDynamicColumnDefinition> Clone(
        IReadOnlyList<ExcelDynamicColumnDefinition> columns) =>
        (columns ?? Array.Empty<ExcelDynamicColumnDefinition>()).Select(column => new ExcelDynamicColumnDefinition
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = (column.Aliases ?? Array.Empty<string>()).ToArray(),
            DataType = column.DataType,
            Order = column.Order,
            Placement = column.Placement,
            PhysicalColumnIndex = column.PhysicalColumnIndex,
            NumberFormat = column.NumberFormat,
            HeaderStyle = column.HeaderStyle,
            BodyStyle = column.BodyStyle,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = (column.ValidationRuleNames ?? Array.Empty<string>()).ToArray(),
            ImageMultiplicity = column.ImageMultiplicity
        }).ToArray();

    /// <summary>
    /// 将请求级动态列定义合并到映射配置。
    /// </summary>
    /// <param name="configuration">待合并的基础映射配置；为 <see langword="null" /> 时创建请求级配置。</param>
    /// <param name="columns">请求级动态列定义集合。</param>
    /// <returns>包含动态列定义的独立映射配置。</returns>
    internal static ExcelMappingConfiguration MergeIntoConfiguration(
        ExcelMappingConfiguration configuration, IReadOnlyList<ExcelDynamicColumnDefinition> columns)
    {
        var result = configuration == null
            ? new ExcelMappingConfiguration { SourceKind = MappingSourceKind.Request }
            : MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Request);
        if (columns == null || columns.Count == 0)
            return result;
        result.DynamicColumns = columns.Select(column => new ExcelMappingDynamicColumnConfiguration
        {
            Key = column.Key,
            Title = column.Title,
            Aliases = (column.Aliases ?? Array.Empty<string>()).ToList(),
            DataTypeName = GetDataTypeName(column.DataType),
            Order = column.Order,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = (column.ValidationRuleNames ?? Array.Empty<string>()).ToList(),
            NumberFormat = column.NumberFormat,
            ColumnIndex = column.PhysicalColumnIndex ?? column.Placement?.PhysicalColumnIndex,
            PlacementKey = GetPlacementKey(column.Placement),
            ImageMultiplicity = column.ImageMultiplicity
        }).ToList();
        return result;
    }

    /// <summary>
    /// 将列位置转换为稳定的相对布局键。
    /// </summary>
    /// <param name="placement">动态列位置定义；为 <see langword="null" /> 时返回 <see langword="null" />。</param>
    /// <returns>相对布局键；未指定相对位置时返回 <see langword="null" />。</returns>
    private static string GetPlacementKey(ExcelColumnPlacement placement)
    {
        if (!string.IsNullOrWhiteSpace(placement?.BeforeKey))
            return $"before:{placement.BeforeKey}";
        if (!string.IsNullOrWhiteSpace(placement?.AfterKey))
            return $"after:{placement.AfterKey}";
        return null;
    }

    /// <summary>
    /// 将 CLR 类型转换为动态列允许的数据类型名称。
    /// </summary>
    /// <param name="type">动态列的 CLR 类型。</param>
    /// <returns>映射文档使用的数据类型名称。</returns>
    private static string GetDataTypeName(Type type)
    {
        type = Nullable.GetUnderlyingType(type) ?? type ?? typeof(string);
        if (type == typeof(object)) return "object";
        if (type == typeof(string)) return "string";
        if (type == typeof(bool)) return "boolean";
        if (type == typeof(byte)) return "byte";
        if (type == typeof(short)) return "int16";
        if (type == typeof(int)) return "int32";
        if (type == typeof(long)) return "int64";
        if (type == typeof(float)) return "single";
        if (type == typeof(double)) return "double";
        if (type == typeof(decimal)) return "decimal";
        if (type == typeof(DateTime)) return "dateTime";
        if (type == typeof(DateTimeOffset)) return "dateTimeOffset";
        if (type == typeof(Guid)) return "guid";
        if (type == typeof(byte[])) return "bytes";
        throw new ArgumentException($"动态列数据类型不在允许列表中: {type.FullName}", nameof(type));
    }
}
