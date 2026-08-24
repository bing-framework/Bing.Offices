using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Configurations;

namespace Bing.Offices.Exports;

internal static class ExcelDynamicColumnCloner
{
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

    private static string GetPlacementKey(ExcelColumnPlacement placement)
    {
        if (!string.IsNullOrWhiteSpace(placement?.BeforeKey))
            return $"before:{placement.BeforeKey}";
        if (!string.IsNullOrWhiteSpace(placement?.AfterKey))
            return $"after:{placement.AfterKey}";
        return null;
    }

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
