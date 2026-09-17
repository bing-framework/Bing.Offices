using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// 为 MiniExcel Provider 创建并缓存 Core 映射计划。
/// </summary>
internal sealed class MiniExcelMappingPlanBuilder
{
    private delegate IExcelMappingPlan CreatePlanInvoker(MiniExcelMappingPlanBuilder target,
        ExcelMappingDocument document, ExcelMappingConfiguration configuration,
        MappingDirection direction);

    private static readonly ConcurrentDictionary<Type, CreatePlanInvoker> Invokers = new();
    private readonly IExcelMappingPlanFactory _factory;

    public MiniExcelMappingPlanBuilder(IExcelMappingPlanFactory factory)
    {
        _factory = factory ?? throw new ArgumentNullException(nameof(factory));
    }

    public IExcelMappingPlan Create(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        if (itemType == null)
            throw new ArgumentNullException(nameof(itemType));
        return Invokers.GetOrAdd(itemType, CreateInvoker)(this, document, configuration, direction);
    }

    private static CreatePlanInvoker CreateInvoker(Type itemType)
    {
        var method = typeof(MiniExcelMappingPlanBuilder).GetMethod(nameof(CreateTyped),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(itemType);
        return (CreatePlanInvoker)method.CreateDelegate(typeof(CreatePlanInvoker));
    }

    private IExcelMappingPlan CreateTyped<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
        where T : class, new()
    {
        return _factory.Create<T>(document ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
    }

    internal static ExcelMappingConfiguration MergeRequestDynamicColumns(
        ExcelMappingConfiguration configuration, IReadOnlyList<ExcelDynamicColumnDefinition> definitions)
    {
        if (definitions == null || definitions.Count == 0)
            return configuration;
        var overlay = new ExcelMappingConfiguration
        {
            DynamicColumns = definitions.Select(definition => new ExcelMappingDynamicColumnConfiguration
            {
                Key = definition.Key,
                Title = definition.Title,
                Aliases = (definition.Aliases ?? Array.Empty<string>()).ToList(),
                DataTypeName = GetDataTypeName(definition.DataType),
                Order = definition.Order,
                ConverterName = definition.ConverterName,
                ValidatorName = definition.ValidatorName,
                ValidationRuleNames = (definition.ValidationRuleNames ?? Array.Empty<string>()).ToList(),
                NumberFormat = definition.NumberFormat,
                ColumnIndex = definition.PhysicalColumnIndex ?? definition.Placement?.PhysicalColumnIndex,
                PlacementKey = GetPlacementKey(definition.Placement),
                ImageMultiplicity = definition.ImageMultiplicity
            }).ToList()
        };
        return MappingConfigurationMerger.Merge(configuration, overlay, MappingSourceKind.Request);
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
