using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 使用 Core 映射工厂为运行时实体类型创建 ClosedXML 计划。
/// </summary>
internal sealed class ClosedXmlMappingPlanBuilder
{
    /// <summary>
    /// 以运行时类型创建映射计划的委托。
    /// </summary>
    /// <param name="target">映射计划构建器。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>创建的映射计划。</returns>
    private delegate IExcelMappingPlan CreatePlanInvoker(ClosedXmlMappingPlanBuilder target,
        ExcelMappingDocument document, ExcelMappingConfiguration configuration, MappingDirection direction);

    /// <summary>
    /// 按实体类型缓存的泛型计划创建委托。
    /// </summary>
    private static readonly ConcurrentDictionary<Type, CreatePlanInvoker> Invokers = new();

    /// <summary>
    /// Core 映射计划工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _factory;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlMappingPlanBuilder" /> 类型的实例。
    /// </summary>
    /// <param name="factory">用于创建公共映射计划的工厂。</param>
    public ClosedXmlMappingPlanBuilder(IExcelMappingPlanFactory factory)
    {
        _factory = factory ?? throw new ArgumentNullException(nameof(factory));
    }

    /// <summary>
    /// 为指定实体类型创建映射计划。
    /// </summary>
    /// <param name="itemType">实体类型。</param>
    /// <param name="document">映射文档；为空时使用约定回退。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>创建的映射计划。</returns>
    public IExcelMappingPlan Create(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        if (itemType == null)
            throw new ArgumentNullException(nameof(itemType));
        return Invokers.GetOrAdd(itemType, CreateInvoker)(this, document, configuration, direction);
    }

    /// <summary>
    /// 创建指定实体类型对应的泛型计划委托。
    /// </summary>
    /// <param name="itemType">实体类型。</param>
    /// <returns>绑定实体类型后的计划创建委托。</returns>
    private static CreatePlanInvoker CreateInvoker(Type itemType)
    {
        var method = typeof(ClosedXmlMappingPlanBuilder).GetMethod(nameof(CreateTyped),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(itemType);
        return (CreatePlanInvoker)method.CreateDelegate(typeof(CreatePlanInvoker));
    }

    /// <summary>
    /// 调用 Core 工厂创建指定实体类型的映射计划。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>创建的映射计划。</returns>
    private IExcelMappingPlan CreateTyped<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction) where T : class, new() =>
        _factory.Create<T>(document ?? new ExcelMappingDocument { UseConventionFallback = true }, configuration, direction);

    /// <summary>
    /// 将请求动态列覆盖合并到 Core 映射配置，保持属性/配置来源由 Core 决定。
    /// </summary>
    /// <param name="configuration">已有映射配置。</param>
    /// <param name="definitions">请求中的动态列定义。</param>
    /// <returns>合并后的映射配置；没有动态列定义时返回原配置实例。</returns>
    public static ExcelMappingConfiguration MergeRequestDynamicColumns(
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

    /// <summary>
    /// 将动态列相对位置转换为配置使用的键。
    /// </summary>
    /// <param name="placement">动态列位置描述。</param>
    /// <returns>位置键；未设置相对位置时返回 <see langword="null" />。</returns>
    private static string GetPlacementKey(ExcelColumnPlacement placement)
    {
        if (!string.IsNullOrWhiteSpace(placement?.BeforeKey))
            return $"before:{placement.BeforeKey}";
        if (!string.IsNullOrWhiteSpace(placement?.AfterKey))
            return $"after:{placement.AfterKey}";
        return null;
    }

    /// <summary>
    /// 将动态列数据类型转换为公共配置名称。
    /// </summary>
    /// <param name="type">动态列数据类型。</param>
    /// <returns>公共配置使用的数据类型名称。</returns>
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
