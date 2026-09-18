using System.Collections.Concurrent;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// 为 MiniExcel Provider 创建并缓存映射计划。
/// </summary>
/// <remarks>
/// 计划由注入的工厂创建；本类型仅负责运行时实体类型的反射调度和请求动态列合并。
/// </remarks>
internal sealed class MiniExcelMappingPlanBuilder
{
    /// <summary>
    /// 表示按实体类型创建映射计划的反射委托。
    /// </summary>
    /// <param name="target">执行计划创建的映射计划构建器。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>指定实体类型和方向的映射计划。</returns>
    private delegate IExcelMappingPlan CreatePlanInvoker(MiniExcelMappingPlanBuilder target,
        ExcelMappingDocument document, ExcelMappingConfiguration configuration,
        MappingDirection direction);

    /// <summary>
    /// 按行实体运行时类型缓存映射计划反射委托。
    /// </summary>
    private static readonly ConcurrentDictionary<Type, CreatePlanInvoker> Invokers = new();
    /// <summary>
    /// 创建 Core 映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _factory;

    /// <summary>
    /// 初始化一个 <see cref="MiniExcelMappingPlanBuilder" /> 类型的实例。
    /// </summary>
    /// <param name="factory">创建映射计划的工厂。</param>
    public MiniExcelMappingPlanBuilder(IExcelMappingPlanFactory factory)
    {
        _factory = factory ?? throw new ArgumentNullException(nameof(factory));
    }

    /// <summary>
    /// 按运行时实体类型创建导入或导出映射计划。
    /// </summary>
    /// <param name="itemType">行实体运行时类型。</param>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>指定实体类型和方向的映射计划。</returns>
    public IExcelMappingPlan Create(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
    {
        if (itemType == null)
            throw new ArgumentNullException(nameof(itemType));
        return Invokers.GetOrAdd(itemType, CreateInvoker)(this, document, configuration, direction);
    }

    /// <summary>
    /// 为运行时实体类型创建泛型计划构建反射委托。
    /// </summary>
    /// <param name="itemType">行实体运行时类型。</param>
    /// <returns>绑定到指定实体类型的计划构建委托。</returns>
    private static CreatePlanInvoker CreateInvoker(Type itemType)
    {
        var method = typeof(MiniExcelMappingPlanBuilder).GetMethod(nameof(CreateTyped),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(itemType);
        return (CreatePlanInvoker)method.CreateDelegate(typeof(CreatePlanInvoker));
    }

    /// <summary>
    /// 使用具体实体类型创建映射计划。
    /// </summary>
    /// <typeparam name="T">行实体类型。</typeparam>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">映射配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>具体实体类型的映射计划。</returns>
    private IExcelMappingPlan CreateTyped<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration configuration, MappingDirection direction)
        where T : class, new()
    {
        return _factory.Create<T>(document ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        }, configuration, direction);
    }

    /// <summary>
    /// 将工作表请求中的动态列合并到映射配置。
    /// </summary>
    /// <param name="configuration">原始映射配置。</param>
    /// <param name="definitions">请求声明的动态列定义。</param>
    /// <returns>包含请求动态列的映射配置。</returns>
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

    /// <summary>
    /// 获取列布局的稳定键文本。
    /// </summary>
    /// <param name="placement">列布局定义。</param>
    /// <returns>布局键文本；未指定布局时返回 null。</returns>
    private static string GetPlacementKey(ExcelColumnPlacement placement)
    {
        if (!string.IsNullOrWhiteSpace(placement?.BeforeKey))
            return $"before:{placement.BeforeKey}";
        if (!string.IsNullOrWhiteSpace(placement?.AfterKey))
            return $"after:{placement.AfterKey}";
        return null;
    }

    /// <summary>
    /// 将 CLR 类型映射为动态列数据类型名称。
    /// </summary>
    /// <param name="type">待映射的 CLR 类型。</param>
    /// <returns>配置使用的数据类型名称。</returns>
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
