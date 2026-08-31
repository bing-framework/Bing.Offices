using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text.Json;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Validations;
using Bing.Offices.Providers;

namespace Bing.Offices.Mappings;

/// <summary>
/// Core 对 Provider-neutral 计划契约的实现。
/// </summary>
public sealed class ExcelMappingPlanFactory : IExcelMappingPlanFactory
{
    /// <summary>按名称解析映射 Profile 的注册表。</summary>
    private readonly Configurations.IMappingProfileResolver _profileRegistry;
    /// <summary>验证模型别名及其目标类型的注册表。</summary>
    private readonly Configurations.ExcelModelAliasRegistry _modelAliases;
    /// <summary>可用于列绑定的值转换器集合。</summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>可用于特性校验绑定的规则集合。</summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    /// <summary>可通过名称绑定的校验规则集合。</summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>映射计划缓存允许保留的最大条目数。</summary>
    private readonly int _cacheCapacity;
    /// <summary>按规范化配置哈希缓存的延迟映射计划。</summary>
    private readonly ConcurrentDictionary<string, Lazy<IExcelMappingPlan>> _planCache = new();
    /// <summary>按创建顺序记录缓存键，用于近似先进先出淘汰。</summary>
    private readonly ConcurrentQueue<string> _cacheOrder = new();

    /// <summary>使用转换器、校验规则和可选 Profile 注册表初始化计划工厂。</summary>
    /// <param name="valueConverters">可绑定到映射列的值转换器。</param>
    /// <param name="validationRules">可绑定到校验特性的规则。</param>
    /// <param name="namedValidationRules">可按名称绑定的校验规则。</param>
    /// <param name="cacheCapacity">映射计划缓存的最大条目数。</param>
    /// <param name="profileRegistry">可选的映射 Profile 解析器。</param>
    /// <param name="modelAliases">可选的模型别名注册表。</param>
    public ExcelMappingPlanFactory(IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        int cacheCapacity = 256,
        Configurations.IMappingProfileResolver profileRegistry = null,
        Configurations.ExcelModelAliasRegistry modelAliases = null)
    {
        _profileRegistry = profileRegistry;
        _modelAliases = modelAliases;
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault().ToArray();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        if (cacheCapacity <= 0)
            throw new ArgumentOutOfRangeException(nameof(cacheCapacity));
        _cacheCapacity = cacheCapacity;
    }

    /// <inheritdoc />
    public IExcelMappingPlan Create<T>(ExcelMappingDocument document, MappingDirection direction)
        where T : class, new()
    {
        return Create<T>(document, null, direction);
    }

    /// <inheritdoc />
    public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document, MappingDirection direction,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        if (sheetNames == null || sheetNames.Count == 0)
            throw new ArgumentException("Workbook 至少需要一个 Sheet。", nameof(sheetNames));
        var mapping = Create<T>(document, direction);
        return new ExcelMappingWorkbookPlan(sheetNames.Select(name =>
            new ExcelMappingSheetPlan(name, mapping)).ToArray());
    }

    /// <inheritdoc />
    public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction,
        IReadOnlyList<string> sheetNames) where T : class, new()
    {
        if (sheetNames == null || sheetNames.Count == 0)
            throw new ArgumentException("Workbook 至少需要一个 Sheet。", nameof(sheetNames));
        var mapping = Create<T>(document, requestConfiguration, direction);
        return new ExcelMappingWorkbookPlan(sheetNames.Select(name =>
            new ExcelMappingSheetPlan(name, mapping)).ToArray());
    }

    /// <inheritdoc />
    public IExcelMappingPlan Create<T>(ExcelMappingDocument document, ExcelMappingConfiguration configuration,
        MappingDirection direction) where T : class, new()
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        var resolved = ResolveDocument<T>(document, configuration, direction);
        var key = CreateCacheKey<T>(document, direction, resolved.Configuration);
        if (_planCache.TryGetValue(key, out var existing))
            return existing.Value;
        var created = new Lazy<IExcelMappingPlan>(() => CreatePlan<T>(resolved.Configuration, resolved.ProfileName,
            resolved.ModelAlias, resolved.AllowImplicitNamedConverters),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
        var cached = _planCache.GetOrAdd(key, created);
        if (ReferenceEquals(cached, created))
        {
            _cacheOrder.Enqueue(key);
            TrimCache();
        }
        return cached.Value;
    }

    /// <summary>按方向合并文档、Profile 和请求级配置，并解析模型别名约束。</summary>
    /// <typeparam name="T">目标实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="requestConfiguration">可选的请求级覆盖配置。</param>
    /// <param name="direction">导入或导出方向。</param>
    /// <returns>已解析的配置及其来源元数据。</returns>
    private ResolvedMapping ResolveDocument<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
        where T : class, new()
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));
        var normalized = ExcelMappingDocumentFactory.Create<T>(document, null, direction);
        var directionConfiguration = direction == MappingDirection.Import ? normalized.Import : normalized.Export;
        if (directionConfiguration == null && requestConfiguration == null && !document.UseConventionFallback)
            throw new InvalidOperationException(
                $"映射文档未提供方向配置: 文档={typeof(ExcelMappingDocument).Name}，方向={direction}，模型={typeof(T).FullName}。"
                + "如需约定映射，请显式设置 UseConventionFallback=true。");
        var profileName = directionConfiguration?.Profile;
        var modelAlias = directionConfiguration?.ModelAlias;
        if (_modelAliases != null && _modelAliases.HasRegistrations
            && !string.IsNullOrWhiteSpace(modelAlias))
        {
            if (!_modelAliases.TryResolve(modelAlias, out var modelType, out var aliasProfile))
                throw new InvalidOperationException($"未知 modelAlias: {modelAlias}");
            if (modelType != null && modelType != typeof(T))
                throw new InvalidOperationException($"modelAlias 与模型类型不匹配: {modelAlias}");
            if (!string.IsNullOrWhiteSpace(aliasProfile) && string.IsNullOrWhiteSpace(profileName))
                throw new InvalidOperationException($"modelAlias 要求方向配置提供 Profile: {modelAlias}");
            if (!string.IsNullOrWhiteSpace(aliasProfile)
                && !string.IsNullOrWhiteSpace(profileName)
                && !string.Equals(aliasProfile, profileName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"modelAlias 与 Profile 不匹配: {modelAlias}");
        }
        ProfileDescriptor profileDescriptor = null;
        if (_profileRegistry != null && !string.IsNullOrWhiteSpace(profileName)
            && !_profileRegistry.TryGetDescriptor(profileName, direction, typeof(T), out profileDescriptor))
            throw new InvalidOperationException(
                $"未找到匹配的 Profile: {profileName}，方向: {direction}，模型: {typeof(T).FullName}");
        var configuration = profileDescriptor == null
            ? directionConfiguration
            : MappingConfigurationMerger.Merge(profileDescriptor.Configuration,
                directionConfiguration, MappingSourceKind.Document);
        if (requestConfiguration != null)
            configuration = MappingConfigurationMerger.Merge(configuration, requestConfiguration,
                MappingSourceKind.Request);
        return new ResolvedMapping(configuration, profileName, modelAlias, profileDescriptor == null);
    }

    /// <summary>将已合并配置编译为不可变的 Provider-neutral 映射计划。</summary>
    /// <typeparam name="T">目标实体类型。</typeparam>
    /// <param name="configuration">已合并的方向配置。</param>
    /// <param name="profileName">来源 Profile 名称。</param>
    /// <param name="modelAlias">来源模型别名。</param>
    /// <param name="allowImplicitNamedConverters">是否允许隐式绑定命名转换器。</param>
    /// <returns>编译后的映射计划。</returns>
    private IExcelMappingPlan CreatePlan<T>(ExcelMappingConfiguration configuration, string profileName,
        string modelAlias, bool allowImplicitNamedConverters) where T : class, new()
    {
        var map = ExcelTypeMapFactory.Get<T>(configuration);
        var dynamicColumns = (configuration?.DynamicColumns ?? new List<ExcelMappingDynamicColumnConfiguration>())
            .Select(CreateDynamicColumn).ToArray();
        ValidateDynamicColumns(dynamicColumns);
        return new ExcelMappingPlan(map.Properties
                .Select(property => CreateColumn(property, allowImplicitNamedConverters)).ToArray(),
            dynamicColumns, new ExcelMappingStyle(configuration?.Style),
            new ExcelMappingLayout(configuration?.Layout), profileName, modelAlias);
    }

    /// <summary>从实体属性映射创建列计划并绑定转换器和校验规则。</summary>
    /// <param name="property">已解析的实体属性映射。</param>
    /// <param name="allowImplicitNamedConverters">是否允许隐式绑定命名转换器。</param>
    /// <returns>编译后的列计划。</returns>
    private ExcelMappingColumn CreateColumn(ExcelPropertyMap property, bool allowImplicitNamedConverters) =>
        new ExcelMappingColumn(property, BindValueConverters(property, allowImplicitNamedConverters),
            BindValidationRules(property));

    /// <summary>解析列声明的值转换器，并按配置过滤隐式命名转换器。</summary>
    /// <param name="property">实体属性映射。</param>
    /// <param name="allowImplicitNamedConverters">是否允许未显式命名的命名转换器。</param>
    /// <returns>适用于该属性的转换器集合。</returns>
    private IReadOnlyList<IExcelValueConverter> BindValueConverters(ExcelPropertyMap property,
        bool allowImplicitNamedConverters)
    {
        if (property.Ignored || property.IsDynamicColumn || IsDynamicValueContainer(property.Property.PropertyType))
            return Array.Empty<IExcelValueConverter>();
        var converters = ExcelValueConverterBindingResolver.Resolve(_valueConverters,
            property.ConverterName, property.Property.PropertyType);
        return string.IsNullOrWhiteSpace(property.ConverterName) && !allowImplicitNamedConverters
            ? converters.Where(converter => converter is not INamedExcelValueConverter).ToArray()
            : converters;
    }

    /// <summary>判断属性是否为承载动态列值的字典容器。</summary>
    /// <param name="propertyType">属性类型。</param>
    /// <returns>可赋值为字符串到对象字典时为 true。</returns>
    private static bool IsDynamicValueContainer(Type propertyType) =>
        typeof(IDictionary<string, object>).IsAssignableFrom(propertyType);

    /// <summary>绑定属性特性校验规则和配置声明的命名校验规则。</summary>
    /// <param name="property">实体属性映射。</param>
    /// <returns>按执行顺序排列的校验绑定集合。</returns>
    private IReadOnlyList<IExcelValidationBinding> BindValidationRules(ExcelPropertyMap property)
    {
        var bindings = new List<IExcelValidationBinding>();
        foreach (var attribute in property.Property.GetCustomAttributes<FilterAttributeBase>())
        {
            var rule = ResolveValidationRule(attribute);
            bindings.Add(ExcelValidationBinding.Attribute(attribute, rule));
        }
        foreach (var ruleName in property.ValidationRuleNames ?? Array.Empty<string>())
        {
            var rule = _namedValidationRules.Where(candidate => string.Equals(candidate.Name, ruleName,
                StringComparison.OrdinalIgnoreCase)).ToArray();
            if (rule.Length != 1)
                throw new InvalidOperationException($"未找到唯一命名校验规则: {ruleName}");
            bindings.Add(ExcelValidationBinding.Named(rule[0]));
        }
        return bindings;
    }

    /// <summary>将动态列配置解析为包含转换器和校验绑定的列计划。</summary>
    /// <param name="column">规范化动态列配置。</param>
    /// <returns>编译后的动态列计划。</returns>
    private IExcelDynamicMappingColumn CreateDynamicColumn(ExcelMappingDynamicColumnConfiguration column)
    {
        if (column == null)
            throw new ArgumentException("动态列配置不能为 null。", nameof(column));
        if (string.IsNullOrWhiteSpace(column.Key) || string.IsNullOrWhiteSpace(column.Title))
            throw new ArgumentException("动态列 Key 和 Title 不能为空。", nameof(column));
        var dataType = ResolveDynamicType(column.DataTypeName);
        var converters = ExcelValueConverterBindingResolver.Resolve(_valueConverters, column.ConverterName, dataType);
        var validations = new List<IExcelValidationBinding>();
        foreach (var validation in column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
        {
            if (validation == null || string.IsNullOrWhiteSpace(validation.Name))
                throw new InvalidOperationException($"动态列 {column.Key} 的内置校验规则名称不能为空。");
            var attribute = CreateDynamicValidationAttribute(validation, column.Key);
            validations.Add(ExcelValidationBinding.Attribute(attribute, ResolveValidationRule(attribute)));
        }
        var validatorNames = (column.ValidationRuleNames ?? new List<string>()).ToList();
        if (!string.IsNullOrWhiteSpace(column.ValidatorName)
            && !validatorNames.Contains(column.ValidatorName, StringComparer.OrdinalIgnoreCase))
            validatorNames.Insert(0, column.ValidatorName);
        foreach (var validatorName in validatorNames)
        {
            var rules = _namedValidationRules.Where(rule => string.Equals(rule.Name, validatorName,
                StringComparison.OrdinalIgnoreCase)).ToArray();
            if (rules.Length != 1)
                throw new InvalidOperationException($"未找到唯一命名校验规则: {validatorName}");
            validations.Add(ExcelValidationBinding.Named(rules[0]));
        }
        return new ExcelDynamicMappingColumn(column, converters, validations);
    }

    /// <summary>解析动态列允许使用的 CLR 类型名称。</summary>
    /// <param name="dataTypeName">配置中的类型名称；为空时使用 string。</param>
    /// <returns>对应的 CLR 类型。</returns>
    private static Type ResolveDynamicType(string dataTypeName)
    {
        switch ((dataTypeName ?? "string").Trim().ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "boolean": case "bool": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int32": case "int": return typeof(int);
            case "int64": case "long": return typeof(long);
            case "single": case "float": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new InvalidOperationException($"动态列数据类型不在允许列表中: {dataTypeName}");
        }
    }

    /// <summary>根据动态列规则配置创建内置校验特性。</summary>
    /// <param name="validation">动态校验规则配置。</param>
    /// <param name="columnKey">所属动态列键，用于错误信息。</param>
    /// <returns>对应的校验特性实例。</returns>
    private static FilterAttributeBase CreateDynamicValidationAttribute(
        ExcelMappingDynamicValidationConfiguration validation, string columnKey)
    {
        switch (validation.Name.Trim().ToLowerInvariant())
        {
            case "required":
                return new ExcelRequiredAttribute();
            case "regex":
                if (string.IsNullOrWhiteSpace(validation.Pattern))
                    throw new InvalidOperationException($"动态列 {columnKey} 的 regex 规则必须提供 pattern。");
                return new ExcelRegexAttribute(validation.Pattern);
            case "date":
                var date = string.IsNullOrWhiteSpace(validation.Format)
                    ? new ExcelDateAttribute()
                    : new ExcelDateAttribute(validation.Format);
                date.CultureName = validation.CultureName;
                return date;
            case "maxvalue":
            case "max-value":
                if (!validation.MaxValue.HasValue)
                    throw new InvalidOperationException($"动态列 {columnKey} 的 maxValue 规则必须提供 maxValue。");
                return new ExcelMaxValueAttribute(validation.MaxValue.Value);
            case "range":
                if (!validation.Min.HasValue || !validation.Max.HasValue)
                    throw new InvalidOperationException($"动态列 {columnKey} 的 range 规则必须提供 min 和 max。");
                return new ExcelRangeAttribute(validation.Min.Value, validation.Max.Value);
            case "maxlength":
            case "max-length":
                if (!validation.MaxLength.HasValue)
                    throw new InvalidOperationException($"动态列 {columnKey} 的 maxLength 规则必须提供 maxLength。");
                return new ExcelMaxLengthAttribute(validation.MaxLength.Value);
            case "unique":
                return new ExcelUniqueAttribute { IgnoreEmpty = validation.IgnoreEmpty };
            default:
                throw new InvalidOperationException($"动态列 {columnKey} 的内置校验规则不受支持: {validation.Name}");
        }
    }

    /// <summary>验证动态列键、标题及标题别名在同一映射中的唯一性。</summary>
    /// <param name="columns">待验证的动态列计划。</param>
    private static void ValidateDynamicColumns(IReadOnlyList<IExcelDynamicMappingColumn> columns)
    {
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (!keys.Add(column.Key))
                throw new InvalidOperationException($"动态列包含重复 Key: {column.Key}");
            if (!titles.Add(column.Title))
                throw new InvalidOperationException($"动态列包含重复标题: {column.Title}");
            foreach (var alias in column.Aliases)
            {
                if (string.IsNullOrWhiteSpace(alias) || !titles.Add(alias))
                    throw new InvalidOperationException($"动态列包含重复或空标题别名: {alias}");
            }
        }
    }

    /// <summary>根据绑定特性或规则能力解析唯一的校验规则实现。</summary>
    /// <param name="attribute">待绑定的校验特性。</param>
    /// <returns>可处理该特性的校验规则。</returns>
    private IExcelValidationRule ResolveValidationRule(FilterAttributeBase attribute)
    {
        var binding = attribute.GetType().GetCustomAttribute<BindFilterAttribute>();
        if (binding != null)
        {
            var boundRule = _validationRules.FirstOrDefault(rule => binding.RuleType.IsInstanceOfType(rule));
            if (boundRule == null || !boundRule.CanValidate(attribute))
                throw new InvalidOperationException($"未注册或不支持校验特性: {attribute.GetType().FullName}");
            return boundRule;
        }
        var rule = _validationRules.FirstOrDefault(candidate => candidate.CanValidate(attribute));
        return rule ?? throw new InvalidOperationException($"未找到特性对应的校验规则: {attribute.GetType().FullName}");
    }

    /// <summary>根据模型、方向、租户和规范化配置创建稳定缓存键。</summary>
    /// <typeparam name="T">目标实体类型。</typeparam>
    /// <param name="document">原始映射文档。</param>
    /// <param name="direction">导入或导出方向。</param>
    /// <param name="configuration">已解析的方向配置。</param>
    /// <returns>配置内容的 SHA-256 Base64 缓存键。</returns>
    private static string CreateCacheKey<T>(ExcelMappingDocument document, MappingDirection direction,
        ExcelMappingConfiguration configuration) where T : class, new()
    {
        var payload = JsonSerializer.SerializeToUtf8Bytes(new
        {
            document.TenantId,
            ModelType = typeof(T).AssemblyQualifiedName,
            Direction = direction,
            document.ConfigurationVersion,
            Configuration = configuration
        }, new JsonSerializerOptions { IgnoreNullValues = false });
        using var sha256 = SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(payload));
    }

    private sealed class ResolvedMapping
    {
        /// <summary>创建已解析的映射来源快照。</summary>
        /// <param name="configuration">合并后的方向配置。</param>
        /// <param name="profileName">来源 Profile 名称。</param>
        /// <param name="modelAlias">来源模型别名。</param>
        /// <param name="allowImplicitNamedConverters">是否允许隐式绑定命名转换器。</param>
        public ResolvedMapping(ExcelMappingConfiguration configuration, string profileName,
            string modelAlias, bool allowImplicitNamedConverters)
        {
            Configuration = configuration;
            ProfileName = profileName;
            ModelAlias = modelAlias;
            AllowImplicitNamedConverters = allowImplicitNamedConverters;
        }

        /// <summary>获取合并后的方向配置。</summary>
        public ExcelMappingConfiguration Configuration { get; }
        /// <summary>获取来源 Profile 名称。</summary>
        public string ProfileName { get; }
        /// <summary>获取来源模型别名。</summary>
        public string ModelAlias { get; }
        /// <summary>获取是否允许隐式绑定命名转换器。</summary>
        public bool AllowImplicitNamedConverters { get; }
    }

    /// <summary>按近似先进先出策略移除超出容量的计划缓存条目。</summary>
    private void TrimCache()
    {
        while (_planCache.Count > _cacheCapacity && _cacheOrder.TryDequeue(out var oldest))
            _planCache.TryRemove(oldest, out _);
    }
}

internal sealed class ExcelMappingPlan : IExcelMappingPlan
{
    /// <summary>从已编译的列、动态列、样式和布局创建不可变映射计划。</summary>
    /// <param name="columns">固定列映射。</param>
    /// <param name="dynamicColumns">动态列映射。</param>
    /// <param name="style">样式配置。</param>
    /// <param name="layout">布局配置。</param>
    /// <param name="profileName">来源 Profile 名称。</param>
    /// <param name="modelAlias">来源模型别名。</param>
    internal ExcelMappingPlan(IReadOnlyList<IExcelMappingColumn> columns,
        IReadOnlyList<IExcelDynamicMappingColumn> dynamicColumns, IExcelMappingStyle style,
        IExcelMappingLayout layout, string profileName, string modelAlias)
    {
        Columns = new ReadOnlyCollection<IExcelMappingColumn>(columns.ToArray());
        DynamicColumns = new ReadOnlyCollection<IExcelDynamicMappingColumn>(dynamicColumns.ToArray());
        Style = style;
        Layout = layout;
        ProfileName = profileName;
        ModelAlias = modelAlias;
    }

    /// <inheritdoc />
    public string ProfileName { get; }
    /// <inheritdoc />
    public string ModelAlias { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelMappingColumn> Columns { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelDynamicMappingColumn> DynamicColumns { get; }
    /// <inheritdoc />
    public IExcelMappingStyle Style { get; }
    /// <inheritdoc />
    public IExcelMappingLayout Layout { get; }
}

internal sealed class ExcelDynamicMappingColumn : IExcelDynamicMappingColumn
{
    /// <summary>从动态列配置及已绑定的转换器和校验器创建不可变动态列计划。</summary>
    /// <param name="column">规范化后的动态列配置。</param>
    /// <param name="converters">已绑定的值转换器。</param>
    /// <param name="validations">已绑定的校验规则。</param>
    internal ExcelDynamicMappingColumn(ExcelMappingDynamicColumnConfiguration column,
        IReadOnlyList<IExcelValueConverter> converters, IReadOnlyList<IExcelValidationBinding> validations)
    {
        Key = column.Key;
        Title = column.Title;
        Aliases = new ReadOnlyCollection<string>((column.Aliases ?? new List<string>()).ToArray());
        DataTypeName = column.DataTypeName ?? "string";
        Order = column.Order;
        ConverterName = column.ConverterName;
        ValidatorName = column.ValidatorName;
        var validationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList();
        if (!string.IsNullOrWhiteSpace(column.ValidatorName)
            && !validationRuleNames.Contains(column.ValidatorName, StringComparer.OrdinalIgnoreCase))
            validationRuleNames.Insert(0, column.ValidatorName);
        ValidationRuleNames = new ReadOnlyCollection<string>(validationRuleNames.ToArray());
        ValidationRules = new ReadOnlyCollection<ExcelMappingDynamicValidationConfiguration>(
            (column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
                .Select(rule => new ExcelMappingDynamicValidationConfiguration
                {
                    Name = rule.Name,
                    Pattern = rule.Pattern,
                    Format = rule.Format,
                    CultureName = rule.CultureName,
                    Min = rule.Min,
                    Max = rule.Max,
                    MaxValue = rule.MaxValue,
                    MaxLength = rule.MaxLength,
                    IgnoreEmpty = rule.IgnoreEmpty
                }).ToArray());
        NumberFormat = column.NumberFormat;
        ColumnIndex = column.ColumnIndex;
        PlacementKey = column.PlacementKey;
        ImageMultiplicity = column.ImageMultiplicity;
        ValueConverters = new ReadOnlyCollection<IExcelValueConverter>(converters.ToArray());
        ValidationBindings = new ReadOnlyCollection<IExcelValidationBinding>(validations.ToArray());
        var unique = (column.ValidationRules ?? new List<ExcelMappingDynamicValidationConfiguration>())
            .FirstOrDefault(rule => rule != null && string.Equals(rule.Name?.Trim(), "unique",
                StringComparison.OrdinalIgnoreCase));
        IsUnique = unique != null;
        UniqueIgnoreEmpty = unique?.IgnoreEmpty ?? true;
    }

    /// <inheritdoc />
    public string Key { get; }
    /// <inheritdoc />
    public string Title { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> Aliases { get; }
    /// <inheritdoc />
    public string DataTypeName { get; }
    /// <inheritdoc />
    public int Order { get; }
    /// <inheritdoc />
    public string ConverterName { get; }
    /// <inheritdoc />
    public string ValidatorName { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>获取动态列配置中声明的内置校验规则快照。</summary>
    public IReadOnlyList<ExcelMappingDynamicValidationConfiguration> ValidationRules { get; }
    /// <inheritdoc />
    public string NumberFormat { get; }
    /// <inheritdoc />
    public int? ColumnIndex { get; }
    /// <inheritdoc />
    public string PlacementKey { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <inheritdoc />
    public bool IsUnique { get; }
    /// <inheritdoc />
    public bool UniqueIgnoreEmpty { get; }
}

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

internal sealed class ExcelMappingColumn : IExcelMappingColumn, IExcelCompiledMappingColumn
{
    /// <summary>从属性映射及已绑定规则创建不可变列计划。</summary>
    /// <param name="property">已解析的实体属性映射。</param>
    /// <param name="valueConverters">已绑定的值转换器。</param>
    /// <param name="validationBindings">已绑定的校验规则。</param>
    internal ExcelMappingColumn(ExcelPropertyMap property,
        IReadOnlyList<IExcelValueConverter> valueConverters,
        IReadOnlyList<IExcelValidationBinding> validationBindings)
    {
        if (property == null)
            throw new ArgumentNullException(nameof(property));
        Name = property.Name;
        Title = property.Title;
        Aliases = new ReadOnlyCollection<string>(property.Aliases.ToArray());
        Formatter = property.Formatter;
        Ignored = property.Ignored;
        IsDynamicColumn = property.IsDynamicColumn;
        ImportWhitespace = property.ImportWhitespace;
        DecimalScale = property.DecimalScale;
        ConverterName = property.ConverterName;
        ValidationRuleNames = new ReadOnlyCollection<string>(property.ValidationRuleNames.ToArray());
        ValueMap = new ReadOnlyDictionary<string, string>(property.ValueMap.ToDictionary(pair => pair.Key,
            pair => Convert.ToString(pair.Value, CultureInfo.InvariantCulture), StringComparer.Ordinal));
        ImageMultiplicity = property.ImageMultiplicity;
        ValueConverters = new ReadOnlyCollection<IExcelValueConverter>(valueConverters.ToArray());
        ValidationBindings = new ReadOnlyCollection<IExcelValidationBinding>(validationBindings.ToArray());
        IsUnique = ValidationBindings.Any(binding => binding.Kind == ExcelValidationBindingKind.Unique);
        UniqueIgnoreEmpty = property.Property.GetCustomAttributes<ExcelUniqueAttribute>().FirstOrDefault()?.IgnoreEmpty
            ?? true;
        Property = property.Property;
        Getter = property.Getter;
        Setter = property.Setter;
        Attributes = new ReadOnlyCollection<Attribute>(property.Property.GetCustomAttributes<Attribute>().ToArray());
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public string Title { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> Aliases { get; }
    /// <inheritdoc />
    public string Formatter { get; }
    /// <inheritdoc />
    public bool Ignored { get; }
    /// <inheritdoc />
    public bool IsDynamicColumn { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelWhitespacePolicy? ImportWhitespace { get; }
    /// <inheritdoc />
    public byte? DecimalScale { get; }
    /// <inheritdoc />
    public string ConverterName { get; }
    /// <inheritdoc />
    public IReadOnlyList<string> ValidationRuleNames { get; }
    /// <inheritdoc />
    public IReadOnlyDictionary<string, string> ValueMap { get; }
    /// <inheritdoc />
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <inheritdoc />
    public bool IsUnique { get; }
    /// <inheritdoc />
    public bool UniqueIgnoreEmpty { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <inheritdoc />
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <summary>获取映射到的实体属性元数据。</summary>
    public PropertyInfo Property { get; }
    /// <summary>获取从实体读取属性值的委托。</summary>
    public Func<object, object> Getter { get; }
    /// <summary>获取将转换后值写入实体属性的委托。</summary>
    public Action<object, object> Setter { get; }
    /// <summary>获取实体属性上的特性快照。</summary>
    public IReadOnlyList<Attribute> Attributes { get; }
}

internal sealed class ExcelMappingWorkbookPlan : IExcelMappingWorkbookPlan
{
    /// <summary>从工作表映射计划创建不可变 Workbook 计划。</summary>
    /// <param name="sheets">按请求顺序排列的工作表计划。</param>
    internal ExcelMappingWorkbookPlan(IReadOnlyList<IExcelMappingSheetPlan> sheets)
    {
        Sheets = new ReadOnlyCollection<IExcelMappingSheetPlan>(sheets.ToArray());
    }

    /// <inheritdoc />
    public IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}

internal sealed class ExcelMappingSheetPlan : IExcelMappingSheetPlan
{
    /// <summary>为指定工作表名称和列映射创建计划。</summary>
    /// <param name="name">工作表名称。</param>
    /// <param name="mapping">工作表使用的列映射计划。</param>
    internal ExcelMappingSheetPlan(string name, IExcelMappingPlan mapping)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(name));
        Name = name;
        Mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public IExcelMappingPlan Mapping { get; }
}
