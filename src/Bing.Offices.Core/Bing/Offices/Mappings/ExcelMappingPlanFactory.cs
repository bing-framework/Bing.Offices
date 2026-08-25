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
    private readonly Configurations.IMappingProfileResolver _profileRegistry;
    private readonly Configurations.ExcelModelAliasRegistry _modelAliases;
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    private readonly int _cacheCapacity;
    private readonly ConcurrentDictionary<string, Lazy<IExcelMappingPlan>> _planCache = new();
    private readonly ConcurrentQueue<string> _cacheOrder = new();

    /// <summary>初始化计划工厂。</summary>
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

    private ExcelMappingColumn CreateColumn(ExcelPropertyMap property, bool allowImplicitNamedConverters) =>
        new ExcelMappingColumn(property, BindValueConverters(property, allowImplicitNamedConverters),
            BindValidationRules(property));

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

    private static bool IsDynamicValueContainer(Type propertyType) =>
        typeof(IDictionary<string, object>).IsAssignableFrom(propertyType);

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
        public ResolvedMapping(ExcelMappingConfiguration configuration, string profileName,
            string modelAlias, bool allowImplicitNamedConverters)
        {
            Configuration = configuration;
            ProfileName = profileName;
            ModelAlias = modelAlias;
            AllowImplicitNamedConverters = allowImplicitNamedConverters;
        }

        public ExcelMappingConfiguration Configuration { get; }
        public string ProfileName { get; }
        public string ModelAlias { get; }
        public bool AllowImplicitNamedConverters { get; }
    }

    private void TrimCache()
    {
        while (_planCache.Count > _cacheCapacity && _cacheOrder.TryDequeue(out var oldest))
            _planCache.TryRemove(oldest, out _);
    }
}

internal sealed class ExcelMappingPlan : IExcelMappingPlan
{
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

    public string ProfileName { get; }
    public string ModelAlias { get; }
    public IReadOnlyList<IExcelMappingColumn> Columns { get; }
    public IReadOnlyList<IExcelDynamicMappingColumn> DynamicColumns { get; }
    public IExcelMappingStyle Style { get; }
    public IExcelMappingLayout Layout { get; }
}

internal sealed class ExcelDynamicMappingColumn : IExcelDynamicMappingColumn
{
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

    public string Key { get; }
    public string Title { get; }
    public IReadOnlyList<string> Aliases { get; }
    public string DataTypeName { get; }
    public int Order { get; }
    public string ConverterName { get; }
    public string ValidatorName { get; }
    public IReadOnlyList<string> ValidationRuleNames { get; }
    public IReadOnlyList<ExcelMappingDynamicValidationConfiguration> ValidationRules { get; }
    public string NumberFormat { get; }
    public int? ColumnIndex { get; }
    public string PlacementKey { get; }
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    public bool IsUnique { get; }
    public bool UniqueIgnoreEmpty { get; }
}

internal sealed class ExcelMappingStyle : IExcelMappingStyle
{
    internal ExcelMappingStyle(ExcelMappingStyleConfiguration style)
    {
        HeaderStyleKey = style?.HeaderStyleKey;
        BodyStyleKey = style?.BodyStyleKey;
        NumberFormat = style?.NumberFormat;
    }
    public string HeaderStyleKey { get; }
    public string BodyStyleKey { get; }
    public string NumberFormat { get; }
}

internal sealed class ExcelMappingLayout : IExcelMappingLayout
{
    internal ExcelMappingLayout(ExcelMappingLayoutConfiguration layout)
    {
        ColumnIndex = layout?.ColumnIndex;
        PlacementKey = layout?.PlacementKey;
    }
    public int? ColumnIndex { get; }
    public string PlacementKey { get; }
}

internal sealed class ExcelMappingColumn : IExcelMappingColumn, IExcelCompiledMappingColumn
{
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

    public string Name { get; }
    public string Title { get; }
    public IReadOnlyList<string> Aliases { get; }
    public string Formatter { get; }
    public bool Ignored { get; }
    public bool IsDynamicColumn { get; }
    public Bing.Offices.Imports.ExcelWhitespacePolicy? ImportWhitespace { get; }
    public byte? DecimalScale { get; }
    public string ConverterName { get; }
    public IReadOnlyList<string> ValidationRuleNames { get; }
    public IReadOnlyDictionary<string, string> ValueMap { get; }
    public Bing.Offices.Imports.ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    public bool IsUnique { get; }
    public bool UniqueIgnoreEmpty { get; }
    public IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    public IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    public PropertyInfo Property { get; }
    public Func<object, object> Getter { get; }
    public Action<object, object> Setter { get; }
    public IReadOnlyList<Attribute> Attributes { get; }
}

internal sealed class ExcelMappingWorkbookPlan : IExcelMappingWorkbookPlan
{
    internal ExcelMappingWorkbookPlan(IReadOnlyList<IExcelMappingSheetPlan> sheets)
    {
        Sheets = new ReadOnlyCollection<IExcelMappingSheetPlan>(sheets.ToArray());
    }

    public IReadOnlyList<IExcelMappingSheetPlan> Sheets { get; }
}

internal sealed class ExcelMappingSheetPlan : IExcelMappingSheetPlan
{
    internal ExcelMappingSheetPlan(string name, IExcelMappingPlan mapping)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Sheet 名称不能为空。", nameof(name));
        Name = name;
        Mapping = mapping ?? throw new ArgumentNullException(nameof(mapping));
    }

    public string Name { get; }
    public IExcelMappingPlan Mapping { get; }
}
