using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Reflection;

namespace Bing.Offices.Mappings;

/// <summary>
/// Excel 类型映射工厂。
/// </summary>
public static class ExcelTypeMapFactory
{
    /// <summary>
    /// 类型静态映射缓存。
    /// </summary>
    private static readonly ConcurrentDictionary<Type, object> TypeMaps = new();

    /// <summary>
    /// 获取类型的不可变静态映射。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    public static ExcelTypeMap<T> Get<T>() => (ExcelTypeMap<T>)TypeMaps.GetOrAdd(typeof(T), _ => Create<T>());

    /// <summary>
    /// 获取应用请求级配置后的不可变类型映射。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="configuration">请求级映射配置。</param>
    public static ExcelTypeMap<T> Get<T>(ExcelMappingConfiguration configuration)
    {
        return ApplyConfiguration(configuration, Get<T>());
    }

    /// <summary>
    /// 获取 normalized Mapping Document 指定方向的类型映射。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="direction">映射方向。</param>
    public static ExcelTypeMap<T> Get<T>(ExcelMappingDocument document, MappingDirection direction)
    {
        if (document == null)
            throw new ArgumentNullException(nameof(document));
        return Get<T>(direction == MappingDirection.Import ? document.Import : document.Export);
    }

    /// <summary>
    /// 获取 normalized Mapping Document 指定方向的类型映射，并应用请求级配置。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="configuration">请求级映射配置。</param>
    /// <param name="direction">映射方向。</param>
    public static ExcelTypeMap<T> Get<T>(ExcelMappingDocument document, ExcelMappingConfiguration configuration,
        MappingDirection direction)
    {
        return Get<T>(configuration, Get<T>(document, direction));
    }

    /// <summary>
    /// 将配置编译到给定的不可变类型映射。
    /// </summary>
    private static ExcelTypeMap<T> ApplyConfiguration<T>(ExcelMappingConfiguration configuration, ExcelTypeMap<T> source)
    {
        if (configuration == null || configuration.Columns == null || configuration.Columns.Count == 0)
            return source;

        var configuredColumns = new Dictionary<string, ExcelColumnConfiguration>(StringComparer.OrdinalIgnoreCase);
        var configuredIndexes = new HashSet<int>();
        foreach (var column in configuration.Columns)
        {
            if (column == null || string.IsNullOrWhiteSpace(column.PropertyName))
                throw new ArgumentException("映射配置的属性名称不能为空。", nameof(configuration));
            if (column.ColumnIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(configuration), "映射配置的列索引不能小于零。");
            if (configuredColumns.ContainsKey(column.PropertyName))
                throw new ArgumentException($"映射配置包含重复属性: {column.PropertyName}", nameof(configuration));
            if (column.ColumnIndex.HasValue && !configuredIndexes.Add(column.ColumnIndex.Value))
                throw new ArgumentException($"映射配置包含重复列索引: {column.ColumnIndex.Value}", nameof(configuration));
            configuredColumns.Add(column.PropertyName, column);
        }

        var properties = new List<(ExcelPropertyMap Property, int Order)>();
        var defaultOrder = 0;
        foreach (var property in source.Properties)
        {
            configuredColumns.TryGetValue(property.Name, out var configurationColumn);
            if (configurationColumn == null)
            {
                properties.Add((property, int.MaxValue - source.Properties.Count + defaultOrder++));
                continue;
            }
            if (property.IsDynamicColumn)
                throw new InvalidOperationException($"动态列属性不支持请求级固定列映射: {property.Name}");
            var title = string.IsNullOrWhiteSpace(configurationColumn.Title) ? property.Title : configurationColumn.Title;
            var values = CreateConfiguredValueMap(property, configurationColumn);
            var aliases = CreateAliases(configurationColumn, property.Aliases);
            var configuredProperty = new ExcelPropertyMap(property.Property, title,
                configurationColumn.Formatter ?? property.Formatter, configurationColumn.Ignored ?? property.Ignored,
                property.IsDynamicColumn, configurationColumn.ImportWhitespace ?? property.ImportWhitespace,
                configurationColumn.DecimalScale ?? property.DecimalScale,
                configurationColumn.ConverterName ?? property.ConverterName,
                CreateValidationRuleNames(configurationColumn, property.ValidationRuleNames), values, aliases,
                property.Getter,
                property.Setter, configurationColumn.ImageMultiplicity ?? property.ImageMultiplicity);
            properties.Add((configuredProperty, configurationColumn.ColumnIndex ?? int.MaxValue - source.Properties.Count + defaultOrder++));
        }
        var unknownProperty = configuredColumns.Keys.FirstOrDefault(name =>
            source.Properties.All(property => !string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)));
        if (unknownProperty != null)
            throw new ArgumentException($"映射配置引用了不存在的属性: {unknownProperty}", nameof(configuration));
        var mappedProperties = properties
            .OrderBy(item => item.Order)
            .Select(item => item.Property)
            .ToList();
        ValidateTitles(mappedProperties, configuration);
        return new ExcelTypeMap<T>(new ReadOnlyCollection<ExcelPropertyMap>(mappedProperties));
    }

    /// <summary>
    /// 获取依次应用 Fluent Profile 与请求配置后的类型映射。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    /// <param name="profile">优先级低于请求配置的 Fluent Profile。</param>
    /// <param name="configuration">优先级最高的请求配置。</param>
    public static ExcelTypeMap<T> Get<T>(ExcelMappingProfile<T> profile, ExcelMappingConfiguration configuration)
        where T : class, new() =>
        Get<T>(configuration, profile == null ? Get<T>() : Get<T>(profile.Configuration));

    /// <summary>
    /// 获取指定方向的双模型 Profile 映射，并按请求配置继续覆盖。
    /// </summary>
    [Obsolete("请改用 ExcelMappingDocument 和 IExcelMappingPlanFactory。", false)]
    public static ExcelTypeMap<T> Get<T>(object profile, ExcelMappingConfiguration configuration,
        MappingDirection direction) where T : class, new()
    {
        if (profile == null)
            return Get<T>(configuration);
        if (profile is ExcelMappingProfile<T> legacy)
            return Get(legacy, configuration);
        if (!(profile is IMappingProfileSnapshot snapshot))
            throw new ArgumentException("映射 Profile 类型不受支持。", nameof(profile));
        var expectedType = direction == MappingDirection.Import ? snapshot.ImportType : snapshot.ExportType;
        if (expectedType != typeof(T))
            throw new ArgumentException($"映射 Profile 的{(direction == MappingDirection.Import ? "导入" : "导出")}模型类型不匹配: {typeof(T).FullName}",
                nameof(profile));
        var profileConfiguration = direction == MappingDirection.Import
            ? snapshot.ImportConfiguration
            : snapshot.ExportConfiguration;
        var mapped = Get<T>(profileConfiguration);
        return Get<T>(configuration, mapped);
    }

    private static ExcelTypeMap<T> Get<T>(ExcelMappingConfiguration configuration, ExcelTypeMap<T> source)
    {
        if (configuration == null || configuration.Columns == null || configuration.Columns.Count == 0)
            return source;

        return ApplyConfiguration(configuration, source);
    }

    /// <summary>
    /// 验证请求配置应用后的固定列标题保持唯一。
    /// </summary>
    private static void ValidateTitles(IEnumerable<ExcelPropertyMap> properties, ExcelMappingConfiguration configuration)
    {
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties.Where(property => !property.Ignored && !property.IsDynamicColumn))
        {
            if (!titles.Add(property.Title))
                throw new ArgumentException($"映射配置包含重复列标题: {property.Title}", nameof(configuration));
        }
    }

    /// <summary>
    /// 创建类型映射。
    /// </summary>
    /// <typeparam name="T">实体类型。</typeparam>
    private static ExcelTypeMap<T> Create<T>()
    {
        var properties = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.CanRead)
            .Select(CreatePropertyMap)
            .ToList();
        return new ExcelTypeMap<T>(new ReadOnlyCollection<ExcelPropertyMap>(properties));
    }

    /// <summary>
    /// 创建属性映射。
    /// </summary>
    /// <param name="property">属性元数据。</param>
    private static ExcelPropertyMap CreatePropertyMap(PropertyInfo property)
    {
        foreach (var attribute in property.GetCustomAttributes<ExcelRegexAttribute>())
            _ = new Regex(attribute.Pattern, RegexOptions.Compiled, TimeSpan.FromSeconds(1));
        foreach (var attribute in property.GetCustomAttributes<RegexAttribute>())
            _ = new Regex(attribute.RegexString, RegexOptions.Compiled, TimeSpan.FromSeconds(1));
        var isDynamicColumn = property.IsDefined(typeof(DynamicColumnAttribute));
        if (isDynamicColumn &&
            (!property.CanWrite || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType)))
            throw new OfficeException($"【{property.Name}】动态列属性必须是可写的 IDictionary<string, object> 类型");
        var mappings = property.GetCustomAttributes<ValueMappingAttribute>().ToList();
        if (isDynamicColumn && mappings.Any())
            throw new OfficeException($"【{property.Name}】该属性已设置动态列，无法再设置值映射");

        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var mapping in mappings)
            AddMapping(values, property.Name, mapping.Text, mapping.Value);
        if (!mappings.Any())
            AddDefaultValues(property.PropertyType, values);

        var format = property.GetCustomAttribute<DataFormatAttribute>()?.CustomFormat
            ?? property.GetCustomAttribute<DisplayFormatAttribute>()?.DataFormatString;
        var scale = Types.IsNumericType(Reflections.GetUnderlyingType(property.PropertyType))
            ? property.GetCustomAttribute<DecimalScaleAttribute>()?.Scale
            : null;
        var title = property.GetCustomAttribute<ColumnNameAttribute>()?.Name ?? property.Name;
        return new ExcelPropertyMap(property, title, format, property.HasIgnore(), isDynamicColumn, null, scale, null,
            Array.Empty<string>(),
            new ReadOnlyDictionary<string, object>(values), Array.Empty<string>(), CreateGetter(property),
            CreateSetter(property));
    }

    /// <summary>
    /// 添加内置布尔值和枚举值映射。
    /// </summary>
    /// <param name="propertyType">属性类型。</param>
    /// <param name="values">目标映射字典。</param>
    private static void AddDefaultValues(Type propertyType, IDictionary<string, object> values)
    {
        var underlyingType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (underlyingType == typeof(bool))
        {
            AddIfMissing(values, Resources.Yes, true);
            AddIfMissing(values, Resources.No, false);
        }
        if (!underlyingType.IsEnum)
            return;
        foreach (var value in underlyingType.GetEnumValueDefinitionList())
            AddIfMissing(values, value.Description ?? value.DisplayName ?? value.Name,
                Enum.ToObject(underlyingType, value.Value));
        if (Nullable.GetUnderlyingType(propertyType) != null)
            AddIfMissing(values, string.Empty, null);
    }

    /// <summary>
    /// 在键不存在时添加映射值。
    /// </summary>
    /// <param name="values">目标映射字典。</param>
    /// <param name="text">显示文本。</param>
    /// <param name="value">业务值。</param>
    private static void AddIfMissing(IDictionary<string, object> values, string text, object value)
    {
        if (!values.ContainsKey(text))
            values.Add(text, value);
    }

    /// <summary>
    /// 添加自定义值映射，并确保显示文本和值均可逆。
    /// </summary>
    /// <param name="values">目标映射字典。</param>
    /// <param name="propertyName">属性名称。</param>
    /// <param name="text">显示文本。</param>
    /// <param name="value">业务值。</param>
    private static void AddMapping(IDictionary<string, object> values, string propertyName, string text, object value)
    {
        if (values.ContainsKey(text))
            throw new OfficeException($"【{propertyName}】存在重复的值映射文本: {text}");
        if (values.Values.Any(mappedValue => Equals(mappedValue, value)))
            throw new OfficeException($"【{propertyName}】存在重复的值映射业务值: {value}");
        values.Add(text, value);
    }

    /// <summary>
    /// 从请求级配置创建显示值映射。
    /// </summary>
    /// <param name="property">默认属性映射。</param>
    /// <param name="configuration">列配置。</param>
    private static IReadOnlyDictionary<string, object> CreateConfiguredValueMap(ExcelPropertyMap property,
        ExcelColumnConfiguration configuration)
    {
        if (configuration.ValueMappings == null || configuration.ValueMappings.Count == 0)
            return property.ValueMap;
        var values = new Dictionary<string, object>(StringComparer.Ordinal);
        if (configuration.ValueMappingMergeMode == ExcelValueMappingMergeMode.Append)
            foreach (var pair in property.ValueMap)
                values.Add(pair.Key, pair.Value);
        foreach (var mapping in configuration.ValueMappings)
        {
            if (mapping == null || string.IsNullOrWhiteSpace(mapping.Text))
                throw new ArgumentException($"属性 {property.Name} 的显示值映射文本不能为空。", nameof(configuration));
            AddMapping(values, property.Name, mapping.Text, ConvertConfigurationValue(mapping.Value,
                property.Property.PropertyType));
        }
        return new ReadOnlyDictionary<string, object>(values);
    }

    private static IReadOnlyList<string> CreateAliases(ExcelColumnConfiguration configuration,
        IReadOnlyList<string> defaults)
    {
        if (configuration.Aliases == null || configuration.Aliases.Count == 0)
            return defaults ?? Array.Empty<string>();
        var aliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var alias in configuration.Aliases)
        {
            if (string.IsNullOrWhiteSpace(alias))
                throw new ArgumentException("映射配置的列别名不能为空。", nameof(configuration));
            aliases.Add(alias);
        }
        return aliases.ToArray();
    }

    /// <summary>
    /// 验证并创建配置的命名校验器引用。
    /// </summary>
    /// <param name="configuration">列配置。</param>
    /// <param name="defaultNames">配置未指定规则时继承的默认规则名称。</param>
    private static IReadOnlyList<string> CreateValidationRuleNames(ExcelColumnConfiguration configuration,
        IReadOnlyList<string> defaultNames = null)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!configuration.ClearValidationRules)
            foreach (var defaultName in defaultNames ?? Array.Empty<string>())
                names.Add(defaultName);
        if (configuration.ValidationRuleMergeMode == ExcelValidationRuleMergeMode.Replace)
            names.Clear();
        foreach (var name in configuration.ValidationRuleNames ?? new List<string>())
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("映射配置的校验规则名称不能为空。", nameof(configuration));
            if (!names.Add(name))
                throw new ArgumentException($"映射配置包含重复校验规则名称: {name}", nameof(configuration));
        }
        foreach (var name in configuration.ValidationRuleNamesToRemove ?? new List<string>())
            names.Remove(name);
        return names.ToArray();
    }

    /// <summary>
    /// 将可序列化配置值转换为属性类型。
    /// </summary>
    /// <param name="value">配置值文本。</param>
    /// <param name="propertyType">目标属性类型。</param>
    private static object ConvertConfigurationValue(string value, Type propertyType)
    {
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (value == null)
        {
            if (!propertyType.IsValueType || Nullable.GetUnderlyingType(propertyType) != null)
                return null;
            throw new ArgumentException($"非空值类型 {propertyType.FullName} 不允许使用空映射值。");
        }
        if (targetType == typeof(string))
            return value;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 编译属性读取器。
    /// </summary>
    /// <param name="property">属性元数据。</param>
    private static Func<object, object> CreateGetter(PropertyInfo property)
    {
        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Property(Expression.Convert(instance, property.DeclaringType!), property);
        return Expression.Lambda<Func<object, object>>(Expression.Convert(value, typeof(object)), instance).Compile();
    }

    /// <summary>
    /// 编译属性写入器。
    /// </summary>
    /// <param name="property">属性元数据。</param>
    private static Action<object, object> CreateSetter(PropertyInfo property)
    {
        if (!property.CanWrite)
            return (_, _) => throw new InvalidOperationException($"属性不可写入: {property.Name}");

        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Parameter(typeof(object), "value");
        var setter = Expression.Assign(Expression.Property(Expression.Convert(instance, property.DeclaringType!), property),
            Expression.Convert(value, property.PropertyType));
        return Expression.Lambda<Action<object, object>>(setter, instance, value).Compile();
    }
}
