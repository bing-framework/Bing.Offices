using System.Globalization;
using System.Reflection;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Mappings;
using Bing.Offices.Validations;
using CsvHelper;
using CsvHelper.Configuration;

namespace Bing.Offices.Csv;

/// <summary>
/// 基于类型映射的 CSV 流式导出器。
/// </summary>
public sealed class CsvEntityExporter : ICsvExporter
{
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    public CsvEntityExporter(IEnumerable<IExcelValueConverter> valueConverters = null) =>
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();

    /// <inheritdoc />
    public void Export<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new CsvExportOptions<T>();
        ValidateOptions(options.Delimiter, options.Quote, options.NewLine, options.Encoding, options.Culture,
            options.FormulaInjectionPolicy);
        var columns = CreateColumns(ExcelTypeMapFactory.Get(options.MappingProfile, options.MappingConfiguration),
            options.DynamicColumns);
        using var writer = new StreamWriter(destination, options.Encoding, 1024, true);
        if (options.IncludeHeader)
            CsvRecordWriter.Write(writer, columns.Select(column => column.Title), options.Delimiter, options.Quote, options.NewLine,
                options.FormulaInjectionPolicy);
        var rowIndex = options.IncludeHeader ? 2 : 1;
        foreach (var item in data)
        {
            cancellationToken.ThrowIfCancellationRequested();
                CsvRecordWriter.Write(writer, columns.Select((column, index) => FormatValue(column, item, rowIndex, index + 1,
                    options.Culture)), options.Delimiter, options.Quote, options.NewLine, options.FormulaInjectionPolicy);
            rowIndex++;
        }
        writer.Flush();
    }

    private static void ValidateOptions(char delimiter, char quote, string newLine, Encoding encoding, CultureInfo culture,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        if (delimiter == quote || delimiter == '\r' || delimiter == '\n')
            throw new ArgumentOutOfRangeException(nameof(delimiter));
        if (quote == '\r' || quote == '\n')
            throw new ArgumentOutOfRangeException(nameof(quote));
        if (newLine != "\r\n" && newLine != "\n")
            throw new ArgumentOutOfRangeException(nameof(newLine));
        if (encoding == null)
            throw new ArgumentNullException(nameof(encoding));
        if (culture == null)
            throw new ArgumentNullException(nameof(culture));
        if (!Enum.IsDefined(typeof(CsvFormulaInjectionPolicy), formulaInjectionPolicy))
            throw new ArgumentOutOfRangeException(nameof(formulaInjectionPolicy));
    }

    private string FormatValue(ExcelPropertyMap column, object value, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        var context = new ExcelConversionContext(value, column.Name, column.Property.PropertyType, null, rowIndex,
            columnIndex, culture);
        foreach (var converter in ResolveValueConverters(column))
        {
            if (converter.TryConvertTo(context, out var convertedValue))
                return Convert.ToString(convertedValue, culture) ?? string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(column.Formatter) && value is IFormattable formattable)
            return formattable.ToString(column.Formatter, culture);
        var mapping = column.ValueMap.FirstOrDefault(pair => IsMappedValue(pair.Value, value));
        if (mapping.Key != null)
            return mapping.Key;
        return Convert.ToString(value, culture) ?? string.Empty;
    }

    private string FormatValue(CsvExportColumn column, object item, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (!column.IsDynamic)
            return FormatValue(column.Property, column.Property.Getter(item), rowIndex, columnIndex, culture);
        var values = column.Property.Getter(item) as IDictionary<string, object>;
        return values != null && values.TryGetValue(column.Title, out var value)
            ? Convert.ToString(value, culture) ?? string.Empty
            : string.Empty;
    }

    private static IReadOnlyList<CsvExportColumn> CreateColumns<T>(ExcelTypeMap<T> map,
        IReadOnlyList<string> dynamicColumns) where T : class, new()
    {
        var columns = new List<CsvExportColumn>();
        foreach (var property in map.Properties.Where(property => !property.Ignored))
        {
            if (!property.IsDynamicColumn)
                columns.Add(new CsvExportColumn(property, property.Title, false));
            else
                foreach (var title in dynamicColumns ?? Array.Empty<string>())
                {
                    if (string.IsNullOrWhiteSpace(title))
                        throw new ArgumentException("动态列名称不能为空。", nameof(dynamicColumns));
                    columns.Add(new CsvExportColumn(property, title, true));
                }
        }
        if (columns.Select(column => column.Title).Distinct(StringComparer.OrdinalIgnoreCase).Count() != columns.Count)
            throw new ArgumentException("CSV 导出列标题重复。", nameof(dynamicColumns));
        return columns;
    }

    private IEnumerable<IExcelValueConverter> ResolveValueConverters(ExcelPropertyMap property)
    {
        var propertyType = property.Property.PropertyType;
        if (string.IsNullOrWhiteSpace(property.ConverterName))
            return _valueConverters.Where(converter => converter.CanConvert(propertyType));
        var converters = _valueConverters.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, property.ConverterName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (converters.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {property.ConverterName}");
        if (!converters[0].CanConvert(propertyType))
            throw new InvalidOperationException($"值转换器 {property.ConverterName} 不支持属性类型: {propertyType.FullName}");
        return converters;
    }

    private static bool IsMappedValue(object mappedValue, object value)
    {
        if (mappedValue == null || value == null)
            return mappedValue == null && value == null;
        if (Equals(mappedValue, value))
            return true;
        if (value is not Enum enumValue)
            return false;
        return Equals(mappedValue, Convert.ChangeType(enumValue, Enum.GetUnderlyingType(enumValue.GetType()),
            CultureInfo.InvariantCulture));
    }

    private sealed class CsvExportColumn
    {
        public CsvExportColumn(ExcelPropertyMap property, string title, bool isDynamic)
        {
            Property = property;
            Title = title;
            IsDynamic = isDynamic;
        }

        public ExcelPropertyMap Property { get; }
        public string Title { get; }
        public bool IsDynamic { get; }
    }
}

/// <summary>
/// 基于类型映射的 CSV 流式导入器。
/// </summary>
public sealed class CsvEntityImporter : ICsvImporter
{
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityImporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="validationRules">属性校验规则集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    public CsvEntityImporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
    }

    /// <inheritdoc />
    public CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new CsvImportOptions<T>();
        if (options.Delimiter == options.Quote || options.Delimiter == '\r' || options.Delimiter == '\n')
            throw new ArgumentOutOfRangeException(nameof(options.Delimiter));
        if (options.Quote == '\r' || options.Quote == '\n')
            throw new ArgumentOutOfRangeException(nameof(options.Quote));
        if (options.Encoding == null)
            throw new ArgumentNullException(nameof(options.Encoding));
        if (options.Culture == null)
            throw new ArgumentNullException(nameof(options.Culture));

        var map = ExcelTypeMapFactory.Get(options.MappingProfile, options.MappingConfiguration);
        var properties = map.Properties.Where(property => !property.Ignored && !property.IsDynamicColumn).ToList();
        var dynamicProperties = map.Properties.Where(property => !property.Ignored && property.IsDynamicColumn).ToList();
        if (dynamicProperties.Count > 1)
            throw new InvalidOperationException($"CSV 模板 {typeof(T).FullName} 只能声明一个动态列属性。");
        ValidateValidationRuleBindings(properties);
        ValidateNamedValidationRuleBindings(properties);
        using var reader = new StreamReader(source, options.Encoding, true, 1024, true);
        var records = CsvRecordReader.Read(reader, options.Delimiter, options.Quote, cancellationToken).GetEnumerator();
        var columns = options.HasHeader
            ? CreateColumns(records, properties, dynamicProperties, options.HeaderMatch)
            : properties.Select((property, index) => new CsvColumn(index, property, property.Title, false)).ToList();
        var items = new List<T>();
        var errors = new List<CsvImportError>();
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var rowIndex = options.HasHeader ? 1 : 0;
        while (records.MoveNext())
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowIndex++;
            var record = records.Current;
            var item = new T();
            var valid = true;
            var rowDuplicateValues = CloneDuplicateValues(duplicateValues);
            Dictionary<string, object> dynamicValues = null;
            foreach (var column in columns)
            {
                var value = column.Index < record.Count ? record[column.Index] : string.Empty;
                try
                {
                    if (column.IsDynamic)
                    {
                        dynamicValues ??= new Dictionary<string, object>(StringComparer.Ordinal);
                        dynamicValues[column.HeaderName] = value;
                        continue;
                    }
                    ValidateRawValue(value, column, rowIndex, rowDuplicateValues, options.Culture);
                    var converted = ConvertValue(value, column.Property, rowIndex, column.Index + 1, options.Culture);
                    ValidateConvertedValue(value, converted, column, rowIndex, rowDuplicateValues, options.Culture);
                    column.Property.Property.SetValue(item, converted);
                }
                catch (Exception exception)
                {
                    errors.Add(new CsvImportError(exception.Message, rowIndex, column.Index + 1, column.Property.Name));
                    valid = false;
                    break;
                }
            }
            if (valid)
            {
                if (dynamicValues != null)
                    dynamicProperties[0].Property.SetValue(item, dynamicValues);
                items.Add(item);
                duplicateValues = rowDuplicateValues;
            }
        }
        return new CsvImportResult<T>(items, errors);
    }

    private static IReadOnlyList<CsvColumn> CreateColumns(IEnumerator<IReadOnlyList<string>> records,
        IReadOnlyCollection<ExcelPropertyMap> properties, IReadOnlyCollection<ExcelPropertyMap> dynamicProperties,
        bool headerMatch)
    {
        if (!records.MoveNext())
            throw new InvalidOperationException("CSV 不包含表头。");
        var columns = new List<CsvColumn>();
        var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < records.Current.Count; index++)
        {
            var header = records.Current[index];
            if (!headers.Add(header))
                throw new InvalidOperationException($"CSV 包含重复表头: {header}");
            var property = properties.FirstOrDefault(candidate =>
                string.Equals(candidate.Title, header, StringComparison.OrdinalIgnoreCase)
                || string.Equals(candidate.Name, header, StringComparison.OrdinalIgnoreCase));
            if (property == null && dynamicProperties.Count == 1)
                property = dynamicProperties.First();
            if (property != null)
            {
                if (!property.Property.CanWrite)
                    throw new InvalidOperationException($"导入模板属性不可写入: {property.Name}");
                columns.Add(new CsvColumn(index, property, header, property.IsDynamicColumn));
            }
        }
        if (headerMatch)
        {
            var missing = properties.Where(property => columns.All(column => column.Property != property)).Select(property => property.Title);
            if (missing.Any())
                throw new InvalidOperationException($"CSV 不存在列: {string.Join(",", missing)}");
        }
        return columns;
    }

    private object ConvertValue(string value, ExcelPropertyMap property, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var type = property.Property.PropertyType;
        var context = new ExcelConversionContext(value, property.Name, type, null, rowIndex, columnIndex, culture);
        foreach (var converter in ResolveValueConverters(property))
        {
            if (converter.TryConvertFrom(context, out var convertedValue))
                return convertedValue;
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            if (!type.IsValueType || Nullable.GetUnderlyingType(type) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {type.FullName}");
        }
        if (property.ValueMap.TryGetValue(value, out var mappedValue))
            return mappedValue;
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        return Convert.ChangeType(value, targetType, culture);
    }

    private IEnumerable<IExcelValueConverter> ResolveValueConverters(ExcelPropertyMap property)
    {
        var propertyType = property.Property.PropertyType;
        if (string.IsNullOrWhiteSpace(property.ConverterName))
            return _valueConverters.Where(converter => converter.CanConvert(propertyType));
        var converters = _valueConverters.OfType<INamedExcelValueConverter>().Where(converter =>
            string.Equals(converter.Name, property.ConverterName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (converters.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名值转换器: {property.ConverterName}");
        if (!converters[0].CanConvert(propertyType))
            throw new InvalidOperationException($"值转换器 {property.ConverterName} 不支持属性类型: {propertyType.FullName}");
        return converters;
    }

    private INamedExcelValidationRule ResolveNamedValidationRule(string ruleName)
    {
        var rules = _namedValidationRules.Where(rule => string.Equals(rule.Name, ruleName,
            StringComparison.OrdinalIgnoreCase)).ToList();
        if (rules.Count != 1)
            throw new InvalidOperationException($"未找到唯一命名校验规则: {ruleName}");
        return rules[0];
    }

    private void ValidateRawValue(string value, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture)
    {
        var context = CreateValidationContext(value, null, column, rowIndex, duplicateValues, culture);
        foreach (var attribute in column.Property.Property.GetCustomAttributes(true).OfType<FilterAttributeBase>()
                 .Where(IsRawValidationAttribute))
        {
            var rule = ResolveValidationRule(attribute);
            if (!rule.Validate(attribute, context))
                throw new InvalidOperationException(attribute.ErrorMsg);
        }
    }

    private void ValidateConvertedValue(string value, object convertedValue, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture)
    {
        var context = CreateValidationContext(value, convertedValue, column, rowIndex, duplicateValues, culture);
        foreach (var attribute in column.Property.Property.GetCustomAttributes(true).OfType<FilterAttributeBase>()
                 .Where(attribute => !IsRawValidationAttribute(attribute)))
        {
            var rule = ResolveValidationRule(attribute);
            if (!rule.Validate(attribute, context))
                throw new InvalidOperationException(attribute.ErrorMsg);
        }
        foreach (var ruleName in column.Property.ValidationRuleNames)
        {
            var rule = ResolveNamedValidationRule(ruleName);
            if (!rule.Validate(context))
                throw new InvalidOperationException(rule.ErrorMessage);
        }
    }

    private static ExcelValidationContext CreateValidationContext(string value, object convertedValue, CsvColumn column,
        int rowIndex, IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture) => new(value, null,
        rowIndex, column.Index + 1, column.Property.Name, duplicateValues, convertedValue,
        column.Property.Property.PropertyType, new ExcelCellValue(value, value, ExcelCellKind.Text), culture,
        column.Property.Property);

    private IExcelValidationRule ResolveValidationRule(FilterAttributeBase attribute)
    {
        var binding = attribute.GetType().GetCustomAttribute<BindFilterAttribute>();
        if (binding != null)
        {
            var boundRule = _validationRules.FirstOrDefault(rule => binding.RuleType.IsInstanceOfType(rule));
            if (boundRule == null)
                throw new InvalidOperationException($"未注册校验规则: {binding.RuleType.FullName}");
            if (!boundRule.CanValidate(attribute))
                throw new InvalidOperationException($"校验规则 {binding.RuleType.FullName} 不支持特性: {attribute.GetType().FullName}");
            return boundRule;
        }
        var rule = _validationRules.FirstOrDefault(candidate => candidate.CanValidate(attribute));
        return rule ?? throw new InvalidOperationException($"未找到特性对应的校验规则: {attribute.GetType().FullName}");
    }

    private void ValidateValidationRuleBindings(IEnumerable<ExcelPropertyMap> properties)
    {
        foreach (var attribute in properties.SelectMany(property => property.Property.GetCustomAttributes(true)
                     .OfType<FilterAttributeBase>()))
            ResolveValidationRule(attribute);
    }

    private void ValidateNamedValidationRuleBindings(IEnumerable<ExcelPropertyMap> properties)
    {
        foreach (var ruleName in properties.SelectMany(property => property.ValidationRuleNames))
            ResolveNamedValidationRule(ruleName);
    }

    private static bool IsRawValidationAttribute(FilterAttributeBase attribute) =>
        attribute is RequiredAttribute || attribute is RegexAttribute;

    private static Dictionary<string, HashSet<string>> CloneDuplicateValues(
        IReadOnlyDictionary<string, HashSet<string>> duplicateValues)
    {
        var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var pair in duplicateValues)
            result[pair.Key] = new HashSet<string>(pair.Value, StringComparer.OrdinalIgnoreCase);
        return result;
    }

    private sealed class CsvColumn
    {
        public CsvColumn(int index, ExcelPropertyMap property, string headerName, bool isDynamic)
        {
            Index = index;
            Property = property;
            HeaderName = headerName;
            IsDynamic = isDynamic;
        }

        public int Index { get; }
        public ExcelPropertyMap Property { get; }
        public string HeaderName { get; }
        public bool IsDynamic { get; }
    }
}

/// <summary>
/// RFC 4180 风格 CSV 记录读取器。
/// </summary>
internal static class CsvRecordReader
{
    public static IEnumerable<IReadOnlyList<string>> Read(TextReader reader, char delimiter, char quote,
        CancellationToken cancellationToken)
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            Mode = CsvMode.RFC4180,
            BadDataFound = _ => throw new InvalidOperationException("CSV 包含不符合 RFC 4180 的字段。")
        };
        using var parser = new CsvParser(reader, configuration, true);
        while (parser.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return parser.Record;
        }
    }
}

/// <summary>
/// 基于 CsvHelper 的 CSV 记录写入器。
/// </summary>
internal static class CsvRecordWriter
{
    /// <summary>
    /// 将一个记录写入调用方拥有的文本写入器。
    /// </summary>
    /// <param name="writer">目标文本写入器。</param>
    /// <param name="fields">记录字段。</param>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="newLine">记录换行符。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
    public static void Write(TextWriter writer, IEnumerable<string> fields, char delimiter, char quote, string newLine,
        CsvFormulaInjectionPolicy formulaInjectionPolicy)
    {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter.ToString(),
            Quote = quote,
            HasHeaderRecord = false,
            NewLine = newLine
        };
        using var csv = new CsvWriter(writer, configuration, true);
        foreach (var field in fields)
            csv.WriteField(ProtectFormula(field ?? string.Empty, formulaInjectionPolicy));
        csv.NextRecord();
    }

    private static string ProtectFormula(string value, CsvFormulaInjectionPolicy policy) =>
        policy == CsvFormulaInjectionPolicy.Escape && value.Length > 0 && "=+-@".IndexOf(value[0]) >= 0
            ? $"'{value}"
            : value;
}
