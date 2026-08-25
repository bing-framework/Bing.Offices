using System.Globalization;
using System.Reflection;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
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
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public CsvEntityExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? new ExcelMappingPlanFactory(
            valueConverters: _valueConverters);
    }

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
        var document = options.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        };
        var map = _mappingPlanFactory.Create<T>(document, options.MappingConfiguration, MappingDirection.Export);
        var columns = CreateColumns<T>(map,
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

    private string FormatValue(CsvPropertyBinding column, object value, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        var context = new ExcelConversionContext(value, column.Name, column.Property.PropertyType, null, rowIndex,
            columnIndex, culture);
        foreach (var converter in column.ValueConverters)
        {
            if (converter.TryConvertTo(context, out var convertedValue))
                return Convert.ToString(convertedValue, culture) ?? string.Empty;
        }
        if (!string.IsNullOrWhiteSpace(column.Formatter) && value is IFormattable formattable)
            return formattable.ToString(column.Formatter, culture);
        var mapping = column.ValueMap.FirstOrDefault(pair => IsMappedValue(pair.Value, value, culture));
        if (mapping.Key != null)
            return mapping.Key;
        return Convert.ToString(value, culture) ?? string.Empty;
    }

    private string FormatValue(CsvExportColumn column, object item, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (!column.IsDynamic)
            return FormatValue(column.Property, column.Property.Getter(item), rowIndex, columnIndex, culture);
        var values = column.Property.Getter(item) as IDictionary<string, object>;
        if (values == null || !values.TryGetValue(column.DynamicColumn?.Key ?? column.Title, out var value))
            return string.Empty;
        var type = CsvDynamicTypeResolver.Resolve(column.DynamicColumn?.DataTypeName);
        var context = new ExcelConversionContext(value, column.DynamicColumn?.Key ?? column.Title, type,
            null, rowIndex, columnIndex, culture);
        foreach (var converter in column.DynamicColumn?.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            if (converter.TryConvertTo(context, out var convertedValue))
                return Convert.ToString(convertedValue, culture) ?? string.Empty;
        }
        return Convert.ToString(value, culture) ?? string.Empty;
    }

    private static IReadOnlyList<CsvExportColumn> CreateColumns<T>(IExcelMappingPlan map,
        IReadOnlyList<string> dynamicColumns) where T : class, new()
    {
        var columns = new List<CsvExportColumn>();
        foreach (var property in map.Columns.Where(property => !property.Ignored))
        {
            var binding = CsvPropertyBinding.Create<T>(property);
            if (!property.IsDynamicColumn)
                columns.Add(new CsvExportColumn(binding, property.Title, false));
            else if (map.DynamicColumns.Count > 0)
                foreach (var dynamicColumn in map.DynamicColumns.OrderBy(column => column.Order)
                             .ThenBy(column => column.Key, StringComparer.Ordinal))
                    columns.Add(new CsvExportColumn(binding, dynamicColumn.Title, true, dynamicColumn));
            else
                foreach (var title in dynamicColumns ?? Array.Empty<string>())
                {
                    if (string.IsNullOrWhiteSpace(title))
                        throw new ArgumentException("动态列名称不能为空。", nameof(dynamicColumns));
                    columns.Add(new CsvExportColumn(binding, title, true));
                }
        }
        if (columns.Select(column => column.Title).Distinct(StringComparer.OrdinalIgnoreCase).Count() != columns.Count)
            throw new ArgumentException("CSV 导出列标题重复。", nameof(dynamicColumns));
        return columns;
    }

    private static bool IsMappedValue(string mappedValue, object value, CultureInfo culture)
    {
        if (mappedValue == null || value == null)
            return mappedValue == null && value == null;
        return string.Equals(mappedValue, Convert.ToString(value, culture), StringComparison.Ordinal);
    }

    private sealed class CsvExportColumn
    {
        public CsvExportColumn(CsvPropertyBinding property, string title, bool isDynamic,
            IExcelDynamicMappingColumn dynamicColumn = null)
        {
            Property = property;
            Title = title;
            IsDynamic = isDynamic;
            DynamicColumn = dynamicColumn;
        }

        public CsvPropertyBinding Property { get; }
        public string Title { get; }
        public bool IsDynamic { get; }
        public IExcelDynamicMappingColumn DynamicColumn { get; }
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
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityImporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="validationRules">属性校验规则集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    public CsvEntityImporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? new ExcelMappingPlanFactory(
            valueConverters: _valueConverters,
            validationRules: _validationRules,
            namedValidationRules: _namedValidationRules);
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

        var document = options.MappingDocument ?? new ExcelMappingDocument
        {
            UseConventionFallback = true
        };
        var map = _mappingPlanFactory.Create<T>(document, options.MappingConfiguration, MappingDirection.Import);
        var properties = map.Columns.Where(property => !property.Ignored && !property.IsDynamicColumn)
            .Select(CsvPropertyBinding.Create<T>).ToList();
        var dynamicProperties = map.Columns.Where(property => !property.Ignored && property.IsDynamicColumn)
            .Select(CsvPropertyBinding.Create<T>).ToList();
        if (dynamicProperties.Count > 1)
            throw new InvalidOperationException($"CSV 模板 {typeof(T).FullName} 只能声明一个动态列属性。");
        using var reader = new StreamReader(source, options.Encoding, true, 1024, true);
        var records = CsvRecordReader.Read(reader, options.Delimiter, options.Quote, cancellationToken).GetEnumerator();
        var columns = options.HasHeader
            ? CsvHeaderBinder.Bind(records, properties, dynamicProperties, map.DynamicColumns, options.HeaderMatch)
            : CsvHeaderBinder.BindByPosition(properties);
        var items = new List<T>();
        var errors = new List<CsvImportError>();
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var uniqueTracker = new UniqueTracker(duplicateValues, options.MaxTrackedUniqueValues,
            CreateStringComparer(options.UniqueComparison));
        var rowIndex = options.HasHeader ? 1 : 0;
        while (records.MoveNext())
        {
            cancellationToken.ThrowIfCancellationRequested();
            rowIndex++;
            var record = records.Current;
            var item = new T();
            var valid = true;
            uniqueTracker.BeginRow();
            Dictionary<string, object> dynamicValues = null;
            foreach (var column in columns)
            {
                var value = column.Index < record.Count ? record[column.Index] : string.Empty;
                value = NormalizeText(value, column.Property.ImportWhitespace);
                try
                {
                    if (column.IsDynamic)
                    {
                        dynamicValues ??= new Dictionary<string, object>(StringComparer.Ordinal);
                        ValidateRawValue(value, column, rowIndex, duplicateValues, options.Culture);
                        var dynamicValue = ConvertDynamicValue(value, column, rowIndex, options.Culture);
                        ValidateConvertedValue(value, dynamicValue, column, rowIndex, duplicateValues,
                            uniqueTracker, options.Culture);
                        dynamicValues[column.DynamicColumn?.Key ?? column.HeaderName] = dynamicValue;
                        continue;
                    }
                    ValidateRawValue(value, column, rowIndex, duplicateValues, options.Culture);
                    var converted = ConvertValue(value, column.Property, rowIndex, column.Index + 1, options.Culture);
                    ValidateConvertedValue(value, converted, column, rowIndex, duplicateValues, uniqueTracker,
                        options.Culture);
                    column.Property.Setter(item, converted);
                }
                catch (Exception exception)
                {
                    int? firstRowNumber = null;
                    var errorColumnKey = column.DynamicColumn?.Key ?? column.Property.Name;
                    if (uniqueTracker.TryGetFirstRowNumber(errorColumnKey, value, out var firstRow))
                        firstRowNumber = firstRow;
                    errors.Add(new CsvImportError(exception.Message, rowIndex, column.Index + 1,
                        errorColumnKey, firstRowNumber));
                    valid = false;
                    break;
                }
            }
            if (valid)
            {
                if (dynamicValues != null)
                    dynamicProperties[0].Setter(item, dynamicValues);
                items.Add(item);
                uniqueTracker.CommitRow();
            }
            else
                uniqueTracker.RollbackRow();
        }
        return new CsvImportResult<T>(items, errors);
    }

    private object ConvertValue(string value, CsvPropertyBinding property, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var type = property.Property.PropertyType;
        var context = new ExcelConversionContext(value, property.Name, type, null, rowIndex, columnIndex, culture);
        foreach (var converter in property.ValueConverters)
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
            return ConvertMappedValue(mappedValue, type, culture);
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        return Convert.ChangeType(value, targetType, culture);
    }

    private static object ConvertDynamicValue(string value, CsvColumn column, int rowIndex, CultureInfo culture)
    {
        if (column.DynamicColumn == null)
            return value;
        var type = CsvDynamicTypeResolver.Resolve(column.DynamicColumn.DataTypeName);
        var context = new ExcelConversionContext(value, column.DynamicColumn.Key, type, null, rowIndex,
            column.Index + 1, culture, new ExcelCellValue(value, value, ExcelCellKind.Text));
        foreach (var converter in column.DynamicColumn.ValueConverters)
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
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType == typeof(Guid)) return Guid.Parse(value);
        if (targetType == typeof(DateTime)) return DateTime.Parse(value, culture);
        if (targetType == typeof(DateTimeOffset)) return DateTimeOffset.Parse(value, culture);
        return Convert.ChangeType(value, targetType, culture);
    }

    private void ValidateRawValue(string value, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture)
    {
        var context = CreateValidationContext(value, null, column, rowIndex, duplicateValues, culture);
        foreach (var binding in GetValidationBindings(column).Where(binding => binding.IsRaw))
        {
            if (!binding.Validate(context))
                throw new InvalidOperationException(binding.ErrorMessage);
        }
    }

    private void ValidateConvertedValue(string value, object convertedValue, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, UniqueTracker uniqueTracker, CultureInfo culture)
    {
        var context = CreateValidationContext(value, convertedValue, column, rowIndex, duplicateValues, culture);
        foreach (var binding in GetValidationBindings(column).Where(binding => !binding.IsRaw
                     && binding.Kind != ExcelValidationBindingKind.Unique))
        {
            if (!binding.Validate(context))
                throw new InvalidOperationException(binding.ErrorMessage);
        }
        var uniqueKey = column.DynamicColumn?.Key ?? column.Property.Name;
        var isUnique = column.DynamicColumn?.IsUnique ?? column.Property.IsUnique;
        var ignoreEmpty = column.DynamicColumn?.UniqueIgnoreEmpty ?? column.Property.UniqueIgnoreEmpty;
        if (isUnique && !uniqueTracker.TryReserve(uniqueKey, value, false, ignoreEmpty, rowIndex))
            throw new InvalidOperationException("重复数据");
    }

    private static ExcelValidationContext CreateValidationContext(string value, object convertedValue, CsvColumn column,
        int rowIndex, IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture) => new(value, null,
        rowIndex, column.Index + 1, column.DynamicColumn?.Key ?? column.Property.Name, convertedValue,
        column.DynamicColumn == null ? column.Property.Property.PropertyType
            : CsvDynamicTypeResolver.Resolve(column.DynamicColumn.DataTypeName),
        new ExcelCellValue(value, value, ExcelCellKind.Text), culture);

    private static IReadOnlyList<IExcelValidationBinding> GetValidationBindings(CsvColumn column) =>
        column.DynamicColumn?.ValidationBindings ?? column.Property.ValidationBindings;

    private static object ConvertMappedValue(string value, Type type, CultureInfo culture)
    {
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType == typeof(string))
            return value;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        if (targetType == typeof(DateTime))
            return DateTime.Parse(value, culture);
        return Convert.ChangeType(value, targetType, culture);
    }

    private static string NormalizeText(string value, ExcelWhitespacePolicy? policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            null or ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => new string(value.Where(character => !char.IsWhiteSpace(character)).ToArray()),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    private static StringComparer CreateStringComparer(StringComparison comparison) => comparison switch
    {
        StringComparison.Ordinal => StringComparer.Ordinal,
        StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
        StringComparison.InvariantCulture => StringComparer.InvariantCulture,
        StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
        StringComparison.CurrentCulture => StringComparer.CurrentCulture,
        StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
        _ => throw new ArgumentOutOfRangeException(nameof(comparison))
    };

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
