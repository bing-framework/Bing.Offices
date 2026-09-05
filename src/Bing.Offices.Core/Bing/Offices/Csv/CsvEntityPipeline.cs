using System.Globalization;
using System.Reflection;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Dates;
using Bing.Offices.Exceptions;
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
internal sealed partial class CsvEntityExporter : ICsvExporter
{
    /// <summary>按优先级用于导出字段的值转换器集合。</summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>将 CSV 请求编译为不可变列映射计划的工厂。</summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>观察并记录公共 CSV 运行异常。</summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    public CsvEntityExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? new ExcelMappingPlanFactory(
            valueConverters: _valueConverters);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public void Export<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        try
        {
            ExportCore(data, destination, options, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesExportException("CSV 导出失败。", exception, "Core",
                BingOfficesStage.Write);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    private void ExportCore<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options,
        CancellationToken cancellationToken) where T : class, new()
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
        IExcelMappingPlan map;
        try
        {
            map = _mappingPlanFactory.Create<T>(document, options.MappingConfiguration, MappingDirection.Export);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException("CSV 映射配置无效。", exception);
        }
        var columns = CreateColumns<T>(map, options.DynamicColumns);
        using var writer = new StreamWriter(destination, options.Encoding, 1024, true);
        if (options.IncludeHeader)
            CsvRecordWriter.Write(writer, columns.Select(column => column.Title), options.Delimiter, options.Quote,
                options.NewLine, options.FormulaInjectionPolicy);
        var rowIndex = options.IncludeHeader ? 2 : 1;
        foreach (var item in data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            CsvRecordWriter.Write(writer, columns.Select((column, index) => FormatValue(column, item, rowIndex,
                index + 1, options.Culture)), options.Delimiter, options.Quote, options.NewLine,
                options.FormulaInjectionPolicy);
            rowIndex++;
        }
        writer.Flush();
    }

    /// <summary>验证 CSV 序列化使用的字符、编码、区域性和公式防护策略。</summary>
    /// <param name="delimiter">字段分隔符。</param>
    /// <param name="quote">字段引用字符。</param>
    /// <param name="newLine">记录换行符。</param>
    /// <param name="encoding">目标流文本编码。</param>
    /// <param name="culture">值格式化使用的区域性。</param>
    /// <param name="formulaInjectionPolicy">潜在公式字段的处理策略。</param>
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

    /// <summary>将固定映射列的实体值格式化为 CSV 字段文本。</summary>
    /// <param name="column">固定列属性绑定。</param>
    /// <param name="value">待格式化的实体值。</param>
    /// <param name="rowIndex">目标记录的一基行号。</param>
    /// <param name="columnIndex">目标字段的一基列号。</param>
    /// <param name="culture">值格式化使用的区域性。</param>
    /// <returns>可写入 CSV 记录的字段文本。</returns>
    private string FormatValue(CsvPropertyBinding column, object value, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (value == null)
            return string.Empty;
        var context = new ExcelConversionContext(value, column.Name, column.Property.PropertyType, null, rowIndex,
            columnIndex, culture);
        foreach (var converter in column.ValueConverters)
        {
            try
            {
                if (converter.TryConvertTo(context, out var convertedValue))
                    return Convert.ToString(convertedValue, culture) ?? string.Empty;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("CSV 值转换器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, propertyName: column.Name, rowIndex: rowIndex,
                    columnIndex: columnIndex, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        if (!string.IsNullOrWhiteSpace(column.Formatter) && value is IFormattable formattable)
            return formattable.ToString(column.Formatter, culture);
        var mapping = column.ValueMap.FirstOrDefault(pair => IsMappedValue(pair.Value, value, culture));
        if (mapping.Key != null)
            return mapping.Key;
        return Convert.ToString(value, culture) ?? string.Empty;
    }

    /// <summary>将固定或动态导出列从实体中读取并格式化为 CSV 字段文本。</summary>
    /// <param name="column">包含映射和动态列定义的导出列。</param>
    /// <param name="item">当前导出实体。</param>
    /// <param name="rowIndex">目标记录的一基行号。</param>
    /// <param name="columnIndex">目标字段的一基列号。</param>
    /// <param name="culture">值格式化使用的区域性。</param>
    /// <returns>可写入 CSV 记录的字段文本。</returns>
    private string FormatValue(CsvExportColumn column, object item, int rowIndex, int columnIndex, CultureInfo culture)
    {
        if (!column.IsDynamic)
            return FormatValue(column.Property, column.Property.Getter(item), rowIndex, columnIndex, culture);
        object rawValue;
        try
        {
            rawValue = column.Property.Getter(item);
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesExportException("CSV 属性读取器执行失败。", exception, "Core",
                BingOfficesStage.Validate, propertyName: column.Property.Name, rowIndex: rowIndex,
                columnIndex: columnIndex, code: BingOfficesErrorCode.UserExtensionFailed);
        }
        var values = rawValue as IDictionary<string, object>;
        if (values == null || !values.TryGetValue(column.DynamicColumn?.Key ?? column.Title, out var value))
            return string.Empty;
        var type = CsvDynamicTypeResolver.Resolve(column.DynamicColumn?.DataTypeName);
        var context = new ExcelConversionContext(value, column.DynamicColumn?.Key ?? column.Title, type,
            null, rowIndex, columnIndex, culture);
        foreach (var converter in column.DynamicColumn?.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            try
            {
                if (converter.TryConvertTo(context, out var convertedValue))
                    return Convert.ToString(convertedValue, culture) ?? string.Empty;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("CSV 动态值转换器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, propertyName: column.Property.Name, rowIndex: rowIndex,
                    columnIndex: columnIndex, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        return Convert.ToString(value, culture) ?? string.Empty;
    }

    /// <summary>将不可变映射计划展开为固定和动态 CSV 导出列。</summary>
    /// <typeparam name="T">导出实体类型。</typeparam>
    /// <param name="map">已编译的实体映射计划。</param>
    /// <param name="dynamicColumns">映射未声明动态列计划时使用的请求级标题。</param>
    /// <returns>按输出顺序排列的 CSV 导出列。</returns>
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

    /// <summary>比较配置映射值与实体值的文本表示。</summary>
    /// <param name="mappedValue">映射配置中的文本值。</param>
    /// <param name="value">待比较的实体值。</param>
    /// <param name="culture">格式化实体值使用的区域性。</param>
    /// <returns>两个值语义相等时为 true。</returns>
    private static bool IsMappedValue(string mappedValue, object value, CultureInfo culture)
    {
        if (mappedValue == null || value == null)
            return mappedValue == null && value == null;
        return string.Equals(mappedValue, Convert.ToString(value, culture), StringComparison.Ordinal);
    }

    private sealed class CsvExportColumn
    {
        /// <summary>使用属性绑定和可选动态列计划创建导出列。</summary>
        /// <param name="property">实体属性绑定。</param>
        /// <param name="title">输出表头标题。</param>
        /// <param name="isDynamic">是否从实体动态字典中读取值。</param>
        /// <param name="dynamicColumn">已绑定的动态列计划。</param>
        public CsvExportColumn(CsvPropertyBinding property, string title, bool isDynamic,
            IExcelDynamicMappingColumn dynamicColumn = null)
        {
            Property = property;
            Title = title;
            IsDynamic = isDynamic;
            DynamicColumn = dynamicColumn;
        }

        /// <summary>获取实体属性绑定。</summary>
        public CsvPropertyBinding Property { get; }
        /// <summary>获取输出表头标题。</summary>
        public string Title { get; }
        /// <summary>获取是否从实体动态字典中读取值。</summary>
        public bool IsDynamic { get; }
        /// <summary>获取已绑定的动态列计划；固定列时为 null。</summary>
        public IExcelDynamicMappingColumn DynamicColumn { get; }
    }
}
/// <summary>
/// 基于类型映射的 CSV 流式导入器。
/// </summary>
internal sealed partial class CsvEntityImporter : ICsvImporter
{
    /// <summary>按优先级用于导入字段的值转换器集合。</summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>用于属性特性校验的规则集合。</summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    /// <summary>可按配置名称绑定的校验规则集合。</summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>将 CSV 请求编译为不可变列映射计划的工厂。</summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>观察并记录公共 CSV 运行异常。</summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityImporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="validationRules">属性校验规则集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    public CsvEntityImporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? new ExcelMappingPlanFactory(
            valueConverters: _valueConverters,
            validationRules: _validationRules,
            namedValidationRules: _namedValidationRules);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public CsvImportResult<T> Import<T>(Stream source, CsvImportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        try
        {
            return ImportCore(source, options, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("CSV 导入失败。", exception, "Core",
                BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    private CsvImportResult<T> ImportCore<T>(Stream source, CsvImportOptions<T> options,
        CancellationToken cancellationToken) where T : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        options ??= new CsvImportOptions<T>();
        options.Validate();
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
        if (options.MaxInputBytes.HasValue && source.CanSeek && source.Length - source.Position > options.MaxInputBytes.Value)
            return CreateResourceLimitResult($"CSV 输入超过最大字节数: {options.MaxInputBytes.Value}", options);
        using var limitedSource = options.MaxInputBytes.HasValue && !source.CanSeek
            ? new CsvLimitedReadStream(source, options.MaxInputBytes.Value)
            : null;
        using var reader = new StreamReader(limitedSource ?? source, options.Encoding, true, 1024, true);
        using var records = CsvRecordReader.Read(reader, options.Delimiter, options.Quote, cancellationToken).GetEnumerator();
        IReadOnlyList<CsvColumn> columns;
        try
        {
            columns = options.HasHeader
                ? CsvHeaderBinder.Bind(records, properties, dynamicProperties, map.DynamicColumns, options.RequireExpectedHeaders,
                    options.MaxColumns)
                : CsvHeaderBinder.BindByPosition(properties);
        }
        catch (CsvResourceLimitException exception)
        {
            return CreateResourceLimitResult(exception.Message, options);
        }
        catch (CsvInvalidHeaderException exception)
        {
            return new CsvImportResult<T>(Array.Empty<T>(), new[]
            {
                new CsvImportError(exception.Message, 1, 0, null, code: CsvImportErrorCode.InvalidHeader)
            }, maxErrors: options.MaxErrors);
        }
        if (options.MaxColumns.HasValue && columns.Count > options.MaxColumns.Value)
            return CreateResourceLimitResult($"CSV 映射列数超过最大列数: {options.MaxColumns.Value}", options);
        var items = new List<T>();
        var errors = new List<CsvImportError>();
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var uniqueTracker = new UniqueTracker(duplicateValues, options.MaxTrackedUniqueValues,
            CreateStringComparer(options.UniqueComparison));
        var rowIndex = options.HasHeader ? 1 : 0;
        var dataRowCount = 0;
        var isTruncated = false;
        try
        {
            while (records.MoveNext())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (options.MaxRows.HasValue && dataRowCount >= options.MaxRows.Value)
                {
                    errors.Add(new CsvImportError($"CSV 数据行数超过限制: {options.MaxRows.Value}", rowIndex + 1, 0,
                        null, code: CsvImportErrorCode.ResourceLimit));
                    isTruncated = true;
                    break;
                }
                dataRowCount++;
                rowIndex++;
                var record = records.Current;
                if (options.MaxColumns.HasValue && record.Count > options.MaxColumns.Value)
                {
                    errors.Add(new CsvImportError($"CSV 第 {rowIndex} 行超过最大列数: {options.MaxColumns.Value}",
                        rowIndex, 0, null, code: CsvImportErrorCode.ResourceLimit));
                    isTruncated = true;
                    break;
                }
                var item = new T();
                var valid = true;
                uniqueTracker.BeginRow();
                Dictionary<string, object> dynamicValues = null;
                foreach (var column in columns)
                {
                    var value = column.Index < record.Count ? record[column.Index] : string.Empty;
                    if (options.MaxFieldLength.HasValue && value.Length > options.MaxFieldLength.Value)
                    {
                        errors.Add(new CsvImportError(
                            $"CSV 第 {rowIndex} 行第 {column.Index + 1} 列超过最大字段长度: {options.MaxFieldLength.Value}",
                            rowIndex, column.Index + 1, column.DynamicColumn?.Key ?? column.Property.Name,
                            code: CsvImportErrorCode.ResourceLimit));
                        valid = false;
                        isTruncated = true;
                        break;
                    }
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
                        try
                        {
                            column.Property.Setter(item, converted);
                        }
                        catch (Exception exception) when (exception is not OperationCanceledException
                            && exception is not OutOfMemoryException && exception is not StackOverflowException)
                        {
                            throw new BingOfficesImportException("CSV 属性写入器执行失败。", exception, "Core",
                                BingOfficesStage.Validate, rowIndex: rowIndex,
                                columnIndex: column.Index + 1, propertyName: column.Property.Name,
                                code: BingOfficesErrorCode.UserExtensionFailed);
                        }
                    }
                    catch (BingOfficesException)
                    {
                        throw;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException
                        && exception is not StackOverflowException)
                    {
                        int? firstRowNumber = null;
                        var errorColumnKey = column.DynamicColumn?.Key ?? column.Property.Name;
                        if (uniqueTracker.TryGetFirstRowNumber(errorColumnKey, value, out var firstRow))
                            firstRowNumber = firstRow;
                        errors.Add(new CsvImportError(exception.Message, rowIndex, column.Index + 1,
                            errorColumnKey, firstRowNumber, ClassifyError(exception)));
                        valid = false;
                        break;
                    }
                }
                if (valid)
                {
                    if (dynamicValues != null)
                    {
                        try
                        {
                            dynamicProperties[0].Setter(item, dynamicValues);
                        }
                        catch (Exception exception) when (exception is not OperationCanceledException
                            && exception is not OutOfMemoryException && exception is not StackOverflowException)
                        {
                            throw new BingOfficesImportException("CSV 动态属性写入器执行失败。", exception, "Core",
                                BingOfficesStage.Validate, rowIndex: rowIndex,
                                propertyName: dynamicProperties[0].Name,
                                code: BingOfficesErrorCode.UserExtensionFailed);
                        }
                    }
                    items.Add(item);
                    uniqueTracker.CommitRow();
                }
                else
                    uniqueTracker.RollbackRow();
                if (isTruncated || options.MaxErrors.HasValue && errors.Count >= options.MaxErrors.Value)
                {
                    isTruncated = true;
                    break;
                }
            }
        }
        catch (CsvResourceLimitException exception)
        {
            errors.Add(new CsvImportError(exception.Message, rowIndex + 1, 0, null,
                code: CsvImportErrorCode.ResourceLimit));
            isTruncated = true;
        }
        return new CsvImportResult<T>(items, errors, isTruncated, options.MaxErrors);
    }

    /// <summary>创建表示输入资源限制已触发的截断导入结果。</summary>
    /// <typeparam name="T">导入实体类型。</typeparam>
    /// <param name="message">描述超出资源限制的错误消息。</param>
    /// <param name="options">提供最大错误数的当前导入选项。</param>
    /// <returns>不包含实体且带有资源限制错误的截断结果。</returns>
    private static CsvImportResult<T> CreateResourceLimitResult<T>(string message, CsvImportOptions<T> options)
        where T : class, new() => new CsvImportResult<T>(Array.Empty<T>(), new[]
        {
            new CsvImportError(message, 0, 0, null, code: CsvImportErrorCode.ResourceLimit)
        }, true, options.MaxErrors);

    /// <summary>将转换和校验异常映射为公开的 CSV 错误代码。</summary>
    /// <param name="exception">处理字段时捕获的异常。</param>
    /// <returns>用于导入结果的错误代码。</returns>
    private static CsvImportErrorCode ClassifyError(Exception exception)
    {
        if (exception is InvalidCastException || exception is FormatException
            || exception is OverflowException || exception is ArgumentException)
            return CsvImportErrorCode.ValueConversion;
        if (exception is InvalidOperationException)
            return CsvImportErrorCode.Validation;
        return CsvImportErrorCode.InvalidInput;
    }

    /// <summary>将固定列文本转换为实体属性的目标类型。</summary>
    /// <param name="value">规范化后的 CSV 字段文本。</param>
    /// <param name="property">固定列属性绑定。</param>
    /// <param name="rowIndex">字段所在的一基行号。</param>
    /// <param name="columnIndex">字段所在的一基列号。</param>
    /// <param name="culture">文本转换使用的区域性。</param>
    /// <returns>可写入实体属性的转换结果。</returns>
    private object ConvertValue(string value, CsvPropertyBinding property, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var type = property.Property.PropertyType;
        var context = new ExcelConversionContext(value, property.Name, type, null, rowIndex, columnIndex, culture);
        foreach (var converter in property.ValueConverters)
        {
            try
            {
                if (converter.TryConvertFrom(context, out var convertedValue))
                    return convertedValue;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("CSV 值转换器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, rowIndex: rowIndex, columnIndex: columnIndex,
                    propertyName: property.Name, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            if (!type.IsValueType || Nullable.GetUnderlyingType(type) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {type.FullName}");
        }
        if (property.ValueMap.TryGetValue(value, out var mappedValue))
            return ConvertMappedValue(mappedValue, type, culture, property.Attributes.OfType<ExcelDateAttribute>()
                .FirstOrDefault());
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType.IsEnum)
            return Enum.Parse(targetType, value, true);
        if (targetType == typeof(Guid))
            return Guid.Parse(value);
        if (targetType == typeof(Version))
            return new Version(value);
        if (targetType == typeof(DateTime) || targetType == typeof(DateTimeOffset))
        {
            if (ExcelDateParser.TryParse(new ExcelCellValue(value, value, ExcelCellKind.Text), value,
                    type, culture, property.Attributes.OfType<ExcelDateAttribute>().FirstOrDefault(), out var parsed))
                return parsed;
            throw new FormatException($"值不是受支持的日期格式: {value}");
        }
        return Convert.ChangeType(value, targetType, culture);
    }

    /// <summary>将动态列文本转换为其配置的逻辑类型。</summary>
    /// <param name="value">规范化后的 CSV 字段文本。</param>
    /// <param name="column">动态列的表头和映射绑定。</param>
    /// <param name="rowIndex">字段所在的一基行号。</param>
    /// <param name="culture">文本转换使用的区域性。</param>
    /// <returns>适合写入动态值字典的转换结果。</returns>
    private static object ConvertDynamicValue(string value, CsvColumn column, int rowIndex, CultureInfo culture)
    {
        if (column.DynamicColumn == null)
            return value;
        var type = CsvDynamicTypeResolver.Resolve(column.DynamicColumn.DataTypeName);
        var context = new ExcelConversionContext(value, column.DynamicColumn.Key, type, null, rowIndex,
            column.Index + 1, culture, new ExcelCellValue(value, value, ExcelCellKind.Text));
        foreach (var converter in column.DynamicColumn.ValueConverters)
        {
            try
            {
                if (converter.TryConvertFrom(context, out var convertedValue))
                    return convertedValue;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("CSV 动态值转换器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, rowIndex: rowIndex, columnIndex: column.Index + 1,
                    propertyName: column.DynamicColumn.Key, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            if (!type.IsValueType || Nullable.GetUnderlyingType(type) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {type.FullName}");
        }
        var targetType = Nullable.GetUnderlyingType(type) ?? type;
        if (targetType == typeof(Guid)) return Guid.Parse(value);
        if (targetType == typeof(DateTime) || targetType == typeof(DateTimeOffset))
        {
            if (ExcelDateParser.TryParse(new ExcelCellValue(value, value, ExcelCellKind.Text), value,
                    targetType, culture, null, out var parsed))
                return parsed;
            throw new FormatException($"值不是受支持的日期格式: {value}");
        }
        return Convert.ChangeType(value, targetType, culture);
    }

    /// <summary>执行需要原始字段文本的校验规则。</summary>
    /// <param name="value">规范化后的 CSV 字段文本。</param>
    /// <param name="column">当前列绑定。</param>
    /// <param name="rowIndex">字段所在的一基行号。</param>
    /// <param name="duplicateValues">兼容旧校验规则的重复值状态。</param>
    /// <param name="culture">校验上下文使用的区域性。</param>
    private void ValidateRawValue(string value, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture)
    {
        var context = CreateValidationContext(value, null, column, rowIndex, duplicateValues, culture);
        foreach (var binding in GetValidationBindings(column).Where(binding => binding.IsRaw))
        {
            try
            {
                var isValid = binding.Validate(context);
                if (!isValid)
                    throw new ValidationFailedException(binding.ErrorMessage);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (ValidationFailedException exception)
            {
                throw new InvalidOperationException(exception.Message);
            }
            catch (Exception exception) when (binding.Kind == ExcelValidationBindingKind.Custom
                && exception is not OperationCanceledException && exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("CSV 自定义校验器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, rowIndex: rowIndex, columnIndex: column.Index + 1,
                    propertyName: column.DynamicColumn?.Key ?? column.Property.Name,
                    code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
    }

    /// <summary>执行转换后校验规则并为唯一性规则预留当前值。</summary>
    /// <param name="value">规范化后的原始字段文本。</param>
    /// <param name="convertedValue">转换后的字段值。</param>
    /// <param name="column">当前列绑定。</param>
    /// <param name="rowIndex">字段所在的一基行号。</param>
    /// <param name="duplicateValues">兼容旧校验规则的重复值状态。</param>
    /// <param name="uniqueTracker">负责当前行提交或回滚的唯一值跟踪器。</param>
    /// <param name="culture">校验上下文使用的区域性。</param>
    private void ValidateConvertedValue(string value, object convertedValue, CsvColumn column, int rowIndex,
        IDictionary<string, HashSet<string>> duplicateValues, UniqueTracker uniqueTracker, CultureInfo culture)
    {
        var context = CreateValidationContext(value, convertedValue, column, rowIndex, duplicateValues, culture);
        foreach (var binding in GetValidationBindings(column).Where(binding => !binding.IsRaw
                     && binding.Kind != ExcelValidationBindingKind.Unique))
        {
            try
            {
                var isValid = binding.Validate(context);
                if (!isValid)
                    throw new ValidationFailedException(binding.ErrorMessage);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (ValidationFailedException exception)
            {
                throw new InvalidOperationException(exception.Message);
            }
            catch (Exception exception) when (binding.Kind == ExcelValidationBindingKind.Custom
                && exception is not OperationCanceledException && exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                throw new BingOfficesImportException("CSV 自定义校验器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, rowIndex: rowIndex, columnIndex: column.Index + 1,
                    propertyName: column.DynamicColumn?.Key ?? column.Property.Name,
                    code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        var uniqueKey = column.DynamicColumn?.Key ?? column.Property.Name;
        var isUnique = column.DynamicColumn?.IsUnique ?? column.Property.IsUnique;
        var ignoreEmpty = column.DynamicColumn?.UniqueIgnoreEmpty ?? column.Property.UniqueIgnoreEmpty;
        if (isUnique && !uniqueTracker.TryReserve(uniqueKey, value, false, ignoreEmpty, rowIndex))
            throw new InvalidOperationException("重复数据");
    }

    /// <summary>创建供 CSV 字段校验规则使用的提供程序无关上下文。</summary>
    /// <param name="value">原始字段文本。</param>
    /// <param name="convertedValue">转换后的字段值。</param>
    /// <param name="column">当前列绑定。</param>
    /// <param name="rowIndex">字段所在的一基行号。</param>
    /// <param name="duplicateValues">兼容旧校验规则的重复值状态。</param>
    /// <param name="culture">校验上下文使用的区域性。</param>
    /// <returns>包含字段位置、类型和文本值的校验上下文。</returns>
    private static ExcelValidationContext CreateValidationContext(string value, object convertedValue, CsvColumn column,
        int rowIndex, IDictionary<string, HashSet<string>> duplicateValues, CultureInfo culture) => new(value, null,
        rowIndex, column.Index + 1, column.DynamicColumn?.Key ?? column.Property.Name, convertedValue,
        column.DynamicColumn == null ? column.Property.Property.PropertyType
            : CsvDynamicTypeResolver.Resolve(column.DynamicColumn.DataTypeName),
        new ExcelCellValue(value, value, ExcelCellKind.Text), culture);

    /// <summary>获取固定列或动态列已经绑定的校验规则。</summary>
    /// <param name="column">当前 CSV 列绑定。</param>
    /// <returns>按执行顺序排列的校验规则集合。</returns>
    private static IReadOnlyList<IExcelValidationBinding> GetValidationBindings(CsvColumn column) =>
        column.DynamicColumn?.ValidationBindings ?? column.Property.ValidationBindings;

    private sealed class ValidationFailedException : Exception
    {
        public ValidationFailedException(string message) : base(message)
        {
        }
    }

    /// <summary>将配置的显示值映射文本转换为目标属性类型。</summary>
    /// <param name="value">映射配置中的文本值。</param>
    /// <param name="type">目标属性类型。</param>
    /// <param name="culture">文本转换使用的区域性。</param>
    /// <param name="dateAttribute">目标属性上的日期输入配置。</param>
    /// <returns>目标类型的映射值。</returns>
    private static object ConvertMappedValue(string value, Type type, CultureInfo culture,
        ExcelDateAttribute dateAttribute = null)
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
        if (targetType == typeof(DateTime) || targetType == typeof(DateTimeOffset))
        {
            if (ExcelDateParser.TryParse(new ExcelCellValue(value, value, ExcelCellKind.Text), value,
                    type, culture, dateAttribute, out var parsed))
                return parsed;
            throw new FormatException($"值不是受支持的日期格式: {value}");
        }
        return Convert.ChangeType(value, targetType, culture);
    }

    /// <summary>按照列配置的空白策略规范化 CSV 字段文本。</summary>
    /// <param name="value">待规范化的文本；为 null 时按空字符串处理。</param>
    /// <param name="policy">空白处理策略；为 null 时保留原始文本。</param>
    /// <returns>规范化后的字段文本。</returns>
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

    /// <summary>创建与导入选项中字符串比较规则等效的比较器。</summary>
    /// <param name="comparison">要支持的字符串比较规则。</param>
    /// <returns>对应的字符串比较器。</returns>
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
