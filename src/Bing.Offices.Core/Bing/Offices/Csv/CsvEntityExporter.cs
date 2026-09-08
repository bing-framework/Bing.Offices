using System.Globalization;
using System.Text;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;

#nullable enable annotations

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
    /// <summary>负责原子文件提交的可替换 SPI。</summary>
    private readonly IFileExportCommitter _fileExportCommitter;
    /// <summary>观察并记录公共 CSV 运行异常。</summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个<see cref="CsvEntityExporter"/>类型的实例。
    /// </summary>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    /// <param name="fileExportCommitter">可替换的原子文件提交器。</param>
    public CsvEntityExporter(IEnumerable<IExcelValueConverter> valueConverters = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null,
        IFileExportCommitter fileExportCommitter = null)
    {
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _mappingPlanFactory = mappingPlanFactory ?? new ExcelMappingPlanFactory(
            valueConverters: _valueConverters);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
    }

    /// <inheritdoc />
    public void Export<T>(IEnumerable<T> data, Stream destination, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        try
        {
            ExportCore(data, destination, options, cancellationToken);
        }
        catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested
            && exception.GetType() != typeof(OperationCanceledException))
        {
            throw new OperationCanceledException(cancellationToken);
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

    /// <inheritdoc />
    public void ExportToFile<T>(IEnumerable<T> data, string path, CsvExportOptions<T> options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));

        try
        {
            _fileExportCommitter.Commit(path,
                destination => Export(data, destination, options, cancellationToken),
                cancellationToken, "CSV");
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
    }

    /// <inheritdoc />
    public async Task ExportAsync<T>(IEnumerable<T> data, Stream destination,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        try
        {
            await ExportCoreAsync(data, destination, options, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested
            && exception.GetType() != typeof(OperationCanceledException))
        {
            throw new OperationCanceledException(cancellationToken);
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

    /// <inheritdoc />
    public async Task ExportToFileAsync<T>(IEnumerable<T> data, string path,
        CsvExportOptions<T> options = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));

        try
        {
            await _fileExportCommitter.CommitAsync(path,
                (destination, token) => ExportAsync(data, destination, options, token),
                cancellationToken, "CSV").ConfigureAwait(false);
        }
        catch (BingOfficesFileCommitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
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

    private async Task ExportCoreAsync<T>(IEnumerable<T> data, Stream destination,
        CsvExportOptions<T> options, CancellationToken cancellationToken) where T : class, new()
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
        using var cancellationDestination = new CsvCancellationStream(destination, cancellationToken,
            suppressSynchronousFlush: true);
        using var writer = new StreamWriter(cancellationDestination, options.Encoding, 1024, true);
        if (options.IncludeHeader)
            await CsvRecordWriter.WriteAsync(writer, columns.Select(column => column.Title), options.Delimiter,
                options.Quote, options.NewLine, options.FormulaInjectionPolicy, cancellationToken)
                .ConfigureAwait(false);
        var rowIndex = options.IncludeHeader ? 2 : 1;
        foreach (var item in data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            await CsvRecordWriter.WriteAsync(writer, columns.Select((column, index) => FormatValue(column, item,
                    rowIndex, index + 1, options.Culture)), options.Delimiter, options.Quote, options.NewLine,
                options.FormulaInjectionPolicy, cancellationToken).ConfigureAwait(false);
            rowIndex++;
        }
        await writer.FlushAsync().ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
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
                    return FormatScalarValue(convertedValue, culture);
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
        return FormatScalarValue(value, culture);
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
                    return FormatScalarValue(convertedValue, culture);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesExportException("CSV 动态值转换器执行失败。", exception, "Core",
                    BingOfficesStage.Validate, propertyName: column.Property.Name, rowIndex: rowIndex,
                    columnIndex: columnIndex, code: BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        return FormatScalarValue(value, culture);
    }

    /// <summary>DateTimeOffset 使用不依赖区域性的往返格式，其他值沿用请求 Culture。</summary>
    private static string FormatScalarValue(object value, CultureInfo culture) =>
        value is DateTimeOffset dateTimeOffset
            ? dateTimeOffset.ToString("O", CultureInfo.InvariantCulture)
            : Convert.ToString(value, culture) ?? string.Empty;

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
