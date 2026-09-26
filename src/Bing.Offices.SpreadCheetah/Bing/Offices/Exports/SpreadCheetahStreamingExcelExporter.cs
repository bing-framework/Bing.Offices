using System.Collections;
using System.Globalization;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Drawing;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.IO;
using Bing.Offices.Internals;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using ExcelStyles = Bing.Offices.Styles;
using SpreadCheetah;
using SpreadCheetah.Styling;
using SpreadCheetah.Tables;
using SpreadCheetah.Worksheets;

namespace Bing.Offices.Exports;

/// <summary>
/// 基于 SpreadCheetah 的前向流式 XLSX 导出器。
/// </summary>
/// <remarks>
/// Provider 只创建新 XLSX，并按 Sheet、批次和行的顺序串行写入，不缓存完整工作簿。
/// </remarks>
public sealed class SpreadCheetahStreamingExcelExporter : IExcelStreamingExporter,
    IExcelProviderFeatureDescriptor
{
    /// <summary>
    /// 用于标识导出错误所属提供程序的名称。
    /// </summary>
    private const string Provider = "SpreadCheetah";
    /// <summary>
    /// 用于生成固定列和动态列映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 负责将导出结果安全提交到目标文件的服务。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;

    /// <summary>
    /// 初始化一个 <see cref="SpreadCheetahStreamingExcelExporter"/> 类型的实例。
    /// </summary>
    /// <param name="mappingPlanFactory">映射计划工厂；为 null 时使用默认工厂。</param>
    /// <param name="fileExportCommitter">文件提交服务；为 null 时使用默认服务。</param>
    public SpreadCheetahStreamingExcelExporter(IExcelMappingPlanFactory mappingPlanFactory = null,
        IFileExportCommitter fileExportCommitter = null)
    {
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault();
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
    }

    /// <inheritdoc />
    public string ProviderName => Provider;

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.StreamingWorkbookCreation
        | ExcelProviderFeatures.Tables | ExcelProviderFeatures.AutoFilter
        | ExcelProviderFeatures.FreezePanes | ExcelProviderFeatures.FormulaText;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats => Array.Empty<ExcelFormat>();

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> WriteFormats { get; } = new[] { ExcelFormat.Xlsx };

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookImport => false;

    /// <inheritdoc />
    public bool SupportsBatchImport => false;

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookExport => false;

    /// <inheritdoc />
    public bool SupportsTrueAsyncIo => true;

    /// <inheritdoc />
    public IReadOnlyList<string> Limitations { get; } = new[]
    {
        "仅创建新的 XLSX 工作簿，不读取或修改已有模板。",
        "底层 ZipArchive 在条目收尾时可能执行少量同步流写入。",
        "支持固定列、动态列、映射转换、基础样式、表格、自动筛选、冻结、PNG/JPEG 图片和基础原生数据校验。",
        "有图片时使用磁盘暂存修正绘图坐标及 JPEG 封装，完成后再复制目标；无图片时直接前向输出。",
        "不支持 AutoFit、条件格式、名称范围、打印布局、图表、批注、宏或加密。"
    };

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;

    /// <inheritdoc />
    public void ExportBatches(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
        => ExportBatchesAsync(request, destination, options, cancellationToken)
            .GetAwaiter().GetResult();

    /// <inheritdoc />
    public async Task ExportBatchesAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
    {
        ValidateArguments(request, destination, options, cancellationToken);
        if (request.Sheets.Any(sheet => sheet.Images.Count > 0))
            await SpreadCheetahSheetContentWriter.ExportStagedAsync(request, destination,
                stream => ExportCoreAsync(request, stream, options, cancellationToken), cancellationToken).ConfigureAwait(false);
        else
            await ExportCoreAsync(request, destination, options, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 异步按批次写入新 XLSX 工作簿。
    /// </summary>
    /// <param name="request">工作簿或工作表导出请求。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="options">流式导出选项。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private async Task ExportCoreAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options, CancellationToken cancellationToken)
    {
        ValidateArguments(request, destination, options, cancellationToken);
        ValidateRequest(request);
        ExcelSheetContent.Validate(request);
        ExcelReportPreflight.Validate(request, ProviderName, 1048576, 16384, false);

        options ??= new ExcelStreamingExportOptions();
        options.Validate();
        // 映射计划和静态请求约束必须在创建工作簿之前完成，避免后续 Sheet 的配置错误
        // 暴露已经写出的前置内容，也避免为预检枚举调用方数据源。
        var preparedSheets = request.Sheets
            .Select(sheet =>
            {
                var bindings = CreateBindings(sheet);
                return (Sheet: sheet, Bindings: bindings,
                    Options: CreateWorksheetOptions(sheet, bindings.Columns.Count),
                    Table: CreateTable(sheet, bindings.Columns.Count));
            })
            .ToArray();

        await using var nonDisposingDestination = new NonDisposingStream(destination);
        await using var spreadsheet = await Spreadsheet.CreateNewAsync(nonDisposingDestination,
            cancellationToken: cancellationToken).ConfigureAwait(false);
        var images = await SpreadCheetahSheetContentWriter.EmbedAsync(spreadsheet, request, cancellationToken).ConfigureAwait(false);
        foreach (var prepared in preparedSheets)
        {
            var sheet = prepared.Sheet;
            var bindingSet = prepared.Bindings;
            var bindings = bindingSet.Columns;
            cancellationToken.ThrowIfCancellationRequested();
            var worksheetOptions = prepared.Options;
            var headerStyle = CreateStyle(spreadsheet, sheet.HeaderStyle);
            var bodyStyle = CreateStyle(spreadsheet, sheet.BodyStyle ?? sheet.SheetStyle);
            await spreadsheet.StartWorksheetAsync(sheet.Name, worksheetOptions, cancellationToken)
                .ConfigureAwait(false);
            SpreadCheetahSheetContentWriter.Apply(spreadsheet, sheet, images, cancellationToken);
            var table = prepared.Table;
            if (table != null)
                spreadsheet.StartTable(table, ToColumnName(0));
            await spreadsheet.AddHeaderRowAsync(bindings.Select(binding => binding.Title).ToArray(),
                headerStyle, cancellationToken).ConfigureAwait(false);

            var batch = new List<Cell[]>(options?.BatchSize ?? 1000);
            var rowIndex = sheet.DataRowStartIndex + 1;
            foreach (var item in sheet.Data)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (table != null && rowIndex - 2 >= sheet.Tables[0].Range.EndRow)
                    throw new BingOfficesConfigurationException(
                        $"Sheet {sheet.Name} 数据行数超过声明的表格区域。", stage: BingOfficesStage.Write);
                if (item == null)
                    throw new ArgumentException($"Sheet {sheet.Name} 包含 null 数据项。", nameof(request));
                batch.Add(CreateRow(item, bindingSet, sheet.Culture, bodyStyle, sheet, rowIndex++));
                if (batch.Count < options.BatchSize)
                    continue;
                await WriteBatchAsync(spreadsheet, batch, cancellationToken).ConfigureAwait(false);
                batch.Clear();
            }

            if (table != null && rowIndex - 2 != sheet.Tables[0].Range.EndRow)
                throw new BingOfficesConfigurationException(
                    $"Sheet {sheet.Name} 数据行数少于声明的表格区域。", stage: BingOfficesStage.Write);
            await WriteBatchAsync(spreadsheet, batch, cancellationToken).ConfigureAwait(false);
            if (table != null)
                await spreadsheet.FinishTableAsync(cancellationToken).ConfigureAwait(false);
        }

        await spreadsheet.FinishAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public void ExportBatchesToFile(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        _fileExportCommitter.Commit(path,
            stream => ExportBatches(request, stream, options, cancellationToken),
            cancellationToken, ProviderName);
    }

    /// <inheritdoc />
    public Task ExportBatchesToFileAsync(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));
        return _fileExportCommitter.CommitAsync(path,
            (stream, token) => ExportBatchesAsync(request, stream, options, token),
            cancellationToken, ProviderName);
    }

    /// <summary>
    /// 异步写入当前批次的单元格行。
    /// </summary>
    /// <param name="spreadsheet">当前输出工作簿。</param>
    /// <param name="batch">待写入的单元格行批次。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static async Task WriteBatchAsync(Spreadsheet spreadsheet, List<Cell[]> batch,
        CancellationToken cancellationToken)
    {
        foreach (var row in batch)
            await spreadsheet.AddRowAsync(row, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 创建工作表的固定列和动态列绑定。
    /// </summary>
    /// <param name="sheet">工作表导出请求。</param>
    /// <returns>按输出顺序排列的列绑定及动态键集合。</returns>
    private BindingSet CreateBindings(ExcelSheetExportRequest sheet)
    {
        var mappingConfiguration = MergeRequestDynamicColumns(sheet.MappingConfiguration,
            sheet.DynamicColumns);
        var create = typeof(IExcelMappingPlanFactory).GetMethods()
            .Single(method => method.Name == nameof(IExcelMappingPlanFactory.Create)
                && method.IsGenericMethodDefinition && method.GetParameters().Length == 3);
        IExcelMappingPlan plan;
        try
        {
            plan = (IExcelMappingPlan)create.MakeGenericMethod(sheet.ItemType).Invoke(_mappingPlanFactory,
                new object[]
                {
                    sheet.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
                    mappingConfiguration,
                    MappingDirection.Export
                });
        }
        catch (TargetInvocationException exception) when (exception.InnerException is BingOfficesException)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            throw new BingOfficesConfigurationException(
                $"Sheet {sheet.Name} 映射计划预检失败: {exception.InnerException.Message}",
                exception.InnerException, BingOfficesStage.Plan);
        }

        var bindings = new List<ColumnBinding>();
        foreach (var column in plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn))
        {
            var property = sheet.ItemType.GetProperty(column.Name, BindingFlags.Instance | BindingFlags.Public)
                ?? throw new BingOfficesConfigurationException($"属性不存在: {column.Name}",
                    stage: BingOfficesStage.Plan);
            if (IsDynamicContainerColumn(property))
                continue;
            bindings.Add(ColumnBinding.CreateFixed(column.Title, property, column));
        }

        var dynamicPlans = plan.DynamicColumns
            .OrderBy(column => column.ColumnIndex ?? int.MaxValue)
            .ThenBy(column => column.Order)
            .ThenBy(column => column.Key, StringComparer.Ordinal)
            .ToArray();
        if (dynamicPlans.Length > 0 && sheet.DynamicGetter == null)
            throw new BingOfficesConfigurationException($"Sheet {sheet.Name} 的动态列缺少动态值读取器。",
                stage: BingOfficesStage.Plan);

        foreach (var dynamicPlan in dynamicPlans)
        {
            var requestColumn = sheet.DynamicColumns.FirstOrDefault(column =>
                string.Equals(column.Key, dynamicPlan.Key, StringComparison.OrdinalIgnoreCase));
            var definition = CreateDynamicDefinition(dynamicPlan, requestColumn, plan);
            var binding = ColumnBinding.CreateDynamic(definition.Title, dynamicPlan.Key, dynamicPlan);
            InsertDynamicBinding(bindings, binding, definition);
        }

        ValidateBindingTitles(bindings, sheet.Name);
        if (bindings.Count == 0)
            throw new BingOfficesConfigurationException($"Sheet {sheet.Name} 没有可导出的固定列。",
                stage: BingOfficesStage.Plan);
        return new BindingSet(bindings);
    }

    /// <summary>
    /// 将数据项转换为导出单元格行。
    /// </summary>
    /// <param name="item">待导出的数据项。</param>
    /// <param name="bindingSet">已解析的列绑定集合。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="bodyStyle">数据单元格样式标识。</param>
    /// <param name="sheet">工作表导出请求。</param>
    /// <param name="rowIndex">用于错误定位的行号。</param>
    /// <returns>按列绑定顺序排列的单元格数组。</returns>
    private static Cell[] CreateRow(object item, BindingSet bindingSet, CultureInfo culture,
        StyleId bodyStyle, ExcelSheetExportRequest sheet, int rowIndex)
    {
        var bindings = bindingSet.Columns;
        var actualCulture = culture ?? CultureInfo.InvariantCulture;
        var result = new Cell[bindings.Count];
        IDictionary<string, object> dynamicValues;
        try
        {
            dynamicValues = sheet.DynamicGetter?.Invoke(item);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (IsNonFatal(exception))
        {
            var dynamicColumnIndex = FindFirstDynamicColumnIndex(bindingSet.Columns);
            var dynamicBinding = dynamicColumnIndex >= 0 ? bindingSet.Columns[dynamicColumnIndex] : null;
            throw CreateExportException("SpreadCheetah 动态值读取器执行失败。", exception,
                sheet.Name, rowIndex, dynamicColumnIndex >= 0 ? dynamicColumnIndex + 1 : null,
                dynamicBinding?.DynamicKey, BingOfficesErrorCode.UserExtensionFailed);
        }
        if (sheet.FailOnUnknownDynamicValues && dynamicValues != null)
        {
            var unknown = dynamicValues.Keys.FirstOrDefault(key => !bindingSet.DynamicKeys.Contains(key));
            if (unknown != null)
                throw new BingOfficesConfigurationException($"Sheet {sheet.Name} 包含未声明动态列值: {unknown}",
                    stage: BingOfficesStage.Validate);
        }
        for (var index = 0; index < bindings.Count; index++)
        {
            var binding = bindings[index];
            object raw;
            if (binding.Property != null)
            {
                try
                {
                    raw = binding.Property.GetValue(item);
                }
                catch (BingOfficesException)
                {
                    throw;
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception exception) when (IsNonFatal(exception))
                {
                    var actual = UnwrapInvocationException(exception);
                    // 反射会包裹属性抛出的异常，取消、公共异常和致命异常须保持原始对象。
                    if (actual is OperationCanceledException || actual is BingOfficesException || !IsNonFatal(actual))
                        ExceptionDispatchInfo.Capture(actual).Throw();
                    throw CreateExportException("SpreadCheetah 属性读取器执行失败。",
                        actual, sheet.Name, rowIndex, index + 1,
                        binding.Property.Name, BingOfficesErrorCode.UserExtensionFailed);
                }
            }
            else if (dynamicValues != null && dynamicValues.TryGetValue(binding.DynamicKey, out var dynamicValue))
                raw = dynamicValue;
            else
                raw = null;

            var value = binding.Fixed != null
                ? ConvertFixedValue(raw, binding.Fixed, binding.Property, sheet.Name, rowIndex, index + 1,
                    actualCulture)
                : ConvertDynamicValue(raw, binding.Dynamic, sheet.Name, rowIndex, index + 1, actualCulture);
            try
            {
                result[index] = CreateCell(value, actualCulture, bodyStyle);
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (IsNonFatal(exception))
            {
                throw CreateExportException("SpreadCheetah 单元格值写入转换失败。", exception,
                    sheet.Name, rowIndex, index + 1,
                    binding.Property?.Name ?? binding.DynamicKey, BingOfficesErrorCode.ExportFailed);
            }
        }
        return result;
    }

    /// <summary>
    /// 按固定列映射转换导出值。
    /// </summary>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="column">对应列的映射配置。</param>
    /// <param name="property">固定列对应的实体属性。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">用于错误定位的行号。</param>
    /// <param name="columnIndex">用于错误定位的列号。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <returns>转换后的导出值；原始值或转换结果为空时为 null。</returns>
    private static object ConvertFixedValue(object raw, IExcelMappingColumn column, PropertyInfo property,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var context = new ExcelConversionContext(raw, column.Name, property.PropertyType, sheetName,
            rowIndex, columnIndex, culture, CreateCellValueForExport(raw, culture, sheetName, rowIndex, columnIndex, column.Name));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            try
            {
                if (converter.TryConvertTo(context, out var converted))
                    return converted;
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (IsNonFatal(exception))
            {
                throw CreateExportException("SpreadCheetah 值转换器执行失败。", exception,
                    sheetName, rowIndex, columnIndex, column.Name,
                    BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        try
        {
            if (raw == null)
                return null;
            if (column.ValueMap != null)
            {
                var rawText = Convert.ToString(raw, culture);
                var mapped = column.ValueMap.FirstOrDefault(pair =>
                    string.Equals(pair.Value, rawText, StringComparison.Ordinal));
                if (!string.IsNullOrEmpty(mapped.Key))
                    return mapped.Key;
            }
            if (!string.IsNullOrWhiteSpace(column.Formatter) && raw is IFormattable formattable)
                return formattable.ToString(column.Formatter, culture);
            return raw is DateTimeOffset dateTimeOffset ? dateTimeOffset.DateTime : raw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (IsNonFatal(exception))
        {
            throw CreateExportException("SpreadCheetah 固定列值转换失败。", exception,
                sheetName, rowIndex, columnIndex, column.Name, BingOfficesErrorCode.ExportFailed);
        }
    }

    /// <summary>
    /// 按动态列映射转换导出值。
    /// </summary>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="column">对应列的映射配置。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">用于错误定位的行号。</param>
    /// <param name="columnIndex">用于错误定位的列号。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <returns>转换后的导出值；原始值或转换结果为空时为 null。</returns>
    private static object ConvertDynamicValue(object raw, IExcelDynamicMappingColumn column,
        string sheetName, int rowIndex, int columnIndex, CultureInfo culture)
    {
        var targetType = ResolveDynamicType(column.DataTypeName);
        var context = new ExcelConversionContext(raw, column.Key, targetType, sheetName,
            rowIndex, columnIndex, culture, CreateCellValueForExport(raw, culture, sheetName, rowIndex, columnIndex, column.Key));
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
        {
            try
            {
                if (converter.TryConvertTo(context, out var converted))
                    return converted;
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (IsNonFatal(exception))
            {
                throw CreateExportException("SpreadCheetah 动态值转换器执行失败。", exception,
                    sheetName, rowIndex, columnIndex, column.Key,
                    BingOfficesErrorCode.UserExtensionFailed);
            }
        }
        try
        {
            if (raw == null)
                return null;
            // 请求级动态列的默认类型为 string，但历史流式 Provider 会保留调用方的
            // 数值/日期 CLR 值以便底层按原生单元格类型写出；只有文本值需要规范化。
            if (targetType == typeof(string))
                return raw;
            if (targetType.IsInstanceOfType(raw))
                return raw is DateTimeOffset dateTimeOffset ? dateTimeOffset.DateTime : raw;
            if (targetType.IsEnum)
                return Enum.Parse(targetType, Convert.ToString(raw, culture), true);
            if (targetType == typeof(Guid))
                return raw is Guid ? raw : Guid.Parse(Convert.ToString(raw, culture));
            if (targetType == typeof(DateTime))
                return raw is DateTimeOffset offset ? offset.DateTime : Convert.ToDateTime(raw, culture);
            if (targetType == typeof(DateTimeOffset))
                return raw is DateTime date ? new DateTimeOffset(DateTime.SpecifyKind(date, DateTimeKind.Unspecified),
                    TimeSpan.Zero) : Convert.ToDateTime(raw, culture);
            if (targetType == typeof(byte[]))
                return raw is byte[] bytes ? bytes : Convert.FromBase64String(Convert.ToString(raw, culture));
            return Convert.ChangeType(raw, Nullable.GetUnderlyingType(targetType) ?? targetType, culture);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception) when (IsNonFatal(exception))
        {
            throw CreateExportException("SpreadCheetah 动态列值转换失败。", exception,
                sheetName, rowIndex, columnIndex, column.Key, BingOfficesErrorCode.ExportFailed);
        }
    }

    /// <summary>
    /// 创建导出值上下文并封装转换错误。
    /// </summary>
    /// <param name="value">待处理的对象值。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">用于错误定位的行号。</param>
    /// <param name="columnIndex">用于错误定位的列号。</param>
    /// <param name="propertyName">用于错误定位的属性名或动态列键。</param>
    /// <returns>包含原始值、显示文本和类型的单元格值。</returns>
    private static ExcelCellValue CreateCellValueForExport(object value, CultureInfo culture,
        string sheetName, int rowIndex, int columnIndex, string propertyName)
    {
        try
        {
            return CreateCellValue(value, culture);
        }
        catch (BingOfficesException) { throw; }
        catch (OperationCanceledException) { throw; }
        catch (Exception exception) when (IsNonFatal(exception))
        {
            throw CreateExportException("SpreadCheetah 值上下文格式化失败。", exception,
                sheetName, rowIndex, columnIndex, propertyName, BingOfficesErrorCode.ExportFailed);
        }
    }

    /// <summary>
    /// 创建包含原始值和格式化文本的单元格值。
    /// </summary>
    /// <param name="value">待处理的对象值。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <returns>包含原始值、显示文本和类型的单元格值。</returns>
    private static ExcelCellValue CreateCellValue(object value, CultureInfo culture)
    {
        var kind = value switch
        {
            null => ExcelCellKind.Empty,
            bool => ExcelCellKind.Boolean,
            DateTime or DateTimeOffset => ExcelCellKind.DateTime,
            byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
                => ExcelCellKind.Number,
            _ => ExcelCellKind.Text
        };
        var text = value is IFormattable formattable
            ? formattable.ToString(null, culture)
            : Convert.ToString(value, culture) ?? string.Empty;
        return new ExcelCellValue(value, text, kind);
    }

    /// <summary>
    /// 解析动态列的数据类型。
    /// </summary>
    /// <param name="name">动态列数据类型名称。</param>
    /// <returns>数据类型对应的 CLR 类型。</returns>
    private static Type ResolveDynamicType(string name)
    {
        switch ((name ?? "string").Trim().ToLowerInvariant())
        {
            case "object": return typeof(object);
            case "string": return typeof(string);
            case "bool":
            case "boolean": return typeof(bool);
            case "byte": return typeof(byte);
            case "int16": return typeof(short);
            case "int":
            case "int32": return typeof(int);
            case "long":
            case "int64": return typeof(long);
            case "float":
            case "single": return typeof(float);
            case "double": return typeof(double);
            case "decimal": return typeof(decimal);
            case "datetime": return typeof(DateTime);
            case "datetimeoffset": return typeof(DateTimeOffset);
            case "guid": return typeof(Guid);
            case "bytes": return typeof(byte[]);
            default: throw new BingOfficesConfigurationException($"动态列数据类型不受支持: {name}",
                stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 将请求级动态列合并到映射配置。
    /// </summary>
    /// <param name="configuration">已有映射配置。</param>
    /// <param name="definitions">请求级动态列定义集合。</param>
    /// <returns>合并后的配置；没有动态列时返回原配置，原配置可为 null。</returns>
    private static ExcelMappingConfiguration MergeRequestDynamicColumns(
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
    /// 获取动态列的相对位置标识。
    /// </summary>
    /// <param name="placement">动态列位置配置。</param>
    /// <returns>带 before 或 after 前缀的位置标识；未指定相对位置时为 null。</returns>
    private static string GetPlacementKey(ExcelColumnPlacement placement)
    {
        if (!string.IsNullOrWhiteSpace(placement?.BeforeKey))
            return $"before:{placement.BeforeKey}";
        if (!string.IsNullOrWhiteSpace(placement?.AfterKey))
            return $"after:{placement.AfterKey}";
        return null;
    }

    /// <summary>
    /// 获取动态列类型对应的配置名称。
    /// </summary>
    /// <param name="type">动态列的 CLR 类型。</param>
    /// <returns>映射配置支持的数据类型名称。</returns>
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
        throw new BingOfficesConfigurationException($"动态列数据类型不在允许列表中: {type.FullName}",
            stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 从映射计划和请求创建动态列定义。
    /// </summary>
    /// <param name="column">对应列的映射配置。</param>
    /// <param name="requestColumn">请求级动态列定义；没有匹配项时为 null。</param>
    /// <param name="mapping">当前工作表的映射计划。</param>
    /// <returns>应用请求级覆盖和布局默认值的动态列定义。</returns>
    private static ExcelDynamicColumnDefinition CreateDynamicDefinition(IExcelDynamicMappingColumn column,
        ExcelDynamicColumnDefinition requestColumn, IExcelMappingPlan mapping)
    {
        var columnIndex = column.ColumnIndex;
        var placementKey = column.PlacementKey;
        if (!columnIndex.HasValue && string.IsNullOrWhiteSpace(placementKey)
            && mapping.DynamicColumns.Count == 1)
        {
            columnIndex = mapping.Layout?.ColumnIndex;
            placementKey = mapping.Layout?.PlacementKey;
        }
        return new ExcelDynamicColumnDefinition
        {
            Key = column.Key,
            Title = requestColumn?.Title ?? column.Title,
            Aliases = requestColumn?.Aliases ?? column.Aliases,
            DataType = ResolveDynamicType(column.DataTypeName),
            Order = column.Order,
            PhysicalColumnIndex = columnIndex,
            Placement = CreatePlacement(placementKey),
            NumberFormat = column.NumberFormat ?? mapping.Style?.NumberFormat,
            HeaderStyle = requestColumn?.HeaderStyle,
            BodyStyle = requestColumn?.BodyStyle,
            ConverterName = column.ConverterName,
            ValidatorName = column.ValidatorName,
            ValidationRuleNames = column.ValidationRuleNames,
            ImageMultiplicity = column.ImageMultiplicity
        };
    }

    /// <summary>
    /// 解析动态列的相对位置。
    /// </summary>
    /// <param name="placementKey">包含 before 或 after 前缀的相对位置标识。</param>
    /// <returns>相对位置配置；标识为空时为 null。</returns>
    private static ExcelColumnPlacement CreatePlacement(string placementKey)
    {
        if (string.IsNullOrWhiteSpace(placementKey))
            return null;
        var separator = placementKey.IndexOfAny(new[] { ':', '-' });
        var key = separator < 0 ? placementKey : placementKey.Substring(separator + 1);
        return placementKey.StartsWith("before:", StringComparison.OrdinalIgnoreCase)
            || placementKey.StartsWith("before-", StringComparison.OrdinalIgnoreCase)
            ? ExcelColumnPlacement.Before(key)
            : ExcelColumnPlacement.After(key);
    }

    /// <summary>
    /// 按物理索引或相对位置插入动态列绑定。
    /// </summary>
    /// <param name="bindings">当前有序列绑定集合。</param>
    /// <param name="binding">待插入的动态列绑定。</param>
    /// <param name="definition">动态列及其位置定义。</param>
    private static void InsertDynamicBinding(List<ColumnBinding> bindings, ColumnBinding binding,
        ExcelDynamicColumnDefinition definition)
    {
        var physicalIndex = definition.PhysicalColumnIndex ?? definition.Placement?.PhysicalColumnIndex;
        if (physicalIndex.HasValue)
        {
            if (physicalIndex.Value < 0 || physicalIndex.Value > bindings.Count)
                throw new ArgumentOutOfRangeException(nameof(definition.PhysicalColumnIndex),
                    $"动态列 {definition.Key} 的物理索引超出当前列计划。");
            bindings.Insert(physicalIndex.Value, binding);
            return;
        }
        if (!string.IsNullOrWhiteSpace(definition.Placement?.BeforeKey))
        {
            var index = bindings.FindIndex(item => string.Equals(item.DynamicKey,
                definition.Placement.BeforeKey, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new BingOfficesConfigurationException(
                    $"动态列 {definition.Key} 的 Before 目标不存在: {definition.Placement.BeforeKey}",
                    stage: BingOfficesStage.Plan);
            bindings.Insert(index, binding);
            return;
        }
        if (!string.IsNullOrWhiteSpace(definition.Placement?.AfterKey))
        {
            var index = bindings.FindIndex(item => string.Equals(item.DynamicKey,
                definition.Placement.AfterKey, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
                throw new BingOfficesConfigurationException(
                    $"动态列 {definition.Key} 的 After 目标不存在: {definition.Placement.AfterKey}",
                    stage: BingOfficesStage.Plan);
            bindings.Insert(index + 1, binding);
            return;
        }
        bindings.Add(binding);
    }

    /// <summary>
    /// 验证导出列标题非空且不重复。
    /// </summary>
    /// <param name="bindings">当前有序列绑定集合。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    private static void ValidateBindingTitles(IReadOnlyList<ColumnBinding> bindings, string sheetName)
    {
        var titles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var binding in bindings)
            if (string.IsNullOrWhiteSpace(binding.Title) || !titles.Add(binding.Title))
                throw new BingOfficesConfigurationException(
                    $"Sheet {sheetName} 导出列标题为空或重复: {binding.Title}", stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 判断属性是否为动态值字典容器。
    /// </summary>
    /// <param name="property">固定列对应的实体属性。</param>
    /// <returns>属性为动态值字典容器时为 true，否则为 false。</returns>
    private static bool IsDynamicContainerColumn(PropertyInfo property) =>
        typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType);

    /// <summary>
    /// 将值转换为带样式的原生单元格。
    /// </summary>
    /// <param name="value">待处理的对象值。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="style">单元格样式标识。</param>
    /// <returns>包含转换后值和样式的原生单元格。</returns>
    private static Cell CreateCell(object value, CultureInfo culture, StyleId style)
    {
        if (value == null)
            return new Cell((string)null, style);
        if (value is string text)
        {
            if (text.StartsWith("=", StringComparison.Ordinal) && text.Length > 1)
                return new Cell(new global::SpreadCheetah.Formula(text.Substring(1)), style);
            return new Cell(text, style);
        }
        if (value is DateTime dateTime) return new Cell(dateTime, style);
        if (value is DateTimeOffset dateTimeOffset) return new Cell(dateTimeOffset.DateTime, style);
        if (value is bool boolean) return new Cell(boolean, style);
        if (value is byte byteValue) return new Cell((int)byteValue, style);
        if (value is short shortValue) return new Cell((int)shortValue, style);
        if (value is int intValue) return new Cell(intValue, style);
        if (value is uint uintValue) return new Cell((long)uintValue, style);
        if (value is long longValue) return new Cell(longValue, style);
        if (value is ulong ulongValue && ulongValue <= long.MaxValue) return new Cell((long)ulongValue, style);
        if (value is float floatValue) return new Cell(floatValue, style);
        if (value is double doubleValue) return new Cell(doubleValue, style);
        if (value is decimal decimalValue) return new Cell(decimalValue, style);
        return new Cell(Convert.ToString(value, culture) ?? string.Empty, style);
    }

    /// <summary>
    /// 验证流式导出参数和取消状态。
    /// </summary>
    /// <param name="request">工作簿或工作表导出请求。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="options">流式导出选项。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ValidateArguments(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options, CancellationToken cancellationToken)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
        options?.Validate();
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 验证请求是否符合前向 XLSX 导出能力。
    /// </summary>
    /// <param name="request">工作簿或工作表导出请求。</param>
    private static void ValidateRequest(ExcelWorkbookExportRequest request)
    {
        if (request.Format != ExcelFormat.Xlsx)
            throw Unsupported("仅支持 XLSX 格式。");
        if (request.Template != null)
            throw Unsupported("不支持模板工作簿。");
        if (request.MetadataSpecified)
            throw Unsupported("暂不支持工作簿元数据。");
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in request.Sheets)
        {
            if (sheet == null)
                throw new ArgumentException("Workbook 包含 null Sheet。", nameof(request));
            if (string.IsNullOrWhiteSpace(sheet.Name))
                throw new BingOfficesConfigurationException("Sheet 名称不能为空。", stage: BingOfficesStage.Plan);
            if (!sheetNames.Add(sheet.Name))
                throw new BingOfficesConfigurationException($"Workbook 包含重复 Sheet 名称: {sheet.Name}",
                    stage: BingOfficesStage.Plan);
            if (sheet.Data == null)
                throw new ArgumentNullException(nameof(sheet.Data), $"Sheet {sheet.Name} 数据源不能为空。");
            if (sheet.HeaderRowIndex != 0 || sheet.DataRowStartIndex != 1)
                throw Unsupported("仅支持默认表头和数据行位置。");
            if (sheet.Charts?.Count > 0 || sheet.HeaderRows?.Count > 0
                || !string.IsNullOrWhiteSpace(sheet.TemplateRegion) || sheet.RowHeight != null
                || sheet.NamedRanges?.Count > 0 || sheet.ConditionalFormats?.Count > 0
                || sheet.PrintLayout != null)
                throw Unsupported("请求包含 SpreadCheetah 前向模型无法表达的布局或完整 DOM 功能。");
            if (sheet.ColumnWidth != null
                && sheet.ColumnWidth.Mode != ExcelColumnWidthMode.None
                && sheet.ColumnWidth.Mode != ExcelColumnWidthMode.Fixed)
                throw Unsupported("SpreadCheetah 不支持 AutoFit 或 Adaptive 列宽。");
            foreach (var filter in sheet.AutoFilters ?? Array.Empty<ExcelAutoFilterDefinition>())
                if (filter.Range.StartRow != 0 || filter.Range.StartColumn != 0)
                    throw Unsupported("SpreadCheetah 自动筛选必须从 A1 开始。");
            if (sheet.Tables?.Count > 0 && sheet.AutoFilters.Any(filter =>
                    !sheet.Tables.Any(table => ExcelReportPreflight.SameRange(table.Range, filter.Range))))
                throw Unsupported("SpreadCheetah 底层格式不允许在同一 Sheet 同时创建表格和独立自动筛选。",
                    BingOfficesOperation.Export);
            if (sheet.FreezePane?.TopRow != null || sheet.FreezePane?.LeftColumn != null)
                throw Unsupported("SpreadCheetah 只支持冻结顶部行和左侧列。");
        }
    }

    /// <summary>
    /// 创建工作表可见性、冻结和列宽选项。
    /// </summary>
    /// <param name="sheet">工作表导出请求。</param>
    /// <param name="columnCount">导出列数量。</param>
    /// <returns>适用于当前工作表的原生选项。</returns>
    private static WorksheetOptions CreateWorksheetOptions(ExcelSheetExportRequest sheet, int columnCount)
    {
        var options = new WorksheetOptions
        {
            Visibility = sheet.Hidden ? WorksheetVisibility.Hidden : WorksheetVisibility.Visible
        };
        if (sheet.FreezePane != null)
        {
            options.FrozenRows = sheet.FreezePane.Rows;
            options.FrozenColumns = sheet.FreezePane.Columns;
        }
        if (sheet.AutoFilters?.Count > 0 && sheet.Tables.Count == 0)
            options.AutoFilter = new AutoFilterOptions(ToA1(sheet.AutoFilters[0].Range));
        if (sheet.ColumnWidth?.Mode == ExcelColumnWidthMode.Fixed)
        {
            for (var index = 0; index < columnCount; index++)
                options.Column(index + 1).Width = sheet.ColumnWidth.FixedWidth;
        }
        return options;
    }

    /// <summary>
    /// 创建工作表的前向表格定义。
    /// </summary>
    /// <param name="sheet">工作表导出请求。</param>
    /// <param name="columnCount">导出列数量。</param>
    /// <returns>原生表格定义；未请求表格时为 null。</returns>
    private static Table CreateTable(ExcelSheetExportRequest sheet, int columnCount)
    {
        if (sheet.Tables == null || sheet.Tables.Count == 0)
            return null;
        if (sheet.Tables.Count > 1)
            throw Unsupported("SpreadCheetah 每个 Sheet 仅支持一个前向表格。", BingOfficesOperation.Export);
        var definition = sheet.Tables[0];
        if (!definition.HasHeaders || definition.ShowTotals
            || definition.Range.StartRow != 0 || definition.Range.StartColumn != 0
            || definition.Range.EndColumn != columnCount - 1)
            throw Unsupported("SpreadCheetah 表格仅支持从 A1 开始、包含表头且不含汇总行的前向区域。",
                BingOfficesOperation.Export);
        var style = ParseTableStyle(definition.StyleName);
        return new Table(style, definition.Name)
        {
            NumberOfColumns = columnCount,
            BandedRows = true
        };
    }

    /// <summary>
    /// 解析原生表格样式。
    /// </summary>
    /// <param name="styleName">表格样式名称；为空时采用 Medium2。</param>
    /// <returns>解析得到的原生表格样式。</returns>
    private static TableStyle ParseTableStyle(string styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
            return TableStyle.Medium2;
        if (styleName.StartsWith("TableStyle", StringComparison.OrdinalIgnoreCase))
            styleName = styleName.Substring("TableStyle".Length);
        var field = typeof(TableStyle).GetField(styleName,
            BindingFlags.Public | BindingFlags.Static | BindingFlags.IgnoreCase);
        if (field == null)
            throw Unsupported($"SpreadCheetah 不支持表格样式: {styleName}。", BingOfficesOperation.Export);
        return (TableStyle)field.GetValue(null);
    }

    /// <summary>
    /// 注册公共单元格样式并返回样式标识。
    /// </summary>
    /// <param name="spreadsheet">当前输出工作簿。</param>
    /// <param name="source">待转换的公共单元格样式。</param>
    /// <returns>注册后的样式标识；未提供样式时为默认标识。</returns>
    private static StyleId CreateStyle(Spreadsheet spreadsheet, ExcelStyles.ExcelCellStyle source)
    {
        if (source == null)
            return default;
        var style = new Style();
        if (!string.IsNullOrWhiteSpace(source.FontName)) style.Font.Name = source.FontName;
        if (source.FontSize.HasValue) style.Font.Size = source.FontSize.Value;
        if (source.Bold.HasValue) style.Font.Bold = source.Bold.Value;
        if (source.Italic.HasValue) style.Font.Italic = source.Italic.Value;
        if (source.Underline.HasValue)
            style.Font.Underline = source.Underline.Value ? Underline.Single : Underline.None;
        style.Font.Color = ToColor(source.FontColor);
        style.Fill.Color = ToColor(source.BackgroundColor ?? source.ForegroundColor);
        style.Alignment.WrapText = source.WrapText ?? false;
        if (source.Indent.HasValue) style.Alignment.Indent = source.Indent.Value;
        style.Alignment.Horizontal = ToHorizontalAlignment(source.HorizontalAlignment);
        style.Alignment.Vertical = ToVerticalAlignment(source.VerticalAlignment);
        if (!string.IsNullOrWhiteSpace(source.NumberFormat))
            style.Format = NumberFormat.Custom(source.NumberFormat);
        style.Border.Top = ToBorder(source.TopBorder);
        style.Border.Bottom = ToBorder(source.BottomBorder);
        style.Border.Left = ToBorder(source.LeftBorder);
        style.Border.Right = ToBorder(source.RightBorder);
        return spreadsheet.AddStyle(style);
    }

    /// <summary>
    /// 转换公共边框样式。
    /// </summary>
    /// <param name="border">待转换的公共边框样式。</param>
    /// <returns>原生边框样式。</returns>
    private static EdgeBorder ToBorder(ExcelStyles.ExcelBorderStyle border)
    {
        var lineStyle = border?.LineStyle switch
        {
            ExcelStyles.ExcelBorderLineStyle.Thin => global::SpreadCheetah.Styling.BorderStyle.Thin,
            ExcelStyles.ExcelBorderLineStyle.Medium => global::SpreadCheetah.Styling.BorderStyle.Medium,
            ExcelStyles.ExcelBorderLineStyle.Thick => global::SpreadCheetah.Styling.BorderStyle.Thick,
            ExcelStyles.ExcelBorderLineStyle.Dashed => global::SpreadCheetah.Styling.BorderStyle.Dashed,
            ExcelStyles.ExcelBorderLineStyle.Dotted => global::SpreadCheetah.Styling.BorderStyle.Dotted,
            ExcelStyles.ExcelBorderLineStyle.Double => global::SpreadCheetah.Styling.BorderStyle.DoubleLine,
            _ => global::SpreadCheetah.Styling.BorderStyle.None
        };
        return new EdgeBorder { BorderStyle = lineStyle, Color = ToColor(border?.Color) };
    }

    /// <summary>
    /// 转换水平对齐方式。
    /// </summary>
    /// <param name="alignment">公共对齐方式。</param>
    /// <returns>原生水平对齐方式。</returns>
    private static global::SpreadCheetah.Styling.HorizontalAlignment ToHorizontalAlignment(
        ExcelStyles.ExcelHorizontalAlignment alignment) => alignment switch
        {
            ExcelStyles.ExcelHorizontalAlignment.Left => global::SpreadCheetah.Styling.HorizontalAlignment.Left,
            ExcelStyles.ExcelHorizontalAlignment.Center => global::SpreadCheetah.Styling.HorizontalAlignment.Center,
            ExcelStyles.ExcelHorizontalAlignment.Right => global::SpreadCheetah.Styling.HorizontalAlignment.Right,
            _ => global::SpreadCheetah.Styling.HorizontalAlignment.None
        };

    /// <summary>
    /// 转换垂直对齐方式。
    /// </summary>
    /// <param name="alignment">公共对齐方式。</param>
    /// <returns>原生垂直对齐方式。</returns>
    private static global::SpreadCheetah.Styling.VerticalAlignment ToVerticalAlignment(
        ExcelStyles.ExcelVerticalAlignment alignment) => alignment switch
        {
            ExcelStyles.ExcelVerticalAlignment.Top => global::SpreadCheetah.Styling.VerticalAlignment.Top,
            ExcelStyles.ExcelVerticalAlignment.Center => global::SpreadCheetah.Styling.VerticalAlignment.Center,
            _ => global::SpreadCheetah.Styling.VerticalAlignment.Bottom
        };

    /// <summary>
    /// 解析十六进制颜色。
    /// </summary>
    /// <param name="color">待解析的公共颜色。</param>
    /// <returns>原生颜色；未指定颜色时为 null。</returns>
    private static Color? ToColor(ExcelStyles.ExcelColor color)
    {
        if (string.IsNullOrWhiteSpace(color?.Argb))
            return null;
        var text = color.Argb.Trim().TrimStart('#');
        if (text.Length == 6) text = "FF" + text;
        if (text.Length != 8 || !int.TryParse(text, NumberStyles.HexNumber,
                CultureInfo.InvariantCulture, out var argb))
            throw new ArgumentException($"颜色必须是有效的 6 或 8 位十六进制值: {color.Argb}");
        return Color.FromArgb(argb);
    }

    /// <summary>
    /// 将单元格区域转换为 A1 引用。
    /// </summary>
    /// <param name="range">从 0 开始的矩形单元格区域。</param>
    /// <returns>包含起止单元格的 A1 区域引用。</returns>
    private static string ToA1(ExcelRangeDefinition range)
    {
        if (range == null) throw new ArgumentNullException(nameof(range));
        return $"{ToColumnName(range.StartColumn)}{range.StartRow + 1}:{ToColumnName(range.EndColumn)}{range.EndRow + 1}";
    }

    /// <summary>
    /// 将列索引转换为 Excel 列名。
    /// </summary>
    /// <param name="zeroBasedColumn">从 0 开始的列索引。</param>
    /// <returns>由字母组成的 Excel 列名。</returns>
    private static string ToColumnName(int zeroBasedColumn)
    {
        var value = zeroBasedColumn + 1;
        var result = string.Empty;
        while (value > 0)
        {
            value--;
            result = (char)('A' + value % 26) + result;
            value /= 26;
        }
        return result;
    }

    /// <summary>
    /// 创建提供程序不支持指定能力的异常。
    /// </summary>
    /// <param name="message">错误说明。</param>
    /// <returns>包含提供程序和预检阶段信息的异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message)
        => new(message, provider: "SpreadCheetah", operation: BingOfficesOperation.Export,
            stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 创建提供程序不支持指定能力的异常。
    /// </summary>
    /// <param name="message">错误说明。</param>
    /// <param name="operation">发生错误的操作类型。</param>
    /// <returns>包含提供程序和预检阶段信息的异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message,
        BingOfficesOperation operation)
        => new(message, provider: "SpreadCheetah", operation: operation,
            stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 判断异常是否允许封装为导出错误。
    /// </summary>
    /// <param name="exception">待判断或解包的异常。</param>
    /// <returns>非内存耗尽且非栈溢出异常时为 true，否则为 false。</returns>
    private static bool IsNonFatal(Exception exception) =>
        exception is not OutOfMemoryException && exception is not StackOverflowException;

    /// <summary>
    /// 提取反射调用包装的原始异常。
    /// </summary>
    /// <param name="exception">待判断或解包的异常。</param>
    /// <returns>存在内部异常时返回该异常，否则返回原异常。</returns>
    private static Exception UnwrapInvocationException(Exception exception) =>
        exception is TargetInvocationException invocation && invocation.InnerException != null
            ? invocation.InnerException
            : exception;

    /// <summary>
    /// 查找首个动态列的位置。
    /// </summary>
    /// <param name="columns">有序导出列绑定集合。</param>
    /// <returns>从 0 开始的首个动态列索引；不存在时为 -1。</returns>
    private static int FindFirstDynamicColumnIndex(IReadOnlyList<ColumnBinding> columns)
    {
        for (var index = 0; index < columns.Count; index++)
            if (columns[index].Dynamic != null)
                return index;
        return -1;
    }

    /// <summary>
    /// 创建带单元格定位信息的导出异常。
    /// </summary>
    /// <param name="message">错误说明。</param>
    /// <param name="innerException">原始异常；没有原始异常时为 null。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowIndex">用于错误定位的行号。</param>
    /// <param name="columnIndex">用于错误定位的列号。</param>
    /// <param name="propertyName">用于错误定位的属性名或动态列键。</param>
    /// <param name="code">导出错误分类。</param>
    /// <returns>包含工作表、行列和属性信息的导出异常。</returns>
    private static BingOfficesExportException CreateExportException(string message, Exception innerException,
        string sheetName, int rowIndex, int? columnIndex, string propertyName,
        BingOfficesErrorCode code) => new(message, innerException, Provider, BingOfficesStage.Validate,
            sheetName, rowIndex, columnIndex, propertyName, code);

    /// <summary>
    /// 导出列绑定及动态键的集合。
    /// </summary>
    private sealed class BindingSet
    {
        /// <summary>
        /// 初始化一个 <see cref="BindingSet"/> 类型的实例。
        /// </summary>
        /// <param name="columns">有序导出列绑定集合。</param>
        public BindingSet(IReadOnlyList<ColumnBinding> columns)
        {
            Columns = columns;
            DynamicKeys = new HashSet<string>(columns.Where(column => column.Dynamic != null)
                .Select(column => column.DynamicKey), StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 获取有序导出列绑定集合。
        /// </summary>
        public IReadOnlyList<ColumnBinding> Columns { get; }
        /// <summary>
        /// 获取忽略大小写的已声明动态列键集合。
        /// </summary>
        public HashSet<string> DynamicKeys { get; }
    }

    /// <summary>
    /// 固定列或动态列的导出绑定。
    /// </summary>
    private sealed class ColumnBinding
    {
        /// <summary>
        /// 初始化一个 <see cref="ColumnBinding"/> 类型的实例。
        /// </summary>
        /// <param name="title">导出列标题。</param>
        /// <param name="property">固定列对应的实体属性。</param>
        /// <param name="dynamicKey">动态列键；固定列为 null。</param>
        /// <param name="fixedColumn">固定列映射；动态列为 null。</param>
        /// <param name="dynamicColumn">动态列映射；固定列为 null。</param>
        private ColumnBinding(string title, PropertyInfo property, string dynamicKey,
            IExcelMappingColumn fixedColumn, IExcelDynamicMappingColumn dynamicColumn)
        {
            Title = title;
            Property = property;
            DynamicKey = dynamicKey;
            Fixed = fixedColumn;
            Dynamic = dynamicColumn;
        }

        /// <summary>
        /// 创建固定属性列的导出绑定。
        /// </summary>
        /// <param name="title">导出列标题。</param>
        /// <param name="property">固定列对应的实体属性。</param>
        /// <param name="column">对应列的映射配置。</param>
        /// <returns>固定属性列绑定。</returns>
        public static ColumnBinding CreateFixed(string title, PropertyInfo property, IExcelMappingColumn column) =>
            new(title, property, null, column, null);

        /// <summary>
        /// 创建动态键列的导出绑定。
        /// </summary>
        /// <param name="title">导出列标题。</param>
        /// <param name="key">动态列键。</param>
        /// <param name="column">对应列的映射配置。</param>
        /// <returns>动态键列绑定。</returns>
        public static ColumnBinding CreateDynamic(string title, string key, IExcelDynamicMappingColumn column) =>
            new(title, null, key, null, column);

        /// <summary>
        /// 获取导出列标题。
        /// </summary>
        public string Title { get; }
        /// <summary>
        /// 获取固定列属性；动态列为 null。
        /// </summary>
        public PropertyInfo Property { get; }
        /// <summary>
        /// 获取动态列键；固定列为 null。
        /// </summary>
        public string DynamicKey { get; }
        /// <summary>
        /// 获取固定列映射；动态列为 null。
        /// </summary>
        public IExcelMappingColumn Fixed { get; }
        /// <summary>
        /// 获取动态列映射；固定列为 null。
        /// </summary>
        public IExcelDynamicMappingColumn Dynamic { get; }
    }

    /// <summary>
    /// 不释放底层流的转发包装器。
    /// </summary>
    private sealed class NonDisposingStream : Stream
    {
        /// <summary>
        /// 由调用方持有并负责释放的底层流。
        /// </summary>
        private readonly Stream _inner;

        /// <summary>
        /// 初始化一个 <see cref="NonDisposingStream"/> 类型的实例。
        /// </summary>
        /// <param name="inner">由调用方负责释放的底层流。</param>
        public NonDisposingStream(Stream inner) => _inner = inner;
        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;
        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;
        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
        /// <inheritdoc />
        public override void Flush() => _inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => _inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default) => _inner.ReadAsync(buffer, cancellationToken);
        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => _inner.ReadAsync(buffer, offset, count, cancellationToken);
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);
        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default) => _inner.WriteAsync(buffer, cancellationToken);
        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => _inner.WriteAsync(buffer, offset, count, cancellationToken);
        /// <inheritdoc />
        protected override void Dispose(bool disposing) { }
    }
}
