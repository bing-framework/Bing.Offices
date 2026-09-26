using System.Collections;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using DataReaderConfiguration = global::ExcelDataReader.ExcelReaderConfiguration;
using DataReaderFactory = global::ExcelDataReader.ExcelReaderFactory;
using DataReader = global::ExcelDataReader.IExcelDataReader;

namespace Bing.Offices.ExcelDataReader;

/// <summary>
/// 基于 ExcelDataReader 的只读 Excel 导入器。
/// </summary>
/// <remarks>
/// ExcelDataReader 只负责前向读取，本类负责输入暂存、映射、转换、校验和结果边界。
/// </remarks>
public sealed class ExcelDataReaderExcelImporter : IExcelImporter, IExcelBatchImporter,
    IExcelProviderFeatureDescriptor
{
    /// <summary>
    /// 用于标识导入错误所属提供程序的名称。
    /// </summary>
    private const string Provider = "ExcelDataReader";
    /// <summary>
    /// 创建默认映射计划时使用的校验规则快照。
    /// </summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    /// <summary>
    /// 创建默认映射计划时使用的值转换器快照。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>
    /// 创建默认映射计划时使用的命名校验规则快照。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>
    /// 用于创建导入列映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 为每次输入暂存生成文件路径的委托。
    /// </summary>
    private readonly Func<string> _stagedPathFactory;

    /// <summary>
    /// 注册读取旧版工作簿所需的代码页编码提供程序。
    /// </summary>
    static ExcelDataReaderExcelImporter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter"/> 类型的实例。
    /// </summary>
    /// <param name="validationRules">默认映射计划使用的校验规则集合。</param>
    /// <param name="valueConverters">默认映射计划使用的值转换器集合。</param>
    /// <param name="namedValidationRules">默认映射计划使用的命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂；为 null 时使用默认工厂。</param>
    public ExcelDataReaderExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory, null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter"/> 类型的实例。
    /// </summary>
    /// <param name="validationRules">默认映射计划使用的校验规则集合。</param>
    /// <param name="valueConverters">默认映射计划使用的值转换器集合。</param>
    /// <param name="namedValidationRules">默认映射计划使用的命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂；为 null 时使用默认工厂。</param>
    /// <param name="stagedPathFactory">暂存文件路径工厂；为 null 时使用默认临时路径。</param>
    internal ExcelDataReaderExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory, Func<string> stagedPathFactory)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault().ToArray();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault(
            _valueConverters, _validationRules, _namedValidationRules);
        _stagedPathFactory = stagedPathFactory ?? CreateDefaultStagedPath;
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter"/> 类型的实例。
    /// </summary>
    /// <param name="stagedPathFactory">暂存文件路径工厂；为 null 时使用默认临时路径。</param>
    internal ExcelDataReaderExcelImporter(Func<string> stagedPathFactory)
        : this(null, null, null, null, stagedPathFactory)
    {
    }

    /// <inheritdoc />
    public string ProviderName => Provider;

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async
        | ExcelProviderCapabilities.Xls | ExcelProviderCapabilities.Xlsx
        | ExcelProviderCapabilities.Xlsb;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.FormulaCachedValues;

    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = new[]
        { ExcelFormat.Xls, ExcelFormat.Xlsx, ExcelFormat.Xlsb };

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> WriteFormats { get; } = Array.Empty<ExcelFormat>();

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookImport => true;

    /// <inheritdoc />
    public bool SupportsBatchImport => true;

    /// <inheritdoc />
    public bool SupportsCompleteWorkbookExport => false;

    /// <inheritdoc />
    public bool SupportsTrueAsyncIo => true;

    /// <inheritdoc />
    public IReadOnlyList<string> Limitations { get; } = new[]
    {
        "只读，不提供 Workbook 导出。",
        "首版只支持固定列；Entity、动态列、关系、图片和原生 Workbook 校验需使用其他 Provider。",
        "公式按读取引擎提供的缓存值处理，不执行公式计算。",
        "XLSX 的 MaxCells/MaxColumnsPerSheet 使用 ZIP 物理 Cell 预检；XLS/XLSB 请求这两项限制时明确返回 UnsupportedFeature。",
        "分批接口只支持单 Sheet，已交付批次不回滚。"
    };

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedRequest(request);
        var path = StageInput(source, request.ResourceLimits, cancellationToken);
        try
        {
            using var input = OpenStagedInput(path);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException exception)
            {
                var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
                errors.Add(ToResourceError(exception));
                return ResourceLimitedResult(new TWorkbook(), errors);
            }
            return ImportReader(input, request, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("ExcelDataReader 导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedRequest(request);
        var path = await StageInputAsync(source, request.ResourceLimits, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            await using var input = OpenStagedInput(path);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException exception)
            {
                var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
                errors.Add(ToResourceError(exception));
                return ResourceLimitedResult(new TWorkbook(), errors);
            }
            return ImportReader(input, request, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("ExcelDataReader 导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <inheritdoc />
    public ExcelBatchImportSummary ImportBatches<TItem>(Stream source,
        ExcelBatchImportRequest<TItem> request, Action<ExcelImportBatch<TItem>> onBatch,
        CancellationToken cancellationToken = default)
        where TItem : class, new()
    {
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedBatchRequest(request);
        ValidateBatchPlan(CreatePlan(request));
        var path = StageInput(source, request.ResourceLimits, cancellationToken);
        try
        {
            using var input = OpenStagedInput(path);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var resourceError = ToResourceError(preflightException);
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(1, 1,
                    Array.Empty<TItem>(), new[] { resourceError }));
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(1, budgetError.RowIndex,
                    Array.Empty<TItem>(), new[] { budgetError }));
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            using var reader = CreateReader(input);
            return ReadBatches(reader, request, onBatch, cancellationToken);
        }
        catch (BatchCallbackException exception)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("ExcelDataReader 分批导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelBatchImportSummary> ImportBatchesAsync<TItem>(Stream source,
        ExcelBatchImportRequest<TItem> request,
        Func<ExcelImportBatch<TItem>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default)
        where TItem : class, new()
    {
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedBatchRequest(request);
        ValidateBatchPlan(CreatePlan(request));
        var path = await StageInputAsync(source, request.ResourceLimits, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            await using var input = OpenStagedInput(path);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var resourceError = ToResourceError(preflightException);
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(1, 1,
                    Array.Empty<TItem>(), new[] { resourceError }), cancellationToken).ConfigureAwait(false);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(1, budgetError.RowIndex,
                    Array.Empty<TItem>(), new[] { budgetError }), cancellationToken).ConfigureAwait(false);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            using var reader = CreateReader(input);
            return await ReadBatchesAsync(reader, request, onBatch, cancellationToken)
                .ConfigureAwait(false);
        }
        catch (BatchCallbackException exception)
        {
            ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("ExcelDataReader 异步分批导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <summary>
    /// 读取工作簿并汇总工作表导入结果。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>工作簿数据、工作表结果和错误信息。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportReader<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var resourceError = ScanWorkbookResourceBudgets(source, request.ResourceLimits, cancellationToken);
        if (resourceError != null)
        {
            var limitedErrors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
            limitedErrors.Add(resourceError);
            return ResourceLimitedResult(new TWorkbook(), limitedErrors);
        }
        using var reader = CreateReader(source);
        var root = new TWorkbook();
        var sheetResults = new List<ExcelSheetImportResult>();
        var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        var rowBudget = new RowBudget(request.ResourceLimits?.MaxRows);
        // XLSX 的物理 Cell/列限制已经由 ZIP 预检按 XML 节点精确执行；这里不再用 FieldCount 重复估算。
        var cellBudget = new CellBudget(null);
        var physicalSheetIndex = 0;
        var matched = new HashSet<ExcelSheetImportRequest>();
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.Name == null)
                break;
            if (request.ResourceLimits?.MaxSheets is int maxSheets && physicalSheetIndex >= maxSheets)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    "工作表数量超过资源限制。", reader.Name, 1, 0, null));
                return ResourceLimitedResult(root, errors);
            }
            var selected = request.Sheets.Where(sheet => IsSelected(sheet, reader.Name, physicalSheetIndex,
                request.SheetNameComparison)).ToArray();
            if (selected.Length > 0)
            {
                foreach (var sheetRequest in selected)
                {
                    var sheetResult = ReadSheet(reader, sheetRequest, root, rowBudget, cellBudget, errors,
                        request, cancellationToken);
                    sheetResults.Add(sheetResult);
                    matched.Add(sheetRequest);
                    if (rowBudget.Exceeded || cellBudget.Exceeded || errors.Truncated || errors.HasResourceLimit)
                        return ResourceLimitedResult(root, errors);
                }
            }
            physicalSheetIndex++;
        } while (reader.NextResult());

        foreach (var sheetRequest in request.Sheets.Where(sheet => !matched.Contains(sheet)))
        {
            errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                $"未找到工作表: {sheetRequest.Name}", sheetRequest.Name, 1, 0, null));
        }
        return new ExcelWorkbookImportResult<TWorkbook>(root, sheetResults, errors.Items,
            errors.Truncated, request.ResourceLimits?.MaxErrors);
    }

    /// <summary>
    /// 读取并校验单个工作表的数据。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="root">工作簿目标对象。</param>
    /// <param name="rowBudget">跨工作表共享的数据行预算。</param>
    /// <param name="cellBudget">列数和单元格数量预算。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="workbookRequest">工作簿级导入请求。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>当前工作表的成功行索引和错误信息。</returns>
    private ExcelSheetImportResult ReadSheet<TWorkbook>(DataReader reader, ExcelSheetImportRequest request,
        TWorkbook root, RowBudget rowBudget, CellBudget cellBudget, ErrorCollector errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var plan = CreatePlan(request);
        if (plan.DynamicColumns.Count > 0 || request.DynamicColumns.Count > 0)
            throw Unsupported("动态列");
        var headers = ReadHeader(reader, request, cancellationToken);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            reader.Name, out var headerResourceError))
        {
            errors.Add(headerResourceError);
            return new ExcelSheetImportResult(reader.Name, request.ItemType, Array.Empty<int>(),
                new[] { headerResourceError });
        }
        var bindings = BindColumns(plan, headers, request, reader.Name, errors);
        var items = new List<object>();
        var rows = new List<int>();
        var sheetErrors = new List<ExcelImportError>();
        var unique = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var zeroBasedRowIndex = request.HeaderRowIndex;
        var physicalRow = request.HeaderRowIndex + 1;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            zeroBasedRowIndex++;
            physicalRow = zeroBasedRowIndex + 1;
            if (!cellBudget.TryConsume(reader.FieldCount, physicalRow, reader.Name,
                out var cellResourceError))
            {
                errors.Add(cellResourceError);
                sheetErrors.Add(cellResourceError);
                break;
            }
            if (zeroBasedRowIndex < request.DataRowStartIndex)
                continue;
            if (!rowBudget.TryConsume())
                break;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeWorkbookRow<TWorkbook>(reader, request, plan, bindings, reader.Name, physicalRow,
                unique, rowErrors, workbookRequest.ResourceLimits, cancellationToken);
            if (rowErrors.Count > 0)
            {
                foreach (var error in rowErrors)
                {
                    sheetErrors.Add(error);
                    errors.Add(error);
                }
                if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    continue;
            }
            else if (item != null)
            {
                items.Add(item);
                rows.Add(zeroBasedRowIndex);
            }
            if (errors.Truncated)
                break;
        }
        if (rowBudget.Exceeded)
        {
            var resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                "Workbook 数据行数超过资源限制。", reader.Name, physicalRow, 0, null);
            errors.Add(resourceError);
            return new ExcelSheetImportResult(reader.Name, request.ItemType, Array.Empty<int>(),
                new[] { resourceError });
        }
        AddItems(root, request.Target(root), items);
        return new ExcelSheetImportResult(reader.Name, request.ItemType, rows, sheetErrors);
    }

    /// <summary>
    /// 读取工作表并异步交付导入批次。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="onBatch">接收已转换数据与错误的批次回调。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已读取、成功、失败和资源限制状态的统计结果。</returns>
    private async Task<ExcelBatchImportSummary> ReadBatchesAsync<TItem>(DataReader reader,
        ExcelBatchImportRequest<TItem> request,
        Func<ExcelImportBatch<TItem>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken)
        where TItem : class, new()
    {
        if (!MoveToSheet(reader, request.SheetName, cancellationToken))
            throw new BingOfficesImportException($"未找到工作表: {request.SheetName}", provider: Provider,
                stage: BingOfficesStage.Read);
        var plan = CreatePlan(request);
        ValidateBatchPlan(plan);
        var headers = ReadHeader(reader, request.HeaderRowIndex, request.DataRowStartIndex, request.Culture,
            request.HeaderWhitespace, cancellationToken);
        var collector = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        // XLSX 的物理 Cell/列限制已经由 ZIP 预检按 XML 节点精确执行；这里不再用 FieldCount 重复估算。
        var cellBudget = new CellBudget(null);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            request.SheetName, out var headerResourceError))
        {
            collector.Add(headerResourceError);
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()),
                cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, true);
        }
        var bindings = BindColumns(plan, headers, request.RequireExpectedHeaders, request.SheetName,
            request.HeaderComparison, collector);
        if (collector.Items.Count > 0)
        {
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()),
                cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, false);
        }
        var unique = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var budget = new RowBudget(request.ResourceLimits?.MaxRows);
        var batch = new List<TItem>();
        var batchErrors = new List<ExcelImportError>();
        var batchFirstRow = 0;
        var batchNumber = 1;
        var rowsRead = 0;
        var succeeded = 0;
        var failed = 0;
        var zeroBasedRowIndex = request.HeaderRowIndex;
        var physicalRow = request.HeaderRowIndex + 1;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            zeroBasedRowIndex++;
            physicalRow = zeroBasedRowIndex + 1;
            if (!cellBudget.TryConsume(reader.FieldCount, physicalRow, request.SheetName,
                out var cellResourceError))
            {
                collector.Add(cellResourceError);
                if (batch.Count > 0 || batchErrors.Count > 0)
                    await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(batchNumber++, batchFirstRow,
                        batch.ToArray(), batchErrors.ToArray()), cancellationToken).ConfigureAwait(false);
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(batchNumber, physicalRow,
                    Array.Empty<TItem>(), new[] { cellResourceError }), cancellationToken).ConfigureAwait(false);
                batch.Clear();
                batchErrors.Clear();
                break;
            }
            if (zeroBasedRowIndex < request.DataRowStartIndex)
                continue;
            if (!budget.TryConsume())
                break;
            rowsRead++;
            batchFirstRow = batchFirstRow == 0 ? physicalRow : batchFirstRow;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeTypedRow<TItem>(reader, request, plan, bindings, request.SheetName,
                physicalRow, unique, rowErrors, cancellationToken);
            if (rowErrors.Count == 0 && item != null)
            {
                batch.Add(item);
                succeeded++;
            }
            else
            {
                failed++;
                batchErrors.AddRange(rowErrors);
            }
            if (batch.Count >= request.BatchSize || batchErrors.Count >= request.BatchSize)
            {
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(batchNumber++, batchFirstRow,
                    batch.ToArray(), batchErrors.ToArray()), cancellationToken).ConfigureAwait(false);
                batch.Clear();
                batchErrors.Clear();
                batchFirstRow = 0;
            }
            foreach (var error in rowErrors)
                collector.Add(error);
            if (collector.Truncated || (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                && rowErrors.Count > 0))
                break;
        }
        if (batch.Count > 0 || batchErrors.Count > 0)
        {
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(batchNumber, batchFirstRow,
                batch.ToArray(), batchErrors.ToArray()), cancellationToken).ConfigureAwait(false);
        }
        if (budget.Exceeded)
            collector.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                "Workbook 数据行数超过资源限制。", request.SheetName, physicalRow, 0, null));
        return new ExcelBatchImportSummary(rowsRead, succeeded, failed, collector.Items.Count,
            collector.Truncated, budget.Exceeded || cellBudget.Exceeded);
    }

    /// <summary>
    /// 读取工作表并交付导入批次。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="onBatch">接收已转换数据与错误的批次回调。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已读取、成功、失败和资源限制状态的统计结果。</returns>
    private ExcelBatchImportSummary ReadBatches<TItem>(DataReader reader,
        ExcelBatchImportRequest<TItem> request, Action<ExcelImportBatch<TItem>> onBatch,
        CancellationToken cancellationToken)
        where TItem : class, new()
    {
        if (!MoveToSheet(reader, request.SheetName, cancellationToken))
            throw new BingOfficesImportException($"未找到工作表: {request.SheetName}", provider: Provider,
                stage: BingOfficesStage.Read);
        var plan = CreatePlan(request);
        ValidateBatchPlan(plan);
        var headers = ReadHeader(reader, request.HeaderRowIndex, request.DataRowStartIndex, request.Culture,
            request.HeaderWhitespace, cancellationToken);
        var collector = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        var cellBudget = new CellBudget(request.ResourceLimits);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            request.SheetName, out var headerResourceError))
        {
            collector.Add(headerResourceError);
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, true);
        }
        var bindings = BindColumns(plan, headers, request.RequireExpectedHeaders, request.SheetName,
            request.HeaderComparison, collector);
        if (collector.Items.Count > 0)
        {
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, false);
        }
        var unique = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        var budget = new RowBudget(request.ResourceLimits?.MaxRows);
        var batch = new List<TItem>();
        var batchErrors = new List<ExcelImportError>();
        var batchFirstRow = 0;
        var batchNumber = 1;
        var rowsRead = 0;
        var succeeded = 0;
        var failed = 0;
        var zeroBasedRowIndex = request.HeaderRowIndex;
        var physicalRow = request.HeaderRowIndex + 1;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            zeroBasedRowIndex++;
            physicalRow = zeroBasedRowIndex + 1;
            if (!cellBudget.TryConsume(reader.FieldCount, physicalRow, request.SheetName,
                out var cellResourceError))
            {
                collector.Add(cellResourceError);
                if (batch.Count > 0 || batchErrors.Count > 0)
                    InvokeBatch(onBatch, new ExcelImportBatch<TItem>(batchNumber++, batchFirstRow,
                        batch.ToArray(), batchErrors.ToArray()));
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(batchNumber, physicalRow,
                    Array.Empty<TItem>(), new[] { cellResourceError }));
                batch.Clear();
                batchErrors.Clear();
                break;
            }
            if (zeroBasedRowIndex < request.DataRowStartIndex)
                continue;
            if (!budget.TryConsume())
                break;
            rowsRead++;
            batchFirstRow = batchFirstRow == 0 ? physicalRow : batchFirstRow;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeTypedRow<TItem>(reader, request, plan, bindings, request.SheetName,
                physicalRow, unique, rowErrors, cancellationToken);
            if (rowErrors.Count == 0 && item != null)
            {
                batch.Add(item);
                succeeded++;
            }
            else
            {
                failed++;
                batchErrors.AddRange(rowErrors);
            }
            foreach (var error in rowErrors)
                collector.Add(error);
            if (batch.Count >= request.BatchSize || batchErrors.Count >= request.BatchSize)
            {
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(batchNumber++, batchFirstRow,
                    batch.ToArray(), batchErrors.ToArray()));
                batch.Clear();
                batchErrors.Clear();
                batchFirstRow = 0;
            }
            if (collector.Truncated || (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                && rowErrors.Count > 0))
                break;
        }
        if (batch.Count > 0 || batchErrors.Count > 0)
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(batchNumber, batchFirstRow,
                batch.ToArray(), batchErrors.ToArray()));
        if (budget.Exceeded)
            collector.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                "Workbook 数据行数超过资源限制。", request.SheetName, physicalRow, 0, null));
        return new ExcelBatchImportSummary(rowsRead, succeeded, failed, collector.Items.Count,
            collector.Truncated, budget.Exceeded || cellBudget.Exceeded);
    }

    /// <summary>
    /// 创建工作表导入映射计划。
    /// </summary>
    /// <param name="request">导入请求。</param>
    /// <returns>用于实体导入的映射计划。</returns>
    private IExcelMappingPlan CreatePlan(ExcelSheetImportRequest request)
    {
        var document = request.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true };
        var method = typeof(ExcelDataReaderExcelImporter).GetMethod(nameof(CreatePlanTyped),
            BindingFlags.Instance | BindingFlags.NonPublic).MakeGenericMethod(request.ItemType);
        return (IExcelMappingPlan)method.Invoke(this, new object[] { document, request.MappingConfiguration });
    }

    /// <summary>
    /// 创建工作表导入映射计划。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="request">导入请求。</param>
    /// <returns>用于实体导入的映射计划。</returns>
    private IExcelMappingPlan CreatePlan<TItem>(ExcelBatchImportRequest<TItem> request)
        where TItem : class, new()
    {
        return _mappingPlanFactory.Create<TItem>(request.MappingDocument ??
            new ExcelMappingDocument { UseConventionFallback = true }, request.MappingConfiguration,
            MappingDirection.Import);
    }

    /// <summary>
    /// 创建指定实体类型的导入映射计划。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="document">映射文档。</param>
    /// <param name="configuration">运行时映射配置。</param>
    /// <returns>用于实体导入的映射计划。</returns>
    private IExcelMappingPlan CreatePlanTyped<TItem>(ExcelMappingDocument document,
        ExcelMappingConfiguration configuration) where TItem : class, new()
        => _mappingPlanFactory.Create<TItem>(document, configuration, MappingDirection.Import);

    /// <summary>
    /// 读取并规范化工作表表头。
    /// </summary>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>按物理列顺序排列的规范化表头。</returns>
    private static IReadOnlyList<string> ReadHeader(DataReader reader, ExcelSheetImportRequest request,
        CancellationToken cancellationToken)
        => ReadHeader(reader, request.HeaderRowIndex, request.DataRowStartIndex, request.Culture,
            request.HeaderWhitespace, cancellationToken);

    /// <summary>
    /// 读取并规范化工作表表头。
    /// </summary>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="headerRowIndex">从 0 开始的表头行索引。</param>
    /// <param name="dataRowStartIndex">请求中的数据起始行索引；本方法不使用该值推进读取器。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="whitespace">表头空白处理策略。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>按物理列顺序排列的规范化表头。</returns>
    private static IReadOnlyList<string> ReadHeader(DataReader reader, int headerRowIndex, int dataRowStartIndex,
        CultureInfo culture, ExcelWhitespacePolicy whitespace, CancellationToken cancellationToken)
    {
        var row = 0;
        while (row <= headerRowIndex)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.Read())
                throw new BingOfficesImportException("工作表缺少表头行。", provider: Provider,
                    stage: BingOfficesStage.Read);
            row++;
        }
        var result = new List<string>();
        for (var index = 0; index < reader.FieldCount; index++)
            result.Add(NormalizeText(ReadText(reader, index, culture), whitespace));
        return result;
    }

    /// <summary>
    /// 将映射列绑定到匹配的表头位置。
    /// </summary>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="headers">按物理列顺序排列的表头文本。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <returns>成功匹配表头的列绑定集合。</returns>
    private static IReadOnlyList<ColumnBinding> BindColumns(IExcelMappingPlan plan,
        IReadOnlyList<string> headers, ExcelSheetImportRequest request, string sheetName,
        ErrorCollector errors)
        => BindColumns(plan, headers, request.RequireExpectedHeaders, sheetName, request.HeaderComparison, errors);

    /// <summary>
    /// 将映射列绑定到匹配的表头位置。
    /// </summary>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="headers">按物理列顺序排列的表头文本。</param>
    /// <param name="requireExpectedHeaders">缺少期望表头时是否记录错误。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <returns>成功匹配表头的列绑定集合。</returns>
    private static IReadOnlyList<ColumnBinding> BindColumns(IExcelMappingPlan plan,
        IReadOnlyList<string> headers, bool requireExpectedHeaders, string sheetName,
        ExcelNameComparison comparison, ErrorCollector errors)
    {
        var result = new List<ColumnBinding>();
        foreach (var column in plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn))
        {
            var index = FindHeader(headers, column.Title, column.Aliases, comparison);
            if (index < 0)
            {
                if (requireExpectedHeaders)
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                        $"缺少表头: {column.Title}", sheetName, 1, 0, column.Name,
                        column.Name, column.Title));
                continue;
            }
            result.Add(new ColumnBinding(column, index, headers[index]));
        }
        return result;
    }

    /// <summary>
    /// 将映射列绑定到匹配的表头位置。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="headers">按物理列顺序排列的表头文本。</param>
    /// <param name="requireExpectedHeaders">缺少期望表头时是否记录错误。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <returns>成功匹配表头的列绑定集合。</returns>
    private static IReadOnlyList<ColumnBinding> BindColumns<TItem>(IExcelMappingPlan plan,
        IReadOnlyList<string> headers, bool requireExpectedHeaders, string sheetName,
        ExcelNameComparison comparison, ErrorCollector errors) where TItem : class, new()
        => BindColumns(plan, headers, requireExpectedHeaders, sheetName, comparison, errors);

    /// <summary>
    /// 将工作表当前行转换为请求指定的实体。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="bindings">当前处理的列绑定或校验绑定集合。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="unique">按属性名维护的唯一值集合。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static object MaterializeWorkbookRow<TWorkbook>(DataReader reader, ExcelSheetImportRequest request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        IDictionary<string, HashSet<string>> unique, ICollection<ExcelImportError> errors,
        ExcelResourceLimits resourceLimits, CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var method = typeof(ExcelDataReaderExcelImporter).GetMethod(nameof(MaterializeWorkbookRowCore),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(request.ItemType);
        return method.Invoke(null, new object[] { reader, request, plan, bindings, sheetName, rowNumber,
            unique, errors, resourceLimits, cancellationToken });
    }

    /// <summary>
    /// 转换并校验当前行的实体数据。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="bindings">当前处理的列绑定或校验绑定集合。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="unique">按属性名维护的唯一值集合。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static TItem MaterializeTypedRow<TItem>(DataReader reader, ExcelBatchImportRequest<TItem> request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        IDictionary<string, HashSet<string>> unique, ICollection<ExcelImportError> errors,
        CancellationToken cancellationToken) where TItem : class, new()
    {
        var item = new TItem();
        var valid = true;
        var initialErrorCount = errors.Count;
        var reservations = new List<UniqueReservation>();
        var validationEnabled = request.ValidationMode == ExcelImportValidationMode.ConfiguredRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook;
        foreach (var binding in bindings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var raw = binding.Index < reader.FieldCount ? reader.GetValue(binding.Index) : null;
            var text = NormalizeText(ReadText(raw, request.Culture),
                binding.Column.ImportWhitespace ?? request.BodyWhitespace);
            var property = typeof(TItem).GetProperty(binding.Column.Name,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (property == null || !property.CanWrite)
                continue;
            var cell = CreateCell(raw, text);
            try
            {
                    if (validationEnabled)
                        ValidateBindings(binding.Column.ValidationBindings, text, null, binding, sheetName,
                        rowNumber, raw, property.PropertyType, cell, request.Culture, errors,
                        unique, request.ResourceLimits, reservations, rawOnly: true);
                var converted = ConvertValue(raw, text, binding.Column, property.PropertyType,
                    sheetName, rowNumber, binding.Index + 1, request.Culture, cell);
                    if (validationEnabled)
                        ValidateBindings(binding.Column.ValidationBindings, text, converted, binding, sheetName,
                        rowNumber, raw, property.PropertyType, cell, request.Culture, errors,
                        unique, request.ResourceLimits, reservations, rawOnly: false);
                property.SetValue(item, converted);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                if (!errors.Any(error => error.RowIndex == rowNumber && error.ColumnIndex == binding.Index + 1
                    && error.PropertyName == binding.Column.Name))
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                        sheetName, rowNumber, binding.Index + 1, binding.Column.Name,
                        binding.Column.Name, binding.Header, raw));
                valid = false;
                if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    break;
            }
        }
        if (!valid || errors.Count > initialErrorCount)
        {
            foreach (var reservation in reservations)
                reservation.Values.Remove(reservation.Value);
            return null;
        }
        return item;
    }

    /// <summary>
    /// 将工作表请求适配为实体行转换请求。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="bindings">当前处理的列绑定或校验绑定集合。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="unique">按属性名维护的唯一值集合。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static object MaterializeWorkbookRowCore<TItem>(DataReader reader, ExcelSheetImportRequest request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        IDictionary<string, HashSet<string>> unique, ICollection<ExcelImportError> errors,
        ExcelResourceLimits resourceLimits, CancellationToken cancellationToken) where TItem : class, new()
    {
        var requestAdapter = new ExcelBatchImportRequest<TItem>(sheetName)
        {
            HeaderRowIndex = request.HeaderRowIndex,
            DataRowStartIndex = request.DataRowStartIndex,
            ValidationFailureMode = request.ValidationFailureMode,
            ValidationMode = ExcelImportValidationMode.ConfiguredRules,
            Culture = request.Culture,
            MappingConfiguration = request.MappingConfiguration,
            MappingDocument = request.MappingDocument,
            BodyWhitespace = request.BodyWhitespace,
            ResourceLimits = resourceLimits
        };
        return MaterializeTypedRow(reader, requestAdapter, plan, bindings, sheetName, rowNumber,
            unique, errors, cancellationToken);
    }

    /// <summary>
    /// 执行列校验并预留当前行的唯一值。
    /// </summary>
    /// <param name="bindings">当前处理的列绑定或校验绑定集合。</param>
    /// <param name="text">按空白策略处理后的单元格文本。</param>
    /// <param name="converted">转换后的属性值。</param>
    /// <param name="binding">当前列或校验规则绑定。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="propertyType">目标实体属性类型。</param>
    /// <param name="cell">单元格原始值和类型上下文。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="unique">按属性名维护的唯一值集合。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="reservations">当前行新增唯一值的回滚记录集合。</param>
    /// <param name="rawOnly">为 true 时只执行原始值校验；否则执行转换后校验和唯一性检查。</param>
    private static void ValidateBindings(IReadOnlyList<IExcelValidationBinding> bindings, string text,
        object converted, ColumnBinding binding, string sheetName, int rowNumber, object raw, Type propertyType,
        ExcelCellValue cell, CultureInfo culture, ICollection<ExcelImportError> errors,
        IDictionary<string, HashSet<string>> unique, ExcelResourceLimits resourceLimits,
        ICollection<UniqueReservation> reservations, bool rawOnly)
    {
        foreach (var validation in bindings ?? Array.Empty<IExcelValidationBinding>())
        {
            if (validation.Kind == ExcelValidationBindingKind.Unique || validation.IsRaw != rawOnly)
                continue;
            var context = new ExcelValidationContext(text, sheetName, rowNumber, binding.Index + 1,
                binding.Column.Name, converted, propertyType, cell, culture);
            if (!validation.Validate(context))
                errors.Add(new ExcelImportError(ErrorCode(validation), validation.ErrorMessage, sheetName,
                    rowNumber, binding.Index + 1, binding.Column.Name, binding.Column.Name, binding.Header, raw));
        }
        if (!rawOnly && binding.Column.IsUnique)
        {
            if (!unique.TryGetValue(binding.Column.Name, out var values))
                unique[binding.Column.Name] = values = new HashSet<string>(
                    GetStringComparer(resourceLimits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrWhiteSpace(text) && !values.Contains(text))
            {
                if (resourceLimits?.MaxTrackedUniqueValues is int maximum
                    && values.Count >= maximum)
                {
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                        $"唯一值跟踪数量超过资源限制: {maximum}", sheetName, rowNumber,
                        binding.Index + 1, binding.Column.Name, binding.Column.Name, binding.Header, raw));
                }
                else
                {
                    values.Add(text);
                    reservations.Add(new UniqueReservation(values, text));
                }
            }
            else if (!string.IsNullOrWhiteSpace(text))
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, "重复数据。", sheetName,
                    rowNumber, binding.Index + 1, binding.Column.Name, binding.Column.Name, binding.Header, raw));
            }
        }
    }

    /// <summary>
    /// 将单元格值转换为目标属性类型。
    /// </summary>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="text">按空白策略处理后的单元格文本。</param>
    /// <param name="column">对应列的映射配置。</param>
    /// <param name="propertyType">目标实体属性类型。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="columnIndex">用于错误定位的列号。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <param name="cell">单元格原始值和类型上下文。</param>
    /// <returns>转换后的值；允许空值且输入为空，或转换器返回空值时为 null。</returns>
    private static object ConvertValue(object raw, string text, IExcelMappingColumn column, Type propertyType,
        string sheetName, int rowNumber, int columnIndex, CultureInfo culture, ExcelCellValue cell)
    {
        var context = new ExcelConversionContext(raw, column.Name, propertyType, sheetName, rowNumber,
            columnIndex, culture, cell);
        foreach (var converter in column.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertFrom(context, out var converted))
                return converted;
        if (column.ValueMap != null && column.ValueMap.TryGetValue(text, out var mapped))
            text = mapped;
        var target = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (string.IsNullOrWhiteSpace(text))
        {
            if (!target.IsValueType || Nullable.GetUnderlyingType(propertyType) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {propertyType.FullName}");
        }
        if (target == typeof(string))
            return text;
        if (target.IsEnum)
            return Enum.Parse(target, text, true);
        if (target == typeof(Guid))
            return Guid.Parse(text);
        if (target == typeof(DateTime))
            return raw is DateTime ? raw : DateTime.Parse(text, culture);
        if (target == typeof(DateTimeOffset))
            return raw is DateTimeOffset ? raw : DateTimeOffset.Parse(text, culture);
        if (target == typeof(TimeSpan))
            return raw is TimeSpan ? raw : TimeSpan.Parse(text, culture);
        return Convert.ChangeType(raw is string ? text : raw, target, culture);
    }

    /// <summary>
    /// 获取指定字符串比较策略的比较器。
    /// </summary>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <returns>与策略对应的字符串比较器。</returns>
    private static StringComparer GetStringComparer(StringComparison comparison)
        => comparison switch
        {
            StringComparison.Ordinal => StringComparer.Ordinal,
            StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
            StringComparison.InvariantCulture => StringComparer.InvariantCulture,
            StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
            StringComparison.CurrentCulture => StringComparer.CurrentCulture,
            StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
            _ => StringComparer.OrdinalIgnoreCase
        };

    /// <summary>
    /// 将读取器移到指定名称的工作表。
    /// </summary>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="name">待匹配的工作表名称。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>找到目标工作表时为 true，否则为 false。</returns>
    private static bool MoveToSheet(DataReader reader, string name, CancellationToken cancellationToken)
    {
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.Equals(reader.Name, name, StringComparison.OrdinalIgnoreCase))
                return true;
        } while (reader.NextResult());
        return false;
    }

    /// <summary>
    /// 判断工作表是否匹配请求的选择条件。
    /// </summary>
    /// <param name="request">导入请求。</param>
    /// <param name="name">待匹配的工作表名称。</param>
    /// <param name="index">从 0 开始的物理索引。</param>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <returns>名称或索引匹配选择条件时为 true，否则为 false。</returns>
    private static bool IsSelected(ExcelSheetImportRequest request, string name, int index,
        ExcelNameComparison comparison)
        => request.Selector.Kind == ExcelSheetSelectorKind.ByIndex
            ? request.Selector.Index == index
            : string.Equals(request.Selector.Name, name,
                comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 查找标题或别名匹配的表头位置。
    /// </summary>
    /// <param name="headers">按物理列顺序排列的表头文本。</param>
    /// <param name="title">导出列标题。</param>
    /// <param name="aliases">允许匹配的表头别名。</param>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <returns>从 0 开始的匹配列索引；未找到时为 -1。</returns>
    private static int FindHeader(IReadOnlyList<string> headers, string title, IReadOnlyList<string> aliases,
        ExcelNameComparison comparison)
    {
        var candidates = new[] { title }.Concat(aliases ?? Array.Empty<string>()).Where(value => !string.IsNullOrWhiteSpace(value));
        for (var index = 0; index < headers.Count; index++)
            if (candidates.Any(candidate => string.Equals(headers[index], candidate,
                comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase)))
                return index;
        return -1;
    }

    /// <summary>
    /// 按指定区域性读取单元格文本。
    /// </summary>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="index">从 0 开始的物理索引。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <returns>区域性格式化后的文本；空单元格返回空字符串。</returns>
    private static string ReadText(DataReader reader, int index, CultureInfo culture)
        => index < reader.FieldCount ? ReadText(reader.GetValue(index), culture) : string.Empty;

    /// <summary>
    /// 按指定区域性读取单元格文本。
    /// </summary>
    /// <param name="value">待处理的值。</param>
    /// <param name="culture">值转换和格式化使用的区域性。</param>
    /// <returns>区域性格式化后的文本；空单元格返回空字符串。</returns>
    private static string ReadText(object value, CultureInfo culture)
        => value == null || value == DBNull.Value ? string.Empty : value is IFormattable formattable
            ? formattable.ToString(null, culture) : Convert.ToString(value, culture);

    /// <summary>
    /// 按空白策略规范化文本。
    /// </summary>
    /// <param name="value">待处理的值。</param>
    /// <param name="policy">文本空白处理策略。</param>
    /// <returns>按策略保留、去除两端或去除全部空白的文本。</returns>
    private static string NormalizeText(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.RemoveAll => string.Concat(value.Where(character => !char.IsWhiteSpace(character))),
            _ => value.Trim()
        };
    }

    /// <summary>
    /// 创建包含原始值、文本和类型的单元格值。
    /// </summary>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="text">按空白策略处理后的单元格文本。</param>
    /// <returns>单元格原始值、文本和类型上下文。</returns>
    private static ExcelCellValue CreateCell(object raw, string text)
    {
        var kind = raw switch
        {
            null => ExcelCellKind.Empty,
            bool => ExcelCellKind.Boolean,
            DateTime => ExcelCellKind.DateTime,
            DateTimeOffset => ExcelCellKind.DateTime,
            string => ExcelCellKind.Text,
            _ when raw is IConvertible => ExcelCellKind.Number,
            _ => ExcelCellKind.Text
        };
        return new ExcelCellValue(raw, text, kind);
    }

    /// <summary>
    /// 获取校验规则对应的导入错误码。
    /// </summary>
    /// <param name="binding">当前列或校验规则绑定。</param>
    /// <returns>对应的导入错误分类。</returns>
    private static ExcelImportErrorCode ErrorCode(IExcelValidationBinding binding)
        => binding.Kind switch
        {
            ExcelValidationBindingKind.MaxLength => ExcelImportErrorCode.MaxLength,
            ExcelValidationBindingKind.MaxValue => ExcelImportErrorCode.MaxValue,
            _ => ExcelImportErrorCode.Validation
        };

    /// <summary>
    /// 向目标集合添加导入实体。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="root">工作簿目标对象。</param>
    /// <param name="target">接收导入实体的目标集合。</param>
    /// <param name="items">待添加的实体集合。</param>
    private static void AddItems<TWorkbook>(TWorkbook root, object target, IEnumerable<object> items)
        where TWorkbook : class, new()
    {
        if (target is IList list)
        {
            foreach (var item in items)
                list.Add(item);
            return;
        }
        var add = target?.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public)
            .Where(method => string.Equals(method.Name, "Add", StringComparison.Ordinal)
                && method.GetParameters().Length == 1)
            .FirstOrDefault();
        if (add == null)
            throw new InvalidOperationException("Sheet 目标集合必须支持 Add 操作。");
        foreach (var item in items)
            add.Invoke(target, new[] { item });
    }

    /// <summary>
    /// 创建不保留部分数据的资源受限结果。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="root">工作簿目标对象。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <returns>包含新建空工作簿对象和已收集错误的受限结果。</returns>
    private static ExcelWorkbookImportResult<TWorkbook> ResourceLimitedResult<TWorkbook>(TWorkbook root,
        ErrorCollector errors) where TWorkbook : class, new()
        => new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(), Array.Empty<ExcelSheetImportResult>(),
            errors.Items, true, errors.MaxErrors);

    /// <summary>
    /// 验证输入流和导入请求参数。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="request">导入请求。</param>
    private static void ValidateArguments<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request) where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        request.ResourceLimits?.Validate();
    }

    /// <summary>
    /// 验证输入流和导入请求参数。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="request">导入请求。</param>
    private static void ValidateArguments<TItem>(Stream source,
        ExcelBatchImportRequest<TItem> request) where TItem : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        request.Validate();
    }

    /// <summary>
    /// 验证工作簿请求是否受只读导入器支持。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="request">导入请求。</param>
    private static void ValidateSupportedRequest<TWorkbook>(ExcelWorkbookImportRequest<TWorkbook> request)
        where TWorkbook : class, new()
    {
        if (request.Relations.Count > 0)
            throw Unsupported("跨 Sheet 关系");
        if (request.FailureOptions != null
            && request.FailureOptions.Mode != ExcelImportFailureWorkbookMode.None)
            throw Unsupported("Failure Workbook");
        if (request.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook)
            throw Unsupported("Workbook 原生校验");
        if (request.Sheets.Any(sheet => sheet.DynamicColumns.Count > 0))
            throw Unsupported("动态列");
        if (request.ResourceLimits != null
            && (request.ResourceLimits.MaxPictures.HasValue
                || request.ResourceLimits.MaxPictureBytes.HasValue
                || request.ResourceLimits.MaxTotalPictureBytes.HasValue))
            throw Unsupported("图片资源限制");
    }

    /// <summary>
    /// 验证请求是否受分批导入器支持。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="request">导入请求。</param>
    private static void ValidateSupportedBatchRequest<TItem>(ExcelBatchImportRequest<TItem> request)
        where TItem : class, new()
    {
        if (request.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook)
            throw Unsupported("Workbook 原生校验");
        if (request.ResourceLimits != null
            && (request.ResourceLimits.MaxPictures.HasValue
                || request.ResourceLimits.MaxPictureBytes.HasValue
                || request.ResourceLimits.MaxTotalPictureBytes.HasValue
                || request.ResourceLimits.MaxTrackedUniqueValues.HasValue))
            throw Unsupported("图片资源限制或跨行唯一性");
    }

    /// <summary>
    /// 验证分批导入计划不包含动态列或跨行唯一性。
    /// </summary>
    /// <param name="plan">实体导入映射计划。</param>
    private static void ValidateBatchPlan(IExcelMappingPlan plan)
    {
        if (plan.DynamicColumns.Count > 0 || plan.Columns.Any(column => column.IsDynamicColumn))
            throw Unsupported("动态列");
        if (plan.DynamicColumns.Any(column => column.IsUnique)
            || plan.Columns.Any(column => column.IsUnique
                || (column.ValidationBindings ?? Array.Empty<IExcelValidationBinding>())
                    .Any(binding => binding.Kind == ExcelValidationBindingKind.Unique)))
            throw Unsupported("跨行唯一性");
    }

    /// <summary>
    /// 创建不支持指定导入功能的异常。
    /// </summary>
    /// <param name="feature">不支持的功能名称。</param>
    /// <returns>带导入和预检阶段信息的不支持功能异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string feature)
        => new BingOfficesUnsupportedFeatureException($"ExcelDataReader 不支持功能: {feature}",
            provider: Provider, operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);

    /// <summary>
    /// 交付导入批次并包装回调异常。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="callback">接收导入批次的调用方回调。</param>
    /// <param name="batch">包含数据和错误的导入批次。</param>
    private static void InvokeBatch<TItem>(Action<ExcelImportBatch<TItem>> callback,
        ExcelImportBatch<TItem> batch) where TItem : class, new()
    {
        try
        {
            callback(batch);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new BatchCallbackException(exception);
        }
    }

    /// <summary>
    /// 异步交付导入批次并包装回调异常。
    /// </summary>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="callback">接收导入批次的调用方回调。</param>
    /// <param name="batch">包含数据和错误的导入批次。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static async Task InvokeBatchAsync<TItem>(
        Func<ExcelImportBatch<TItem>, CancellationToken, Task> callback,
        ExcelImportBatch<TItem> batch, CancellationToken cancellationToken)
        where TItem : class, new()
    {
        try
        {
            await callback(batch, cancellationToken).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception exception)
        {
            throw new BatchCallbackException(exception);
        }
    }

    /// <summary>
    /// 预检工作簿的工作表数量限制。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="limits">本次导入的资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>工作表数量超限错误；未限制或未超限时为 null。</returns>
    private static ExcelImportError ScanWorkbookResourceBudgets(Stream source,
        ExcelResourceLimits limits, CancellationToken cancellationToken)
    {
        if (limits?.MaxSheets == null)
            return null;
        if (!source.CanSeek)
            throw new InvalidOperationException("ExcelDataReader 资源预检需要可定位的暂存输入。");

        ExcelImportError resourceError = null;
        var sheetIndex = 0;
        try
        {
            using (var reader = CreateReader(source, leaveOpen: true))
            {
                do
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (reader.Name == null)
                        break;
                    if (limits.MaxSheets is int maximumSheets && sheetIndex >= maximumSheets)
                    {
                        resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                            "工作表数量超过资源限制。", reader.Name, 1, 0, null);
                        break;
                    }
                    var physicalRow = 0;
                    while (reader.Read())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        physicalRow++;
                    }
                    if (resourceError != null)
                        break;
                    sheetIndex++;
                } while (reader.NextResult());
            }
        }
        finally
        {
            source.Position = 0;
        }
        return resourceError;
    }

    /// <summary>
    /// 执行输入格式和物理资源限制预检。
    /// </summary>
    /// <param name="input">待读取或预检的输入流。</param>
    /// <param name="limits">本次导入的资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void ValidateProviderPreflight(Stream input, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        var physicalCellLimitsRequested = limits?.MaxColumnsPerSheet != null || limits?.MaxCells != null;
        var isXlsx = physicalCellLimitsRequested && IsXlsxXmlWorkbook(input);
        if (physicalCellLimitsRequested && !isXlsx)
            throw Unsupported("XLS/XLSB 无法提供精确的物理单元格/列资源限制");

        ExcelXlsxZipPreflight.Validate(input, limits, Provider, requireZip: false,
            cancellationToken, enforceSheetLimit: false, enforceCellLimits: isXlsx,
            requireWorkbookXml: false);
    }

    /// <summary>
    /// 判断输入是否为 XML 形式的 XLSX 工作簿。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <returns>包含 XML 工作簿且不含二进制工作簿部件时为 true，否则为 false。</returns>
    private static bool IsXlsxXmlWorkbook(Stream source)
    {
        if (!source.CanSeek || source.Length < 4)
            return false;
        var originalPosition = source.Position;
        source.Position = 0;
        try
        {
            var signature = new byte[4];
            if (source.Read(signature, 0, signature.Length) != signature.Length
                || signature[0] != 0x50 || signature[1] != 0x4B)
                return false;
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            return archive.GetEntry("xl/workbook.xml") != null
                && archive.GetEntry("xl/workbook.bin") == null;
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 将资源超限异常转换为导入错误。
    /// </summary>
    /// <param name="exception">待判断或解包的异常。</param>
    /// <returns>资源超限分类的导入错误。</returns>
    private static ExcelImportError ToResourceError(BingOfficesResourceLimitException exception)
        => new ExcelImportError(ExcelImportErrorCode.ResourceLimit, exception.Message, null, 1, 0, null);

    /// <summary>
    /// 创建前向单遍工作簿读取器。
    /// </summary>
    /// <param name="input">待读取或预检的输入流。</param>
    /// <param name="leaveOpen">释放读取器后是否保留输入流。</param>
    /// <returns>配置为单遍读取的工作簿读取器。</returns>
    private static DataReader CreateReader(Stream input, bool leaveOpen = false)
        => DataReaderFactory.CreateReader(input, new DataReaderConfiguration
        {
            LeaveOpen = leaveOpen,
            FallbackEncoding = Encoding.GetEncoding(1252),
            SinglePassMode = true
        });

    /// <summary>
    /// 打开暂存文件用于顺序读取。
    /// </summary>
    /// <param name="path">输入暂存文件路径。</param>
    /// <returns>支持异步顺序读取的文件流。</returns>
    private static FileStream OpenStagedInput(string path)
        => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 64 * 1024,
            FileOptions.SequentialScan | FileOptions.Asynchronous);

    /// <summary>
    /// 将输入流暂存到受资源限制约束的文件。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="limits">本次导入的资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已写入输入内容的暂存文件路径。</returns>
    private string StageInput(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        var path = _stagedPathFactory();
        try
        {
            using var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                64 * 1024, FileOptions.SequentialScan);
            CopySync(source, destination, limits?.MaxInputBytes, cancellationToken);
            return path;
        }
        catch
        {
            DeleteStagedInput(path);
            throw;
        }
    }

    /// <summary>
    /// 异步将输入流暂存到受资源限制约束的文件。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="limits">本次导入的资源限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已写入输入内容的暂存文件路径。</returns>
    private async Task<string> StageInputAsync(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        var path = _stagedPathFactory();
        try
        {
            await using var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None,
                64 * 1024, FileOptions.Asynchronous | FileOptions.SequentialScan);
            await CopyAsync(source, destination, limits?.MaxInputBytes, cancellationToken).ConfigureAwait(false);
            return path;
        }
        catch
        {
            DeleteStagedInput(path);
            throw;
        }
    }

    /// <summary>
    /// 复制输入流并检查输入字节限制。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="maximum">允许消耗的数量上限；null 表示不限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void CopySync(Stream source, Stream destination, long? maximum,
        CancellationToken cancellationToken)
    {
        var buffer = new byte[64 * 1024];
        long total = 0;
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            total += read;
            if (maximum.HasValue && total > maximum.Value)
                throw new BingOfficesResourceLimitException("输入文件超过资源限制。", provider: Provider);
            destination.Write(buffer, 0, read);
        }
    }

    /// <summary>
    /// 异步复制输入流并检查输入字节限制。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="maximum">允许消耗的数量上限；null 表示不限制。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static async Task CopyAsync(Stream source, Stream destination, long? maximum,
        CancellationToken cancellationToken)
    {
        var buffer = new byte[64 * 1024];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
        {
            total += read;
            if (maximum.HasValue && total > maximum.Value)
                throw new BingOfficesResourceLimitException("输入文件超过资源限制。", provider: Provider);
            await destination.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// 生成默认输入暂存文件路径。
    /// </summary>
    /// <returns>系统临时目录中的唯一文件路径。</returns>
    private static string CreateDefaultStagedPath()
        => Path.Combine(Path.GetTempPath(), "bing-offices-excel-" + Guid.NewGuid().ToString("N") + ".tmp");

    /// <summary>
    /// 尝试删除输入暂存文件。
    /// </summary>
    /// <param name="path">输入暂存文件路径。</param>
    private static void DeleteStagedInput(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            return;
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch
        {
            // 清理失败不能覆盖原始导入异常。
        }
    }

    /// <summary>
    /// 映射列与物理表头的绑定。
    /// </summary>
    private sealed class ColumnBinding
    {
        /// <summary>
        /// 初始化一个 <see cref="ColumnBinding"/> 类型的实例。
        /// </summary>
        /// <param name="column">对应列的映射配置。</param>
        /// <param name="index">从 0 开始的物理索引。</param>
        /// <param name="header">匹配到的表头文本。</param>
        internal ColumnBinding(IExcelMappingColumn column, int index, string header)
        {
            Column = column;
            Index = index;
            Header = header;
        }

        /// <summary>
        /// 获取当前绑定的映射列。
        /// </summary>
        internal IExcelMappingColumn Column { get; }
        /// <summary>
        /// 获取从 0 开始的物理列索引。
        /// </summary>
        internal int Index { get; }
        /// <summary>
        /// 获取匹配到的表头文本。
        /// </summary>
        internal string Header { get; }
    }

    /// <summary>
    /// 当前行新增唯一值的回滚记录。
    /// </summary>
    private sealed class UniqueReservation
    {
        /// <summary>
        /// 初始化一个 <see cref="UniqueReservation"/> 类型的实例。
        /// </summary>
        /// <param name="values">用于唯一性检查的共享值集合。</param>
        /// <param name="value">待处理的值。</param>
        internal UniqueReservation(HashSet<string> values, string value)
        {
            Values = values;
            Value = value;
        }

        /// <summary>
        /// 获取持有已预留唯一值的集合。
        /// </summary>
        internal HashSet<string> Values { get; }
        /// <summary>
        /// 获取当前行预留的唯一值。
        /// </summary>
        internal string Value { get; }
    }

    /// <summary>
    /// 工作簿数据行数量预算。
    /// </summary>
    private sealed class RowBudget
    {
        /// <summary>
        /// 允许读取的数据行数上限；null 表示不限制。
        /// </summary>
        private readonly int? _maximum;
        /// <summary>
        /// 已消耗的数据行预算数量。
        /// </summary>
        private int _count;

        /// <summary>
        /// 初始化一个 <see cref="RowBudget"/> 类型的实例。
        /// </summary>
        /// <param name="maximum">允许消耗的数量上限；null 表示不限制。</param>
        internal RowBudget(int? maximum) => _maximum = maximum;
        /// <summary>
        /// 获取或设置是否已触发资源超限。
        /// </summary>
        internal bool Exceeded { get; private set; }

        /// <summary>
        /// 尝试消耗当前数据的资源预算。
        /// </summary>
        /// <returns>预算允许并完成计数时为 true；超限时为 false。</returns>
        internal bool TryConsume()
        {
            if (_maximum.HasValue && _count >= _maximum.Value)
            {
                Exceeded = true;
                return false;
            }
            _count++;
            return true;
        }
    }

    /// <summary>
    /// 工作表列数和工作簿单元格数量预算。
    /// </summary>
    private sealed class CellBudget
    {
        /// <summary>
        /// 每张工作表允许的列数上限；null 表示不限制。
        /// </summary>
        private readonly int? _maximumColumns;
        /// <summary>
        /// 工作簿允许的单元格总数上限；null 表示不限制。
        /// </summary>
        private readonly long? _maximumCells;
        /// <summary>
        /// 已累计消耗的单元格预算数量。
        /// </summary>
        private long _cells;

        /// <summary>
        /// 初始化一个 <see cref="CellBudget"/> 类型的实例。
        /// </summary>
        /// <param name="limits">本次导入的资源限制。</param>
        internal CellBudget(ExcelResourceLimits limits)
        {
            _maximumColumns = limits?.MaxColumnsPerSheet;
            _maximumCells = limits?.MaxCells;
        }

        /// <summary>
        /// 获取或设置是否已触发资源超限。
        /// </summary>
        internal bool Exceeded { get; private set; }

        /// <summary>
        /// 尝试消耗当前数据的资源预算。
        /// </summary>
        /// <param name="columns">当前行的物理列数。</param>
        /// <param name="rowNumber">从 1 开始的物理行号。</param>
        /// <param name="sheetName">用于错误定位的工作表名称。</param>
        /// <param name="error">资源超限时产生的错误；成功消耗预算时为 null。</param>
        /// <returns>预算允许并完成计数时为 true；超限时为 false。</returns>
        internal bool TryConsume(int columns, int rowNumber, string sheetName,
            out ExcelImportError error)
        {
            if (_maximumColumns.HasValue && columns > _maximumColumns.Value)
            {
                Exceeded = true;
                error = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    $"工作表列数超过资源限制: {_maximumColumns.Value}", sheetName, rowNumber, 0, null);
                return false;
            }
            if (_maximumCells.HasValue && columns > _maximumCells.Value - _cells)
            {
                Exceeded = true;
                error = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    $"Workbook 物理单元格数量超过资源限制: {_maximumCells.Value}", sheetName,
                    rowNumber, 0, null);
                return false;
            }
            _cells += columns;
            error = null;
            return true;
        }
    }

    /// <summary>
    /// 受错误数量上限约束的导入错误收集器。
    /// </summary>
    private sealed class ErrorCollector
    {
        /// <summary>
        /// 初始化一个 <see cref="ErrorCollector"/> 类型的实例。
        /// </summary>
        /// <param name="maxErrors">允许保留的错误数量上限；null 表示不限制。</param>
        internal ErrorCollector(int? maxErrors) => MaxErrors = maxErrors;
        /// <summary>
        /// 获取允许保留的错误数量上限。
        /// </summary>
        internal int? MaxErrors { get; }
        /// <summary>
        /// 获取已收集的导入错误。
        /// </summary>
        internal List<ExcelImportError> Items { get; } = new List<ExcelImportError>();
        /// <summary>
        /// 获取或设置是否因数量上限截断错误记录。
        /// </summary>
        internal bool Truncated { get; private set; }
        /// <summary>
        /// 获取已收集错误中是否包含资源超限错误。
        /// </summary>
        internal bool HasResourceLimit => Items.Any(item => item.Code == ExcelImportErrorCode.ResourceLimit);

        /// <summary>
        /// 在错误数量上限内收集导入错误。
        /// </summary>
        /// <param name="error">待收集的导入错误。</param>
        internal void Add(ExcelImportError error)
        {
            if (MaxErrors.HasValue && Items.Count >= MaxErrors.Value)
            {
                Truncated = true;
                return;
            }
            Items.Add(error);
        }
    }

    /// <summary>
    /// 用于区分调用方批次回调失败的异常。
    /// </summary>
    private sealed class BatchCallbackException : Exception
    {
        /// <summary>
        /// 初始化一个 <see cref="BatchCallbackException"/> 类型的实例。
        /// </summary>
        /// <param name="innerException">原始异常；没有原始异常时为 null。</param>
        internal BatchCallbackException(Exception innerException)
            : base("分批回调执行失败。", innerException)
        {
        }
    }
}
