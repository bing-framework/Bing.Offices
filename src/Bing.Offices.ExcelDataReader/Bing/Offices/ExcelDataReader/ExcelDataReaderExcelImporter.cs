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
    IExcelWorkbookBatchImporter,
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
    /// 负责失败工作簿目标路径原子提交的服务。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;

    /// <summary>
    /// 注册读取旧版工作簿所需的代码页编码提供程序。
    /// </summary>
    static ExcelDataReaderExcelImporter()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="validationRules">默认映射计划使用的校验规则集合。</param>
    /// <param name="valueConverters">默认映射计划使用的值转换器集合。</param>
    /// <param name="namedValidationRules">默认映射计划使用的命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂；为 null 时使用默认工厂。</param>
    public ExcelDataReaderExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory, null, null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter" /> 类型的实例。
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
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory,
            stagedPathFactory, null)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="validationRules">默认映射计划使用的校验规则集合。</param>
    /// <param name="valueConverters">默认映射计划使用的值转换器集合。</param>
    /// <param name="namedValidationRules">默认映射计划使用的命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂；为 null 时使用默认工厂。</param>
    /// <param name="stagedPathFactory">暂存文件路径工厂；为 null 时使用默认临时路径。</param>
    /// <param name="fileExportCommitter">失败工作簿路径输出使用的提交服务；为 null 时使用默认实现。</param>
    internal ExcelDataReaderExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory, Func<string> stagedPathFactory,
        IFileExportCommitter fileExportCommitter)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault().ToArray();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault(
            _valueConverters, _validationRules, _namedValidationRules);
        _stagedPathFactory = stagedPathFactory ?? CreateDefaultStagedPath;
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
    }

    /// <summary>
    /// 初始化一个 <see cref="ExcelDataReaderExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="stagedPathFactory">暂存文件路径工厂；为 null 时使用默认临时路径。</param>
    internal ExcelDataReaderExcelImporter(Func<string> stagedPathFactory)
        : this(null, null, null, null, stagedPathFactory, null)
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
    public ExcelProviderFeatures Features => ExcelProviderFeatures.FormulaCachedValues
        | ExcelProviderFeatures.NamedListValidation
        | ExcelProviderFeatures.MultiSheetBatchImport;

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
        "只读；动态列和跨 Sheet 关系仅用于 Workbook 分批导入，Entity、图片和完整导入的原生 Workbook 校验需使用其他 Provider。",
        "公式按读取引擎提供的缓存值处理，不执行公式计算。",
        "XLSX 的 MaxCells/MaxColumnsPerSheet 使用 ZIP 物理 Cell 预检；XLS/XLSB 请求这两项限制时明确返回 UnsupportedFeature。",
        "单 Sheet 与 Workbook 分批接口支持固定列、动态列和跨 Sheet 关系，已交付批次不回滚；XLSX 完整 Workbook 和分批入口支持 ErrorRowsOnly 与 AnnotatedOriginal，后者保留源部件并追加错误批注，XLS/XLSB 仍返回结构化 UnsupportedFeature。"
    };

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedRequest(request, allowFailureWorkbook: true);
        var path = StageInput(source, request.ResourceLimits, cancellationToken);
        try
        {
            using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException exception)
            {
                var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
                errors.Add(ToResourceError(exception));
                var limited = ResourceLimitedResult(new TWorkbook(), errors);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, limited.Errors,
                    capture, cancellationToken, _fileExportCommitter, path);
                return limited;
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var result = ImportReader(input, request, capture, cancellationToken);
            ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, result.Errors,
                capture, cancellationToken, _fileExportCommitter, path);
            return result;
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
        ValidateSupportedRequest(request, allowFailureWorkbook: true);
        var path = await StageInputAsync(source, request.ResourceLimits, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            await using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException exception)
            {
                var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
                errors.Add(ToResourceError(exception));
                var limited = ResourceLimitedResult(new TWorkbook(), errors);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions,
                    limited.Errors, capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
                return limited;
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var result = ImportReader(input, request, capture, cancellationToken);
            await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions, result.Errors,
                capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
            return result;
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
        ValidateBatchPlan(CreatePlan(request), typeof(TItem));
        var path = StageInput(source, request.ResourceLimits, cancellationToken);
        try
        {
            using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var resourceError = ToResourceError(preflightException);
                capture.CaptureErrors(new[] { resourceError });
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1, 1,
                    Array.Empty<TItem>(), new[] { resourceError }));
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                capture.CaptureErrors(new[] { budgetError.Error });
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1, budgetError.Error.RowIndex,
                    Array.Empty<TItem>(), new[] { budgetError.Error }));
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var workbookValidation = CreateWorkbookValidation(input, request.ValidationMode,
                cancellationToken);
            using var reader = CreateReader(input);
            var result = ReadBatches(reader, request, workbookValidation, onBatch, capture,
                cancellationToken);
            ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                capture, cancellationToken, _fileExportCommitter, path);
            return result;
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
        ValidateBatchPlan(CreatePlan(request), typeof(TItem));
        var path = await StageInputAsync(source, request.ResourceLimits, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            await using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var resourceError = ToResourceError(preflightException);
                capture.CaptureErrors(new[] { resourceError });
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1, 1,
                    Array.Empty<TItem>(), new[] { resourceError }), cancellationToken).ConfigureAwait(false);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions,
                    capture.Errors, capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                capture.CaptureErrors(new[] { budgetError.Error });
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1, budgetError.Error.RowIndex,
                    Array.Empty<TItem>(), new[] { budgetError.Error }), cancellationToken).ConfigureAwait(false);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions,
                    capture.Errors, capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
                return new ExcelBatchImportSummary(0, 0, 0, 1, false, true);
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var workbookValidation = CreateWorkbookValidation(input, request.ValidationMode,
                cancellationToken);
            using var reader = CreateReader(input);
            var result = await ReadBatchesAsync(reader, request, workbookValidation, onBatch, capture,
                cancellationToken).ConfigureAwait(false);
            await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions,
                capture.Errors, capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
            return result;
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

    /// <inheritdoc />
    public ExcelWorkbookBatchImportSummary ImportWorkbookBatches<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, Action<ExcelWorkbookImportBatch> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedRequest(request, allowWorkbookValidation: true, allowDynamicColumns: true,
            allowWorkbookRelations: true, allowFailureWorkbook: true);
        var plans = ValidateSupportedWorkbookBatchRequest(request);
        var path = StageInput(source, request.ResourceLimits, cancellationToken);
        try
        {
            using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var preflightBudgetError = new WorkbookResourceBudgetError(ToResourceError(preflightException), -1);
                capture.CaptureErrors(new[] { preflightBudgetError.Error });
                var limited = DeliverWorkbookResourceError(request, onBatch, preflightBudgetError);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path);
                return limited;
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                capture.CaptureErrors(new[] { budgetError.Error });
                var limited = DeliverWorkbookResourceError(request, onBatch, budgetError);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path);
                return limited;
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var workbookValidation = CreateWorkbookValidation(input, request.ValidationMode,
                cancellationToken);
            input.Position = 0;
            using var reader = CreateReader(input);
            var result = ReadWorkbookBatches(reader, request, plans, workbookValidation, onBatch, capture,
                cancellationToken);
            ExcelDataReaderFailureWorkbookWriter.Write(request.FailureOptions, capture.Errors,
                capture, cancellationToken, _fileExportCommitter, path);
            return result;
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
            throw new BingOfficesImportException("ExcelDataReader Workbook 分批导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelWorkbookBatchImportSummary> ImportWorkbookBatchesAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateArguments(source, request);
        cancellationToken.ThrowIfCancellationRequested();
        ValidateSupportedRequest(request, allowWorkbookValidation: true, allowDynamicColumns: true,
            allowWorkbookRelations: true, allowFailureWorkbook: true);
        var plans = ValidateSupportedWorkbookBatchRequest(request);
        var path = await StageInputAsync(source, request.ResourceLimits, cancellationToken)
            .ConfigureAwait(false);
        try
        {
            await using var input = OpenStagedInput(path);
            ValidateFailureWorkbookSource(request.FailureOptions, input);
            var capture = new ExcelDataReaderFailureWorkbookCapture(request.FailureOptions);
            try
            {
                ValidateProviderPreflight(input, request.ResourceLimits, cancellationToken);
            }
            catch (BingOfficesResourceLimitException preflightException)
            {
                var preflightBudgetError = new WorkbookResourceBudgetError(ToResourceError(preflightException), -1);
                capture.CaptureErrors(new[] { preflightBudgetError.Error });
                var limited = await DeliverWorkbookResourceErrorAsync(request, onBatch, preflightBudgetError,
                    cancellationToken).ConfigureAwait(false);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
                return limited;
            }
            var budgetError = ScanWorkbookResourceBudgets(input, request.ResourceLimits, cancellationToken);
            if (budgetError != null)
            {
                capture.CaptureErrors(new[] { budgetError.Error });
                var limited = await DeliverWorkbookResourceErrorAsync(request, onBatch, budgetError,
                    cancellationToken).ConfigureAwait(false);
                CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
                await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions, capture.Errors,
                    capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
                return limited;
            }
            CaptureFailureWorkbookMetadata(input, request.FailureOptions, capture, cancellationToken);
            var workbookValidation = CreateWorkbookValidation(input, request.ValidationMode,
                cancellationToken);
            input.Position = 0;
            using var reader = CreateReader(input);
            var result = await ReadWorkbookBatchesAsync(reader, request, plans, workbookValidation, onBatch,
                capture, cancellationToken).ConfigureAwait(false);
            await ExcelDataReaderFailureWorkbookWriter.WriteAsync(request.FailureOptions, capture.Errors,
                capture, cancellationToken, _fileExportCommitter, path).ConfigureAwait(false);
            return result;
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
            throw new BingOfficesImportException("ExcelDataReader 异步 Workbook 分批导入失败。", exception, Provider,
                BingOfficesStage.Read);
        }
        finally
        {
            DeleteStagedInput(path);
        }
    }

    /// <summary>
    /// 按物理工作表顺序读取工作簿，并交付非泛型批次上下文。
    /// </summary>
    /// <remarks>
    /// 通过同步批次回调交付结果。
    /// </remarks>
    private ExcelWorkbookBatchImportSummary ReadWorkbookBatches<TWorkbook>(DataReader reader,
        ExcelWorkbookImportRequest<TWorkbook> request,
        IReadOnlyDictionary<ExcelSheetImportRequest, IExcelMappingPlan> plans,
        ExcelXlsxWorkbookValidation workbookValidation,
        Action<ExcelWorkbookImportBatch> onBatch,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var summaries = new List<ExcelBatchImportSheetSummary>();
        var summaryIndexes = new Dictionary<ExcelSheetImportRequest, int>();
        var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        var relationCoordinator = request.Relations.Count == 0
            ? null
            : new ExcelDataReaderBatchRelationCoordinator<TWorkbook>(request);
        var rowBudget = new RowBudget(request.ResourceLimits?.MaxRows);
        // XLSX 的物理 Cell/列限制已由 ZIP 预检精确执行；XLS/XLSB 在请求这些限制时已提前拒绝。
        var cellBudget = new CellBudget(null);
        var matched = new HashSet<ExcelSheetImportRequest>();
        var batchNumber = new BatchNumberState();
        var physicalSheetIndex = 0;
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.Name == null)
                break;
            var selected = request.Sheets.Where(sheet => IsSelected(sheet, reader.Name, physicalSheetIndex,
                request.SheetNameComparison)).ToArray();
            if (selected.Length > 1)
                throw new BingOfficesConfigurationException(
                    $"多个 Sheet selector 指向同一物理 Sheet: {string.Join(", ", selected.Select(DescribeSelector))}; "
                    + $"实际 Sheet=#{physicalSheetIndex} {reader.Name}", stage: BingOfficesStage.Plan);
            foreach (var sheetRequest in selected)
            {
                Action<ExcelWorkbookImportBatch> sheetCallback = onBatch;
                if (relationCoordinator != null)
                {
                    sheetCallback = batch =>
                    {
                        relationCoordinator.AcceptBatch(sheetRequest, batch);
                        onBatch(batch);
                    };
                }
                var summary = ReadWorkbookSheetBatches(reader, sheetRequest, plans[sheetRequest], request, errors, rowBudget,
                    cellBudget, workbookValidation, sheetCallback, batchNumber, failureCapture, cancellationToken);
                summaryIndexes[sheetRequest] = summaries.Count;
                summaries.Add(new ExcelBatchImportSheetSummary(reader.Name, sheetRequest.ItemType, summary));
                matched.Add(sheetRequest);
                if (rowBudget.Exceeded || cellBudget.Exceeded || errors.Truncated || errors.HasResourceLimit)
                    return new ExcelWorkbookBatchImportSummary(summaries);
            }
            physicalSheetIndex++;
        } while (reader.NextResult());

        foreach (var sheetRequest in request.Sheets.Where(sheet => !matched.Contains(sheet)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                $"未找到工作表: {sheetRequest.Name}", sheetRequest.Name, 1, 0, null);
            var errorCountBefore = errors.Items.Count;
            errors.Add(error);
            failureCapture?.CaptureErrors(new[] { error });
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(sheetRequest.Name, sheetRequest.ItemType,
                batchNumber.Next(), 1, Array.Empty<object>(), new[] { error }));
            summaryIndexes[sheetRequest] = summaries.Count;
            summaries.Add(new ExcelBatchImportSheetSummary(sheetRequest.Name, sheetRequest.ItemType,
                new ExcelBatchImportSummary(0, 0, 1, errors.Items.Count - errorCountBefore,
                    errors.Truncated, false)));
        }
        AppendRelationBindingResults(relationCoordinator, request, summaries, summaryIndexes, errors,
            onBatch, batchNumber, failureCapture, cancellationToken);
        return new ExcelWorkbookBatchImportSummary(summaries);
    }

    /// <summary>
    /// 按物理工作表顺序读取工作簿，并交付非泛型批次上下文。
    /// </summary>
    /// <remarks>
    /// 等待每个异步批次回调完成。
    /// </remarks>
    private async Task<ExcelWorkbookBatchImportSummary> ReadWorkbookBatchesAsync<TWorkbook>(DataReader reader,
        ExcelWorkbookImportRequest<TWorkbook> request,
        IReadOnlyDictionary<ExcelSheetImportRequest, IExcelMappingPlan> plans,
        ExcelXlsxWorkbookValidation workbookValidation,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var summaries = new List<ExcelBatchImportSheetSummary>();
        var summaryIndexes = new Dictionary<ExcelSheetImportRequest, int>();
        var errors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        var relationCoordinator = request.Relations.Count == 0
            ? null
            : new ExcelDataReaderBatchRelationCoordinator<TWorkbook>(request);
        var rowBudget = new RowBudget(request.ResourceLimits?.MaxRows);
        // XLSX 的物理 Cell/列限制已由 ZIP 预检精确执行；XLS/XLSB 在请求这些限制时已提前拒绝。
        var cellBudget = new CellBudget(null);
        var matched = new HashSet<ExcelSheetImportRequest>();
        var batchNumber = new BatchNumberState();
        var physicalSheetIndex = 0;
        do
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.Name == null)
                break;
            var selected = request.Sheets.Where(sheet => IsSelected(sheet, reader.Name, physicalSheetIndex,
                request.SheetNameComparison)).ToArray();
            if (selected.Length > 1)
                throw new BingOfficesConfigurationException(
                    $"多个 Sheet selector 指向同一物理 Sheet: {string.Join(", ", selected.Select(DescribeSelector))}; "
                    + $"实际 Sheet=#{physicalSheetIndex} {reader.Name}", stage: BingOfficesStage.Plan);
            foreach (var sheetRequest in selected)
            {
                Func<ExcelWorkbookImportBatch, CancellationToken, Task> sheetCallback = onBatch;
                if (relationCoordinator != null)
                {
                    sheetCallback = (batch, callbackToken) =>
                    {
                        relationCoordinator.AcceptBatch(sheetRequest, batch);
                        return onBatch(batch, callbackToken);
                    };
                }
                var summary = await ReadWorkbookSheetBatchesAsync(reader, sheetRequest, plans[sheetRequest], request, errors,
                    rowBudget, cellBudget, workbookValidation, sheetCallback, batchNumber, failureCapture,
                    cancellationToken).ConfigureAwait(false);
                summaryIndexes[sheetRequest] = summaries.Count;
                summaries.Add(new ExcelBatchImportSheetSummary(reader.Name, sheetRequest.ItemType, summary));
                matched.Add(sheetRequest);
                if (rowBudget.Exceeded || cellBudget.Exceeded || errors.Truncated || errors.HasResourceLimit)
                    return new ExcelWorkbookBatchImportSummary(summaries);
            }
            physicalSheetIndex++;
        } while (reader.NextResult());

        foreach (var sheetRequest in request.Sheets.Where(sheet => !matched.Contains(sheet)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                $"未找到工作表: {sheetRequest.Name}", sheetRequest.Name, 1, 0, null);
            var errorCountBefore = errors.Items.Count;
            errors.Add(error);
            failureCapture?.CaptureErrors(new[] { error });
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(sheetRequest.Name,
                sheetRequest.ItemType, batchNumber.Next(), 1, Array.Empty<object>(), new[] { error }),
                cancellationToken).ConfigureAwait(false);
            summaryIndexes[sheetRequest] = summaries.Count;
            summaries.Add(new ExcelBatchImportSheetSummary(sheetRequest.Name, sheetRequest.ItemType,
                new ExcelBatchImportSummary(0, 0, 1, errors.Items.Count - errorCountBefore,
                    errors.Truncated, false)));
        }
        await AppendRelationBindingResultsAsync(relationCoordinator, request, summaries, summaryIndexes, errors,
            onBatch, batchNumber, failureCapture, cancellationToken).ConfigureAwait(false);
        return new ExcelWorkbookBatchImportSummary(summaries);
    }

    /// <summary>
    /// 追加工作簿分批关系绑定结果。
    /// </summary>
    /// <remarks>
    /// 通过同步批次回调交付结果。
    /// </remarks>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="coordinator">关系协调器；未配置关系时为 null。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="summaries">工作表摘要集合。</param>
    /// <param name="summaryIndexes">Sheet 请求到摘要索引的映射。</param>
    /// <param name="errors">全局错误收集器。</param>
    /// <param name="onBatch">批次回调。</param>
    /// <param name="batchNumber">全局批次号状态。</param>
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static void AppendRelationBindingResults<TWorkbook>(
        ExcelDataReaderBatchRelationCoordinator<TWorkbook> coordinator,
        ExcelWorkbookImportRequest<TWorkbook> request,
        List<ExcelBatchImportSheetSummary> summaries,
        IReadOnlyDictionary<ExcelSheetImportRequest, int> summaryIndexes,
        ErrorCollector errors, Action<ExcelWorkbookImportBatch> onBatch,
        BatchNumberState batchNumber, ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        if (coordinator == null)
            return;
        var remainingErrors = request.ResourceLimits?.MaxErrors is int maxErrors
            ? Math.Max(maxErrors - errors.Items.Count, 0)
            : (int?)null;
        var result = coordinator.Complete(remainingErrors, cancellationToken);
        foreach (var group in result.Groups)
        {
            var acceptedCount = 0;
            foreach (var error in group.Errors)
            {
                var before = errors.Items.Count;
                errors.Add(error);
                if (errors.Items.Count > before)
                    acceptedCount++;
            }
            if (acceptedCount == 0)
            {
                if (errors.Truncated)
                    UpdateRelationSummary(summaries, summaryIndexes, group.Sheet, 0, true);
                continue;
            }
            var acceptedErrors = group.Errors.Take(acceptedCount).ToArray();
            failureCapture?.CaptureErrors(acceptedErrors);
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(group.Sheet.Name, group.Sheet.ItemType,
                batchNumber.Next(), 0, Array.Empty<object>(), acceptedErrors));
            UpdateRelationSummary(summaries, summaryIndexes, group.Sheet, acceptedCount, errors.Truncated);
        }
        if (result.Truncated)
        {
            errors.MarkTruncated();
            if (result.TruncatedSheet != null)
                UpdateRelationSummary(summaries, summaryIndexes, result.TruncatedSheet, 0, true);
        }
    }

    /// <summary>
    /// 追加工作簿分批关系绑定结果。
    /// </summary>
    /// <remarks>
    /// 等待异步批次回调完成。
    /// </remarks>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="coordinator">关系协调器；未配置关系时为 null。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="summaries">工作表摘要集合。</param>
    /// <param name="summaryIndexes">工作表请求到摘要索引的映射。</param>
    /// <param name="errors">全局错误收集器。</param>
    /// <param name="onBatch">异步批次回调。</param>
    /// <param name="batchNumber">全局批次号状态。</param>
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private static async Task AppendRelationBindingResultsAsync<TWorkbook>(
        ExcelDataReaderBatchRelationCoordinator<TWorkbook> coordinator,
        ExcelWorkbookImportRequest<TWorkbook> request,
        List<ExcelBatchImportSheetSummary> summaries,
        IReadOnlyDictionary<ExcelSheetImportRequest, int> summaryIndexes,
        ErrorCollector errors, Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch,
        BatchNumberState batchNumber, ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        if (coordinator == null)
            return;
        var remainingErrors = request.ResourceLimits?.MaxErrors is int maxErrors
            ? Math.Max(maxErrors - errors.Items.Count, 0)
            : (int?)null;
        var result = coordinator.Complete(remainingErrors, cancellationToken);
        foreach (var group in result.Groups)
        {
            var acceptedCount = 0;
            foreach (var error in group.Errors)
            {
                var before = errors.Items.Count;
                errors.Add(error);
                if (errors.Items.Count > before)
                    acceptedCount++;
            }
            if (acceptedCount == 0)
            {
                if (errors.Truncated)
                    UpdateRelationSummary(summaries, summaryIndexes, group.Sheet, 0, true);
                continue;
            }
            var acceptedErrors = group.Errors.Take(acceptedCount).ToArray();
            failureCapture?.CaptureErrors(acceptedErrors);
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(group.Sheet.Name,
                group.Sheet.ItemType, batchNumber.Next(), 0, Array.Empty<object>(), acceptedErrors),
                cancellationToken).ConfigureAwait(false);
            UpdateRelationSummary(summaries, summaryIndexes, group.Sheet, acceptedCount, errors.Truncated);
        }
        if (result.Truncated)
        {
            errors.MarkTruncated();
            if (result.TruncatedSheet != null)
                UpdateRelationSummary(summaries, summaryIndexes, result.TruncatedSheet, 0, true);
        }
    }

    /// <summary>
    /// 将关系级错误数量合并到对应工作表摘要。
    /// </summary>
    /// <param name="summaries">工作表摘要集合。</param>
    /// <param name="summaryIndexes">Sheet 请求到摘要索引的映射。</param>
    /// <param name="sheet">关系错误所属的工作表请求。</param>
    /// <param name="errorCount">新增错误数量。</param>
    /// <param name="truncated">是否达到错误预算。</param>
    private static void UpdateRelationSummary(List<ExcelBatchImportSheetSummary> summaries,
        IReadOnlyDictionary<ExcelSheetImportRequest, int> summaryIndexes,
        ExcelSheetImportRequest sheet, int errorCount, bool truncated)
    {
        if (!summaryIndexes.TryGetValue(sheet, out var index))
            return;
        var current = summaries[index];
        var summary = current.Summary;
        summaries[index] = new ExcelBatchImportSheetSummary(current.SheetName, current.ItemType,
            new ExcelBatchImportSummary(summary.RowsRead, summary.SucceededRows,
                summary.FailedRows, summary.ErrorCount + errorCount,
                summary.ErrorsTruncated || truncated, summary.ResourceLimitExceeded));
    }

    /// <summary>
    /// 读取单个工作表并交付批次。
    /// </summary>
    /// <remarks>
    /// 通过同步批次回调交付结果。
    /// </remarks>
    private ExcelBatchImportSummary ReadWorkbookSheetBatches<TWorkbook>(DataReader reader,
        ExcelSheetImportRequest sheetRequest, IExcelMappingPlan plan,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        ErrorCollector errors, RowBudget rowBudget, CellBudget cellBudget,
        ExcelXlsxWorkbookValidation workbookValidation,
        Action<ExcelWorkbookImportBatch> onBatch, BatchNumberState batchNumber,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var headers = ReadHeader(reader, sheetRequest, cancellationToken);
        failureCapture?.CaptureSheet(reader.Name, ReadCurrentRow(reader), sheetRequest.HeaderRowIndex);
        var sheetErrors = new List<ExcelImportError>();
        var sheetErrorStart = errors.Items.Count;
        if (!cellBudget.TryConsume(reader.FieldCount, sheetRequest.HeaderRowIndex + 1, reader.Name,
            out var headerResourceError))
        {
            errors.Add(headerResourceError);
            sheetErrors.Add(headerResourceError);
            failureCapture?.CaptureErrors(new[] { headerResourceError });
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(reader.Name, sheetRequest.ItemType,
                batchNumber.Next(), sheetRequest.HeaderRowIndex + 1, Array.Empty<object>(), sheetErrors.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, errors.Items.Count - sheetErrorStart,
                errors.Truncated, true);
        }
        var headerErrorCount = errors.Items.Count;
        var bindings = BindColumns(plan, headers, sheetRequest, reader.Name, errors);
        if (errors.Items.Count > headerErrorCount)
        {
            var headerErrors = errors.Items.Skip(headerErrorCount).ToArray();
            sheetErrors.AddRange(headerErrors);
            failureCapture?.CaptureErrors(headerErrors);
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(reader.Name, sheetRequest.ItemType,
                batchNumber.Next(), sheetRequest.HeaderRowIndex + 1, Array.Empty<object>(), sheetErrors.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, errors.Items.Count - sheetErrorStart,
                errors.Truncated, false);
        }

        var unique = CreateUniqueTracker(workbookRequest.ResourceLimits);
        var batch = new List<object>();
        var batchErrors = new List<ExcelImportError>();
        var batchFirstRow = 0;
        var rowsRead = 0;
        var succeeded = 0;
        var failed = 0;
        var zeroBasedRowIndex = sheetRequest.HeaderRowIndex;
        var physicalRow = sheetRequest.HeaderRowIndex + 1;
        ExcelImportError resourceError = null;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            zeroBasedRowIndex++;
            physicalRow = zeroBasedRowIndex + 1;
            if (!cellBudget.TryConsume(reader.FieldCount, physicalRow, reader.Name, out resourceError))
            {
                errors.Add(resourceError);
                sheetErrors.Add(resourceError);
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { resourceError });
                break;
            }
            if (zeroBasedRowIndex < sheetRequest.DataRowStartIndex)
                continue;
            if (!rowBudget.TryConsume())
            {
                resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    "Workbook 数据行数超过资源限制。", reader.Name, physicalRow, 0, null);
                errors.Add(resourceError);
                sheetErrors.Add(resourceError);
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { resourceError });
                break;
            }
            rowsRead++;
            batchFirstRow = batchFirstRow == 0 ? physicalRow : batchFirstRow;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeWorkbookRow<TWorkbook>(reader, sheetRequest, plan, bindings, reader.Name,
                physicalRow, unique, rowErrors, workbookRequest.ResourceLimits, workbookRequest.ValidationMode,
                workbookRequest.UnsupportedFeaturePolicy, workbookValidation, cancellationToken);
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
            if (rowErrors.Count > 0)
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader), rowErrors);
            foreach (var error in rowErrors)
            {
                errors.Add(error);
                sheetErrors.Add(error);
            }
            if (batch.Count >= sheetRequest.BatchSize || batchErrors.Count >= sheetRequest.BatchSize)
            {
                InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(reader.Name, sheetRequest.ItemType,
                    batchNumber.Next(), batchFirstRow, batch.ToArray(), batchErrors.ToArray()));
                batch.Clear();
                batchErrors.Clear();
                batchFirstRow = 0;
            }
            if (errors.Truncated || (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                && rowErrors.Count > 0))
                break;
        }
        if (batch.Count > 0 || batchErrors.Count > 0)
        {
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(reader.Name, sheetRequest.ItemType,
                batchNumber.Next(), batchFirstRow, batch.ToArray(), batchErrors.ToArray()));
        }
        if (resourceError != null)
        {
            InvokeWorkbookBatch(onBatch, new ExcelWorkbookImportBatch(reader.Name, sheetRequest.ItemType,
                batchNumber.Next(), resourceError.RowIndex, Array.Empty<object>(), new[] { resourceError }));
        }
        return new ExcelBatchImportSummary(rowsRead, succeeded, failed, errors.Items.Count - sheetErrorStart,
            errors.Truncated, resourceError != null || rowBudget.Exceeded || cellBudget.Exceeded);
    }

    /// <summary>
    /// 读取单个工作表并交付批次。
    /// </summary>
    /// <remarks>
    /// 等待异步批次回调完成。
    /// </remarks>
    private async Task<ExcelBatchImportSummary> ReadWorkbookSheetBatchesAsync<TWorkbook>(DataReader reader,
        ExcelSheetImportRequest sheetRequest, IExcelMappingPlan plan,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        ErrorCollector errors, RowBudget rowBudget, CellBudget cellBudget,
        ExcelXlsxWorkbookValidation workbookValidation,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> onBatch, BatchNumberState batchNumber,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var headers = ReadHeader(reader, sheetRequest, cancellationToken);
        failureCapture?.CaptureSheet(reader.Name, ReadCurrentRow(reader), sheetRequest.HeaderRowIndex);
        var sheetErrors = new List<ExcelImportError>();
        var sheetErrorStart = errors.Items.Count;
        if (!cellBudget.TryConsume(reader.FieldCount, sheetRequest.HeaderRowIndex + 1, reader.Name,
            out var headerResourceError))
        {
            errors.Add(headerResourceError);
            sheetErrors.Add(headerResourceError);
            failureCapture?.CaptureErrors(new[] { headerResourceError });
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(reader.Name,
                sheetRequest.ItemType, batchNumber.Next(), sheetRequest.HeaderRowIndex + 1,
                Array.Empty<object>(), sheetErrors.ToArray()), cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, errors.Items.Count - sheetErrorStart,
                errors.Truncated, true);
        }
        var headerErrorCount = errors.Items.Count;
        var bindings = BindColumns(plan, headers, sheetRequest, reader.Name, errors);
        if (errors.Items.Count > headerErrorCount)
        {
            var headerErrors = errors.Items.Skip(headerErrorCount).ToArray();
            sheetErrors.AddRange(headerErrors);
            failureCapture?.CaptureErrors(headerErrors);
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(reader.Name,
                sheetRequest.ItemType, batchNumber.Next(), sheetRequest.HeaderRowIndex + 1,
                Array.Empty<object>(), sheetErrors.ToArray()), cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, errors.Items.Count - sheetErrorStart,
                errors.Truncated, false);
        }

        var unique = CreateUniqueTracker(workbookRequest.ResourceLimits);
        var batch = new List<object>();
        var batchErrors = new List<ExcelImportError>();
        var batchFirstRow = 0;
        var rowsRead = 0;
        var succeeded = 0;
        var failed = 0;
        var zeroBasedRowIndex = sheetRequest.HeaderRowIndex;
        var physicalRow = sheetRequest.HeaderRowIndex + 1;
        ExcelImportError resourceError = null;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            zeroBasedRowIndex++;
            physicalRow = zeroBasedRowIndex + 1;
            if (!cellBudget.TryConsume(reader.FieldCount, physicalRow, reader.Name, out resourceError))
            {
                errors.Add(resourceError);
                sheetErrors.Add(resourceError);
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { resourceError });
                break;
            }
            if (zeroBasedRowIndex < sheetRequest.DataRowStartIndex)
                continue;
            if (!rowBudget.TryConsume())
            {
                resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    "Workbook 数据行数超过资源限制。", reader.Name, physicalRow, 0, null);
                errors.Add(resourceError);
                sheetErrors.Add(resourceError);
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { resourceError });
                break;
            }
            rowsRead++;
            batchFirstRow = batchFirstRow == 0 ? physicalRow : batchFirstRow;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeWorkbookRow<TWorkbook>(reader, sheetRequest, plan, bindings, reader.Name,
                physicalRow, unique, rowErrors, workbookRequest.ResourceLimits, workbookRequest.ValidationMode,
                workbookRequest.UnsupportedFeaturePolicy, workbookValidation, cancellationToken);
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
            if (rowErrors.Count > 0)
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader), rowErrors);
            foreach (var error in rowErrors)
            {
                errors.Add(error);
                sheetErrors.Add(error);
            }
            if (batch.Count >= sheetRequest.BatchSize || batchErrors.Count >= sheetRequest.BatchSize)
            {
                await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(reader.Name,
                    sheetRequest.ItemType, batchNumber.Next(), batchFirstRow, batch.ToArray(), batchErrors.ToArray()),
                    cancellationToken).ConfigureAwait(false);
                batch.Clear();
                batchErrors.Clear();
                batchFirstRow = 0;
            }
            if (errors.Truncated || (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                && rowErrors.Count > 0))
                break;
        }
        if (batch.Count > 0 || batchErrors.Count > 0)
        {
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(reader.Name,
                sheetRequest.ItemType, batchNumber.Next(), batchFirstRow, batch.ToArray(), batchErrors.ToArray()),
                cancellationToken).ConfigureAwait(false);
        }
        if (resourceError != null)
        {
            await InvokeWorkbookBatchAsync(onBatch, new ExcelWorkbookImportBatch(reader.Name,
                sheetRequest.ItemType, batchNumber.Next(), resourceError.RowIndex,
                Array.Empty<object>(), new[] { resourceError }), cancellationToken).ConfigureAwait(false);
        }
        return new ExcelBatchImportSummary(rowsRead, succeeded, failed, errors.Items.Count - sheetErrorStart,
            errors.Truncated, resourceError != null || rowBudget.Exceeded || cellBudget.Exceeded);
    }

    /// <summary>
    /// 读取工作簿并汇总工作表导入结果。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿目标类型。</typeparam>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>工作簿数据、工作表结果和错误信息。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportReader<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var resourceError = ScanWorkbookResourceBudgets(source, request.ResourceLimits, cancellationToken);
        if (resourceError != null)
        {
            var limitedErrors = new ErrorCollector(request.ResourceLimits?.MaxErrors);
            limitedErrors.Add(resourceError.Error);
            failureCapture?.CaptureErrors(new[] { resourceError.Error });
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
                var error = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    "工作表数量超过资源限制。", reader.Name, 1, 0, null);
                errors.Add(error);
                failureCapture?.CaptureErrors(new[] { error });
                return ResourceLimitedResult(root, errors);
            }
            var selected = request.Sheets.Where(sheet => IsSelected(sheet, reader.Name, physicalSheetIndex,
                request.SheetNameComparison)).ToArray();
            if (selected.Length > 0)
            {
                foreach (var sheetRequest in selected)
                {
                    var sheetResult = ReadSheet(reader, sheetRequest, root, rowBudget, cellBudget, errors,
                        request, failureCapture, cancellationToken);
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
            var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                $"未找到工作表: {sheetRequest.Name}", sheetRequest.Name, 1, 0, null);
            errors.Add(error);
            failureCapture?.CaptureErrors(new[] { error });
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
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>当前工作表的成功行索引和错误信息。</returns>
    private ExcelSheetImportResult ReadSheet<TWorkbook>(DataReader reader, ExcelSheetImportRequest request,
        TWorkbook root, RowBudget rowBudget, CellBudget cellBudget, ErrorCollector errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var plan = CreatePlan(request);
        if (plan.DynamicColumns.Count > 0 || request.DynamicColumns.Count > 0)
            throw Unsupported("动态列");
        var headers = ReadHeader(reader, request, cancellationToken);
        failureCapture?.CaptureSheet(reader.Name, ReadCurrentRow(reader), request.HeaderRowIndex);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            reader.Name, out var headerResourceError))
        {
            errors.Add(headerResourceError);
            failureCapture?.CaptureErrors(new[] { headerResourceError });
            return new ExcelSheetImportResult(reader.Name, request.ItemType, Array.Empty<int>(),
                new[] { headerResourceError });
        }
        var bindings = BindColumns(plan, headers, request, reader.Name, errors);
        var items = new List<object>();
        var rows = new List<int>();
        var sheetErrors = new List<ExcelImportError>();
        var unique = CreateUniqueTracker(workbookRequest.ResourceLimits);
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
                failureCapture?.CaptureErrors(new[] { cellResourceError });
                break;
            }
            if (zeroBasedRowIndex < request.DataRowStartIndex)
                continue;
            if (!rowBudget.TryConsume())
                break;
            var rowErrors = new List<ExcelImportError>();
            var item = MaterializeWorkbookRow<TWorkbook>(reader, request, plan, bindings, reader.Name, physicalRow,
                unique, rowErrors, workbookRequest.ResourceLimits, workbookRequest.ValidationMode,
                workbookRequest.UnsupportedFeaturePolicy, null, cancellationToken);
            if (rowErrors.Count > 0)
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader), rowErrors);
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
            failureCapture?.CaptureErrors(new[] { resourceError });
            return new ExcelSheetImportResult(reader.Name, request.ItemType, Array.Empty<int>(),
                new[] { resourceError });
        }
        AddItems(root, request.Target(root), items);
        return new ExcelSheetImportResult(reader.Name, request.ItemType, rows, sheetErrors);
    }

    /// <summary>
    /// 读取工作表并交付导入批次。
    /// </summary>
    /// <remarks>
    /// 等待异步批次回调完成。
    /// </remarks>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="workbookValidation">XLSX 原生校验快照。</param>
    /// <param name="onBatch">接收已转换数据与错误的批次回调。</param>
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已读取、成功、失败和资源限制状态的统计结果。</returns>
    private async Task<ExcelBatchImportSummary> ReadBatchesAsync<TItem>(DataReader reader,
        ExcelBatchImportRequest<TItem> request,
        ExcelXlsxWorkbookValidation workbookValidation,
        Func<ExcelImportBatch<TItem>, CancellationToken, Task> onBatch,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TItem : class, new()
    {
        if (!MoveToSheet(reader, request.SheetName, cancellationToken))
            throw new BingOfficesImportException($"未找到工作表: {request.SheetName}", provider: Provider,
                stage: BingOfficesStage.Read);
        var plan = CreatePlan(request);
        ValidateBatchPlan(plan, typeof(TItem));
        var headers = ReadHeader(reader, request.HeaderRowIndex, request.DataRowStartIndex, request.Culture,
            request.HeaderWhitespace, cancellationToken);
        failureCapture?.CaptureSheet(reader.Name, ReadCurrentRow(reader), request.HeaderRowIndex);
        var collector = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        // XLSX 的物理 Cell/列限制已经由 ZIP 预检按 XML 节点精确执行；这里不再用 FieldCount 重复估算。
        var cellBudget = new CellBudget(null);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            request.SheetName, out var headerResourceError))
        {
            collector.Add(headerResourceError);
            failureCapture?.CaptureErrors(new[] { headerResourceError });
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()),
                cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, true);
        }
        var bindings = BindColumns(plan, headers, request.RequireExpectedHeaders, request.SheetName,
            request.HeaderComparison, collector);
        if (collector.Items.Count > 0)
        {
            failureCapture?.CaptureErrors(collector.Items);
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()),
                cancellationToken).ConfigureAwait(false);
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, false);
        }
        var unique = CreateUniqueTracker(request.ResourceLimits);
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
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { cellResourceError });
                if (batch.Count > 0 || batchErrors.Count > 0)
                    await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber++, batchFirstRow,
                        batch.ToArray(), batchErrors.ToArray()), cancellationToken).ConfigureAwait(false);
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber, physicalRow,
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
                physicalRow, unique, rowErrors, workbookValidation, ExcelUnsupportedFeaturePolicy.Fail,
                null, cancellationToken);
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
            if (rowErrors.Count > 0)
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader), rowErrors);
            if (batch.Count >= request.BatchSize || batchErrors.Count >= request.BatchSize)
            {
                await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber++, batchFirstRow,
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
            await InvokeBatchAsync(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber, batchFirstRow,
                batch.ToArray(), batchErrors.ToArray()), cancellationToken).ConfigureAwait(false);
        }
        if (budget.Exceeded)
        {
            var resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                "Workbook 数据行数超过资源限制。", request.SheetName, physicalRow, 0, null);
            collector.Add(resourceError);
            failureCapture?.CaptureErrors(new[] { resourceError });
        }
        return new ExcelBatchImportSummary(rowsRead, succeeded, failed, collector.Items.Count,
            collector.Truncated, budget.Exceeded || cellBudget.Exceeded);
    }

    /// <summary>
    /// 读取工作表并交付导入批次。
    /// </summary>
    /// <remarks>
    /// 通过同步批次回调交付结果。
    /// </remarks>
    /// <typeparam name="TItem">具有公共无参构造函数的导入实体类型。</typeparam>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="workbookValidation">XLSX 原生校验快照。</param>
    /// <param name="onBatch">接收已转换数据与错误的批次回调。</param>
    /// <param name="failureCapture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已读取、成功、失败和资源限制状态的统计结果。</returns>
    private ExcelBatchImportSummary ReadBatches<TItem>(DataReader reader,
        ExcelBatchImportRequest<TItem> request, ExcelXlsxWorkbookValidation workbookValidation,
        Action<ExcelImportBatch<TItem>> onBatch,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        CancellationToken cancellationToken)
        where TItem : class, new()
    {
        if (!MoveToSheet(reader, request.SheetName, cancellationToken))
            throw new BingOfficesImportException($"未找到工作表: {request.SheetName}", provider: Provider,
                stage: BingOfficesStage.Read);
        var plan = CreatePlan(request);
        ValidateBatchPlan(plan, typeof(TItem));
        var headers = ReadHeader(reader, request.HeaderRowIndex, request.DataRowStartIndex, request.Culture,
            request.HeaderWhitespace, cancellationToken);
        failureCapture?.CaptureSheet(reader.Name, ReadCurrentRow(reader), request.HeaderRowIndex);
        var collector = new ErrorCollector(request.ResourceLimits?.MaxErrors);
        var cellBudget = new CellBudget(request.ResourceLimits);
        if (!cellBudget.TryConsume(reader.FieldCount, request.HeaderRowIndex + 1,
            request.SheetName, out var headerResourceError))
        {
            collector.Add(headerResourceError);
            failureCapture?.CaptureErrors(new[] { headerResourceError });
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, true);
        }
        var bindings = BindColumns(plan, headers, request.RequireExpectedHeaders, request.SheetName,
            request.HeaderComparison, collector);
        if (collector.Items.Count > 0)
        {
            failureCapture?.CaptureErrors(collector.Items);
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, 1,
                request.HeaderRowIndex + 1, Array.Empty<TItem>(), collector.Items.ToArray()));
            return new ExcelBatchImportSummary(0, 0, 0, collector.Items.Count,
                collector.Truncated, false);
        }
        var unique = CreateUniqueTracker(request.ResourceLimits);
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
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader),
                    new[] { cellResourceError });
                if (batch.Count > 0 || batchErrors.Count > 0)
                    InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber++, batchFirstRow,
                        batch.ToArray(), batchErrors.ToArray()));
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber, physicalRow,
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
                physicalRow, unique, rowErrors, workbookValidation, ExcelUnsupportedFeaturePolicy.Fail,
                null, cancellationToken);
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
            if (rowErrors.Count > 0)
                failureCapture?.CaptureErrorRow(reader.Name, physicalRow, ReadCurrentRow(reader), rowErrors);
            foreach (var error in rowErrors)
                collector.Add(error);
            if (batch.Count >= request.BatchSize || batchErrors.Count >= request.BatchSize)
            {
                InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber++, batchFirstRow,
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
            InvokeBatch(onBatch, new ExcelImportBatch<TItem>(request.SheetName, batchNumber, batchFirstRow,
                batch.ToArray(), batchErrors.ToArray()));
        if (budget.Exceeded)
        {
            var resourceError = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                "Workbook 数据行数超过资源限制。", request.SheetName, physicalRow, 0, null);
            collector.Add(resourceError);
            failureCapture?.CaptureErrors(new[] { resourceError });
        }
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
    /// 快照读取器当前所在行的原始单元格值。
    /// </summary>
    /// <param name="reader">当前工作簿读取器。</param>
    /// <returns>按物理列顺序排列的原始单元格值。</returns>
    private static IReadOnlyList<object> ReadCurrentRow(DataReader reader)
    {
        var values = new object[reader.FieldCount];
        for (var index = 0; index < values.Length; index++)
            values[index] = reader.GetValue(index);
        return values;
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
        => BindColumns(plan, headers, request.RequireExpectedHeaders, sheetName, request.HeaderComparison, errors,
            request.FailOnUnknownDynamicColumns);

    /// <summary>
    /// 将映射列绑定到匹配的表头位置。
    /// </summary>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="headers">按物理列顺序排列的表头文本。</param>
    /// <param name="requireExpectedHeaders">缺少期望表头时是否记录错误。</param>
    /// <param name="sheetName">用于错误定位的工作表名称。</param>
    /// <param name="comparison">名称或文本的比较策略。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="failOnUnknownDynamicColumns">未知动态表头是否记录结构化错误。</param>
    /// <returns>成功匹配表头的列绑定集合。</returns>
    private static IReadOnlyList<ColumnBinding> BindColumns(IExcelMappingPlan plan,
        IReadOnlyList<string> headers, bool requireExpectedHeaders, string sheetName,
        ExcelNameComparison comparison, ErrorCollector errors, bool failOnUnknownDynamicColumns = false)
    {
        var result = new List<ColumnBinding>();
        var dynamicProperties = plan.Columns.Where(column => !column.Ignored && column.IsDynamicColumn).ToArray();
        if (dynamicProperties.Length > 1)
            throw new BingOfficesConfigurationException("一个导入实体只能声明一个动态列属性。",
                stage: BingOfficesStage.Plan);
        var dynamicProperty = dynamicProperties.SingleOrDefault();
        var dynamicDefinitions = plan.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>();
        var fixedColumns = plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn).ToArray();
        var matchedFixedColumns = new HashSet<IExcelMappingColumn>();
        for (var index = 0; index < headers.Count; index++)
        {
            var header = headers[index];
            var fixedColumn = fixedColumns.FirstOrDefault(column =>
                FindHeader(new[] { header }, column.Title, column.Aliases, comparison) == 0);
            if (fixedColumn != null)
            {
                matchedFixedColumns.Add(fixedColumn);
                result.Add(new ColumnBinding(fixedColumn, index, header));
                continue;
            }
            if (dynamicProperty == null)
                continue;
            var definition = FindDynamicDefinition(header, dynamicDefinitions, comparison);
            if (dynamicDefinitions.Count > 0 && definition == null)
            {
                if (failOnUnknownDynamicColumns)
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                        $"导入包含未知动态列: {header}", sheetName, 1, index + 1,
                        dynamicProperty.Name, header, header));
                continue;
            }
            result.Add(new ColumnBinding(dynamicProperty, index, header, definition));
        }
        if (requireExpectedHeaders)
        {
            foreach (var column in fixedColumns.Where(column => !matchedFixedColumns.Contains(column)))
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少表头: {column.Title}", sheetName, 1, 0, column.Name,
                    column.Name, column.Title));
        }
        return result;
    }

    /// <summary>
    /// 按标题或别名查找动态列定义。
    /// </summary>
    /// <param name="header">待匹配的表头文本。</param>
    /// <param name="definitions">已编译的动态列定义。</param>
    /// <param name="comparison">表头名称比较策略。</param>
    /// <returns>匹配的动态列定义；未匹配时返回 <see langword="null" />。</returns>
    private static IExcelDynamicMappingColumn FindDynamicDefinition(string header,
        IReadOnlyList<IExcelDynamicMappingColumn> definitions, ExcelNameComparison comparison)
    {
        var stringComparison = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return (definitions ?? Array.Empty<IExcelDynamicMappingColumn>()).FirstOrDefault(definition =>
            string.Equals(definition.Title, header, stringComparison)
            || (definition.Aliases ?? Array.Empty<string>()).Any(alias =>
                string.Equals(alias, header, stringComparison)));
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
    /// <param name="unique">维护已提交和当前行待提交值的唯一性跟踪器。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="validationMode">当前导入校验模式。</param>
    /// <param name="unsupportedFeaturePolicy">不支持原生规则的处理策略。</param>
    /// <param name="workbookValidation">XLSX 原生校验快照。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static object MaterializeWorkbookRow<TWorkbook>(DataReader reader, ExcelSheetImportRequest request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        UniqueTracker unique, ICollection<ExcelImportError> errors,
        ExcelResourceLimits resourceLimits, ExcelImportValidationMode validationMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        ExcelXlsxWorkbookValidation workbookValidation, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var method = typeof(ExcelDataReaderExcelImporter).GetMethod(nameof(MaterializeWorkbookRowCore),
            BindingFlags.Static | BindingFlags.NonPublic).MakeGenericMethod(request.ItemType);
        return method.Invoke(null, new object[] { reader, request, plan, bindings, sheetName, rowNumber,
            unique, errors, resourceLimits, validationMode, unsupportedFeaturePolicy, workbookValidation,
            cancellationToken });
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
    /// <param name="unique">维护已提交和当前行待提交值的唯一性跟踪器。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="workbookValidation">XLSX 原生校验快照。</param>
    /// <param name="unsupportedFeaturePolicy">不支持原生规则的处理策略。</param>
    /// <param name="dynamicTargetGetter">读取请求指定动态目标字典的委托。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static TItem MaterializeTypedRow<TItem>(DataReader reader, ExcelBatchImportRequest<TItem> request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        UniqueTracker unique, ICollection<ExcelImportError> errors,
        ExcelXlsxWorkbookValidation workbookValidation,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        Func<object, object> dynamicTargetGetter,
        CancellationToken cancellationToken) where TItem : class, new()
    {
        var item = new TItem();
        var valid = true;
        var initialErrorCount = errors.Count;
        unique.BeginRow();
        var validationEnabled = request.ValidationMode == ExcelImportValidationMode.ConfiguredRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook;
        var workbookValidationEnabled = request.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook;
        Dictionary<string, object> dynamicValues = null;
        foreach (var binding in bindings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var raw = binding.Index < reader.FieldCount ? reader.GetValue(binding.Index) : null;
            var text = NormalizeText(ReadText(raw, request.Culture),
                binding.Column.ImportWhitespace ?? request.BodyWhitespace);
            var property = binding.IsDynamic ? null : typeof(TItem).GetProperty(binding.Column.Name,
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
            if (!binding.IsDynamic && (property == null || !property.CanWrite))
                continue;
            var cell = CreateCell(raw, text);
            var bindingName = binding.Key;
            var propertyType = binding.IsDynamic
                ? ResolveDynamicType(binding.DynamicColumn?.DataTypeName)
                : property.PropertyType;
            var validationBindings = binding.IsDynamic
                ? binding.DynamicColumn?.ValidationBindings ?? binding.Column.ValidationBindings
                : binding.Column.ValidationBindings;
            try
            {
                if (workbookValidationEnabled && workbookValidation != null)
                {
                    var nativeValidationFailed = false;
                    foreach (var nativeResult in workbookValidation.ValidateCell(sheetName, rowNumber - 1,
                        binding.Index, raw, text, request.Culture, cancellationToken))
                    {
                        if (nativeResult.IsValid)
                            continue;
                        errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation,
                            nativeResult.Message, sheetName, rowNumber, binding.Index + 1,
                            bindingName, bindingName, binding.Header, raw));
                        var reportUnsupported = nativeResult.IsUnsupported
                            && unsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report;
                        if (reportUnsupported)
                            continue;
                        nativeValidationFailed = true;
                        valid = false;
                        if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                    }
                    if (nativeValidationFailed)
                        continue;
                }
                if (validationEnabled)
                    ValidateBindings(validationBindings, text, null, binding, sheetName,
                        rowNumber, raw, propertyType, cell, request.Culture, errors,
                        unique, request.ResourceLimits, rawOnly: true);
                var converted = binding.IsDynamic
                    ? ConvertDynamicValue(raw, text, binding.DynamicColumn, bindingName, propertyType,
                        sheetName, rowNumber, binding.Index + 1, request.Culture, cell)
                    : ConvertValue(raw, text, binding.Column, propertyType,
                        sheetName, rowNumber, binding.Index + 1, request.Culture, cell);
                if (validationEnabled)
                    ValidateBindings(validationBindings, text, converted, binding, sheetName,
                        rowNumber, raw, propertyType, cell, request.Culture, errors,
                        unique, request.ResourceLimits, rawOnly: false);
                if (binding.IsDynamic)
                {
                    dynamicValues ??= new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                    dynamicValues[bindingName] = converted;
                }
                else
                    property.SetValue(item, converted);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                if (!errors.Any(error => error.RowIndex == rowNumber && error.ColumnIndex == binding.Index + 1
                    && error.PropertyName == bindingName))
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                        sheetName, rowNumber, binding.Index + 1, bindingName,
                        bindingName, binding.Header, raw));
                valid = false;
                if (request.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    break;
            }
        }
        if (valid && dynamicValues != null)
        {
            try
            {
                SetDynamicValues(item, dynamicValues, dynamicTargetGetter,
                    bindings.First(binding => binding.IsDynamic).Column.Name);
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                var dynamicBinding = bindings.First(binding => binding.IsDynamic);
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, exception.Message,
                    sheetName, rowNumber, dynamicBinding.Index + 1, dynamicBinding.Key,
                    dynamicBinding.Key, dynamicBinding.Header));
                valid = false;
            }
        }
        if (!valid || errors.Skip(initialErrorCount).Any(error =>
            error.Code != ExcelImportErrorCode.WorkbookValidation))
        {
            unique.RollbackRow();
            return null;
        }
        unique.CommitRow();
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
    /// <param name="unique">维护已提交和当前行待提交值的唯一性跟踪器。</param>
    /// <param name="errors">用于收集导入错误的容器。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="validationMode">当前导入校验模式。</param>
    /// <param name="unsupportedFeaturePolicy">不支持原生规则的处理策略。</param>
    /// <param name="workbookValidation">XLSX 原生校验快照。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>转换后的实体；当前行转换或校验失败时为 null。</returns>
    private static object MaterializeWorkbookRowCore<TItem>(DataReader reader, ExcelSheetImportRequest request,
        IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings, string sheetName, int rowNumber,
        UniqueTracker unique, ICollection<ExcelImportError> errors,
        ExcelResourceLimits resourceLimits, ExcelImportValidationMode validationMode,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        ExcelXlsxWorkbookValidation workbookValidation, CancellationToken cancellationToken)
        where TItem : class, new()
    {
        var requestAdapter = new ExcelBatchImportRequest<TItem>(sheetName)
        {
            HeaderRowIndex = request.HeaderRowIndex,
            DataRowStartIndex = request.DataRowStartIndex,
            ValidationFailureMode = request.ValidationFailureMode,
            ValidationMode = validationMode,
            Culture = request.Culture,
            MappingConfiguration = request.MappingConfiguration,
            MappingDocument = request.MappingDocument,
            BodyWhitespace = request.BodyWhitespace,
            ResourceLimits = resourceLimits
        };
        return MaterializeTypedRow(reader, requestAdapter, plan, bindings, sheetName, rowNumber,
            unique, errors, workbookValidation, unsupportedFeaturePolicy, request.DynamicTargetGetter,
            cancellationToken);
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
    /// <param name="unique">维护已提交和当前行待提交值的唯一性跟踪器。</param>
    /// <param name="resourceLimits">导入资源限制。</param>
    /// <param name="rawOnly">为 true 时只执行原始值校验；否则执行转换后校验和唯一性检查。</param>
    private static void ValidateBindings(IReadOnlyList<IExcelValidationBinding> bindings, string text,
        object converted, ColumnBinding binding, string sheetName, int rowNumber, object raw, Type propertyType,
        ExcelCellValue cell, CultureInfo culture, ICollection<ExcelImportError> errors,
        UniqueTracker unique, ExcelResourceLimits resourceLimits, bool rawOnly)
    {
        var bindingName = binding.Key;
        var isUnique = binding.IsDynamic
            ? binding.DynamicColumn?.IsUnique ?? binding.Column.IsUnique
            : binding.Column.IsUnique;
        var uniqueIgnoreEmpty = binding.IsDynamic
            ? binding.DynamicColumn?.UniqueIgnoreEmpty ?? binding.Column.UniqueIgnoreEmpty
            : binding.Column.UniqueIgnoreEmpty;
        foreach (var validation in bindings ?? Array.Empty<IExcelValidationBinding>())
        {
            if (validation.Kind == ExcelValidationBindingKind.Unique || validation.IsRaw != rawOnly)
                continue;
            var context = new ExcelValidationContext(text, sheetName, rowNumber, binding.Index + 1,
                bindingName, converted, propertyType, cell, culture);
            if (!validation.Validate(context))
                errors.Add(new ExcelImportError(ErrorCode(validation), validation.ErrorMessage, sheetName,
                    rowNumber, binding.Index + 1, bindingName, bindingName, binding.Header, raw));
        }
        if (!rawOnly && isUnique)
        {
            try
            {
                if (!unique.TryReserve(bindingName, text,
                    ignoreNull: uniqueIgnoreEmpty,
                    ignoreEmpty: uniqueIgnoreEmpty, rowNumber))
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.Validation, "重复数据。", sheetName,
                        rowNumber, binding.Index + 1, bindingName, bindingName, binding.Header, raw));
            }
            catch (InvalidOperationException) when (resourceLimits?.MaxTrackedUniqueValues is int maximum)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    $"唯一值跟踪数量超过资源限制: {maximum}", sheetName, rowNumber,
                    binding.Index + 1, bindingName, bindingName, binding.Header, raw));
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
    /// 将动态列单元格值转换为定义指定的目标类型。
    /// </summary>
    /// <param name="raw">转换前的原始值。</param>
    /// <param name="text">按空白策略规范化后的文本。</param>
    /// <param name="column">动态列定义；未定义时按字符串处理。</param>
    /// <param name="key">动态列稳定键。</param>
    /// <param name="propertyType">动态列目标类型。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="rowNumber">从 1 开始的物理行号。</param>
    /// <param name="columnIndex">从 1 开始的物理列号。</param>
    /// <param name="culture">值转换使用的区域性。</param>
    /// <param name="cell">单元格原始值和类型上下文。</param>
    /// <returns>转换后的动态列值。</returns>
    private static object ConvertDynamicValue(object raw, string text, IExcelDynamicMappingColumn column,
        string key, Type propertyType, string sheetName, int rowNumber, int columnIndex,
        CultureInfo culture, ExcelCellValue cell)
    {
        var context = new ExcelConversionContext(raw, key, propertyType, sheetName, rowNumber,
            columnIndex, culture, cell);
        foreach (var converter in column?.ValueConverters ?? Array.Empty<IExcelValueConverter>())
            if (converter.TryConvertFrom(context, out var converted))
                return converted;
        var target = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (string.IsNullOrWhiteSpace(text))
        {
            if (!target.IsValueType || Nullable.GetUnderlyingType(propertyType) != null)
                return null;
            throw new InvalidCastException($"值转换失败。输入值为空，目标类型为: {propertyType.FullName}");
        }
        if (target == typeof(string) || target == typeof(object))
            return target == typeof(string) ? text : raw ?? text;
        if (target.IsEnum)
            return Enum.Parse(target, text, true);
        if (target == typeof(Guid))
            return Guid.Parse(text);
        if (target == typeof(DateTime))
            return raw is DateTime ? raw : DateTime.Parse(text, culture);
        if (target == typeof(DateTimeOffset))
            return raw is DateTimeOffset ? raw : DateTimeOffset.Parse(text, culture);
        return Convert.ChangeType(raw is string ? text : raw, target, culture);
    }

    /// <summary>
    /// 将动态列类型名称解析为 CLR 类型。
    /// </summary>
    /// <param name="name">允许列表中的类型名称；为空时按字符串处理。</param>
    /// <returns>动态列使用的 CLR 类型。</returns>
    private static Type ResolveDynamicType(string name) => (name ?? "string").Trim().ToLowerInvariant() switch
    {
        "object" => typeof(object),
        "string" => typeof(string),
        "boolean" or "bool" => typeof(bool),
        "byte" => typeof(byte),
        "int16" => typeof(short),
        "int32" or "int" => typeof(int),
        "int64" or "long" => typeof(long),
        "single" or "float" => typeof(float),
        "double" => typeof(double),
        "decimal" => typeof(decimal),
        "datetime" => typeof(DateTime),
        "datetimeoffset" => typeof(DateTimeOffset),
        "guid" => typeof(Guid),
        "bytes" => typeof(byte[]),
        _ => throw new BingOfficesConfigurationException($"动态列数据类型不在允许列表中: {name}",
            stage: BingOfficesStage.Plan)
    };

    /// <summary>
    /// 将当前行动态值写入请求指定或映射声明的字典属性。
    /// </summary>
    /// <param name="item">当前导入实体。</param>
    /// <param name="values">当前行动态列值。</param>
    /// <param name="dynamicTargetGetter">请求指定的动态目标读取器。</param>
    /// <param name="dynamicPropertyName">映射声明的动态字典属性名称。</param>
    private static void SetDynamicValues<TItem>(TItem item, IReadOnlyDictionary<string, object> values,
        Func<object, object> dynamicTargetGetter, string dynamicPropertyName) where TItem : class, new()
    {
        if (dynamicTargetGetter != null)
        {
            var requestedTarget = dynamicTargetGetter.Invoke(item) as IDictionary<string, object>;
            if (requestedTarget == null)
                return;
            foreach (var pair in values)
                requestedTarget[pair.Key] = pair.Value;
            return;
        }
        var property = typeof(TItem).GetProperty(dynamicPropertyName,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
        if (property == null || !property.CanWrite)
            return;
        var target = property.GetValue(item) as IDictionary<string, object>;
        if (target == null)
        {
            target = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            property.SetValue(item, target);
        }
        foreach (var pair in values)
            target[pair.Key] = pair.Value;
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
    /// 创建固定列导入使用的唯一值跟踪器。
    /// </summary>
    /// <param name="resourceLimits">当前导入的资源限制。</param>
    /// <returns>按资源限制配置的唯一值跟踪器。</returns>
    private static UniqueTracker CreateUniqueTracker(ExcelResourceLimits resourceLimits)
        => new UniqueTracker(new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase),
            resourceLimits?.MaxTrackedUniqueValues,
            GetStringComparer(resourceLimits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase));

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
    /// 获取工作表选择器的诊断文本。
    /// </summary>
    /// <param name="request">工作表导入请求。</param>
    /// <returns>选择器名称或物理索引。</returns>
    private static string DescribeSelector(ExcelSheetImportRequest request)
        => request.Selector.Kind == ExcelSheetSelectorKind.ByIndex
            ? $"#{request.Selector.Index.Value}"
            : request.Selector.Name;

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
    /// <param name="allowWorkbookValidation">是否允许批量路径使用 Workbook 原生校验。</param>
    /// <param name="allowDynamicColumns">是否允许批量路径使用动态列。</param>
    /// <param name="allowWorkbookRelations">是否允许 Workbook 分批路径使用跨 Sheet 关系。</param>
    /// <param name="allowFailureWorkbook">是否允许当前入口生成支持的失败工作簿模式。</param>
    private static void ValidateSupportedRequest<TWorkbook>(ExcelWorkbookImportRequest<TWorkbook> request,
        bool allowWorkbookValidation = false, bool allowDynamicColumns = false,
        bool allowWorkbookRelations = false, bool allowFailureWorkbook = false)
        where TWorkbook : class, new()
    {
        if (!allowWorkbookRelations && request.Relations.Count > 0)
            throw Unsupported("跨 Sheet 关系");
        if (request.FailureOptions != null
            && request.FailureOptions.Mode != ExcelImportFailureWorkbookMode.None)
        {
            if (!allowFailureWorkbook)
                throw Unsupported("Failure Workbook");
        }
        if (!allowWorkbookValidation && (request.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook))
            throw Unsupported("Workbook 原生校验");
        if (!allowDynamicColumns && request.Sheets.Any(sheet => sheet.DynamicColumns.Count > 0))
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
        if (request.ResourceLimits != null
            && (request.ResourceLimits.MaxPictures.HasValue
                || request.ResourceLimits.MaxPictureBytes.HasValue
                || request.ResourceLimits.MaxTotalPictureBytes.HasValue))
            throw Unsupported("图片资源限制");
    }

    /// <summary>
    /// 验证工作簿分批请求的资源和工作表配置边界。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿类型。</typeparam>
    /// <param name="request">工作簿分批请求。</param>
    /// <returns>按请求项缓存的映射计划。</returns>
    private Dictionary<ExcelSheetImportRequest, IExcelMappingPlan> ValidateSupportedWorkbookBatchRequest<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request)
        where TWorkbook : class, new()
    {
        if (request.Sheets.Count == 0)
            throw new ArgumentException("Workbook 至少需要一个 Sheet。", nameof(request));
        if (request.ResourceLimits != null
            && (request.ResourceLimits.MaxPictures.HasValue
                || request.ResourceLimits.MaxPictureBytes.HasValue
                || request.ResourceLimits.MaxTotalPictureBytes.HasValue))
            throw Unsupported("图片资源限制");
        var plans = new Dictionary<ExcelSheetImportRequest, IExcelMappingPlan>();
        foreach (var sheet in request.Sheets)
        {
            if (sheet.BatchSize <= 0)
                throw new ArgumentOutOfRangeException(nameof(sheet.BatchSize));
            var plan = CreatePlan(sheet);
            ValidateBatchPlan(plan, sheet.ItemType);
            plans.Add(sheet, plan);
        }
        return plans;
    }

    /// <summary>
    /// 交付工作簿级资源错误并生成对应摘要。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿类型。</typeparam>
    /// <param name="request">工作簿分批请求。</param>
    /// <param name="callback">同步批次回调。</param>
    /// <param name="resourceError">资源限制错误及其工作表上下文。</param>
    /// <returns>包含资源错误的工作簿摘要。</returns>
    private static ExcelWorkbookBatchImportSummary DeliverWorkbookResourceError<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request, Action<ExcelWorkbookImportBatch> callback,
        WorkbookResourceBudgetError resourceError) where TWorkbook : class, new()
    {
        var sheet = ResolveResourceErrorSheet(request, resourceError);
        var sheetName = resourceError.Error.SheetName ?? sheet.Name;
        InvokeWorkbookBatch(callback, new ExcelWorkbookImportBatch(sheetName, sheet.ItemType, 1,
            resourceError.Error.RowIndex, Array.Empty<object>(), new[] { resourceError.Error }));
        return new ExcelWorkbookBatchImportSummary(new[]
        {
            new ExcelBatchImportSheetSummary(sheetName, sheet.ItemType,
                new ExcelBatchImportSummary(0, 0, 0, 1, false, true))
        });
    }

    /// <summary>
    /// 异步交付工作簿级资源错误并生成对应摘要。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿类型。</typeparam>
    /// <param name="request">工作簿分批请求。</param>
    /// <param name="callback">异步批次回调。</param>
    /// <param name="resourceError">资源限制错误及其工作表上下文。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>包含资源错误的工作簿摘要任务。</returns>
    private static async Task<ExcelWorkbookBatchImportSummary> DeliverWorkbookResourceErrorAsync<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request,
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> callback,
        WorkbookResourceBudgetError resourceError, CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var sheet = ResolveResourceErrorSheet(request, resourceError);
        var sheetName = resourceError.Error.SheetName ?? sheet.Name;
        await InvokeWorkbookBatchAsync(callback, new ExcelWorkbookImportBatch(sheetName, sheet.ItemType, 1,
            resourceError.Error.RowIndex, Array.Empty<object>(), new[] { resourceError.Error }),
            cancellationToken).ConfigureAwait(false);
        return new ExcelWorkbookBatchImportSummary(new[]
        {
            new ExcelBatchImportSheetSummary(sheetName, sheet.ItemType,
                new ExcelBatchImportSummary(0, 0, 0, 1, false, true))
        });
    }

    /// <summary>
    /// 根据资源错误的物理工作表上下文解析请求项。
    /// </summary>
    /// <typeparam name="TWorkbook">具有公共无参构造函数的工作簿类型。</typeparam>
    /// <param name="request">工作簿分批请求。</param>
    /// <param name="resourceError">包含物理工作表上下文的资源错误。</param>
    /// <returns>与错误工作表匹配的请求项；无法匹配时返回第一个请求项。</returns>
    private static ExcelSheetImportRequest ResolveResourceErrorSheet<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request, WorkbookResourceBudgetError resourceError)
        where TWorkbook : class, new()
    {
        var matched = request.Sheets.FirstOrDefault(sheet => resourceError.SheetIndex >= 0
            ? IsSelected(sheet, resourceError.Error.SheetName, resourceError.SheetIndex,
                request.SheetNameComparison)
            : string.Equals(sheet.Name, resourceError.Error.SheetName,
                request.SheetNameComparison == ExcelNameComparison.Ordinal
                    ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase));
        return matched ?? request.Sheets[0];
    }

    /// <summary>
    /// 验证分批导入计划中的动态列目标。
    /// </summary>
    /// <param name="plan">实体导入映射计划。</param>
    /// <param name="itemType">当前工作表实体类型。</param>
    private static void ValidateBatchPlan(IExcelMappingPlan plan, Type itemType)
    {
        var dynamicProperties = plan.Columns.Where(column => !column.Ignored && column.IsDynamicColumn).ToArray();
        if (dynamicProperties.Length > 1)
            throw new BingOfficesConfigurationException("一个导入实体只能声明一个动态列属性。",
                stage: BingOfficesStage.Plan);
        if (plan.DynamicColumns.Count == 0 && dynamicProperties.Length == 0)
            return;
        if (dynamicProperties.Length == 0)
            throw Unsupported("动态列缺少可写字典属性");
        var property = itemType.GetProperty(dynamicProperties[0].Name,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
        if (property == null || !property.CanWrite
            || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
            throw new BingOfficesConfigurationException(
                $"动态列属性不可写入 IDictionary<string, object>: {dynamicProperties[0].Name}",
                stage: BingOfficesStage.Plan);
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
    /// 交付工作簿批次并包装回调异常。
    /// </summary>
    /// <param name="callback">接收工作簿批次的调用方回调。</param>
    /// <param name="batch">包含工作表、实体和错误的批次。</param>
    private static void InvokeWorkbookBatch(Action<ExcelWorkbookImportBatch> callback,
        ExcelWorkbookImportBatch batch)
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
    /// 异步交付工作簿批次并包装回调异常。
    /// </summary>
    /// <param name="callback">接收工作簿批次的调用方回调。</param>
    /// <param name="batch">包含工作表、实体和错误的批次。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static async Task InvokeWorkbookBatchAsync(
        Func<ExcelWorkbookImportBatch, CancellationToken, Task> callback,
        ExcelWorkbookImportBatch batch, CancellationToken cancellationToken)
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
    /// <returns>工作表数量超限错误及物理上下文；未限制或未超限时为 null。</returns>
    private static WorkbookResourceBudgetError ScanWorkbookResourceBudgets(Stream source,
        ExcelResourceLimits limits, CancellationToken cancellationToken)
    {
        if (limits?.MaxSheets == null)
            return null;
        if (!source.CanSeek)
            throw new InvalidOperationException("ExcelDataReader 资源预检需要可定位的暂存输入。");

        WorkbookResourceBudgetError resourceError = null;
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
                        resourceError = new WorkbookResourceBudgetError(
                            new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                                "工作表数量超过资源限制。", reader.Name, 1, 0, null), sheetIndex);
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
    /// 按请求校验模式创建 XLSX 原生校验快照。
    /// </summary>
    /// <param name="input">已暂存的输入流。</param>
    /// <param name="validationMode">导入校验模式。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>启用原生校验时返回快照，否则返回 null。</returns>
    private static ExcelXlsxWorkbookValidation CreateWorkbookValidation(Stream input,
        ExcelImportValidationMode validationMode,
        CancellationToken cancellationToken)
    {
        var enabled = validationMode == ExcelImportValidationMode.WorkbookRules
            || validationMode == ExcelImportValidationMode.ConfiguredAndWorkbook;
        if (!enabled)
            return null;
        if (!IsXlsxXmlWorkbook(input))
            throw Unsupported("XLS/XLSB 无法提供 Workbook 原生校验");
        cancellationToken.ThrowIfCancellationRequested();
        input.Position = 0;
        var snapshot = ExcelXlsxWorkbookValidation.Load(input, cancellationToken);
        input.Position = 0;
        return snapshot;
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
    /// 验证 AnnotatedOriginal 的来源是可保留部件的 XML XLSX。
    /// </summary>
    /// <param name="options">失败工作簿选项。</param>
    /// <param name="input">已暂存的输入流。</param>
    private static void ValidateFailureWorkbookSource(ExcelImportFailureOptions options, Stream input)
    {
        if (options?.Mode != ExcelImportFailureWorkbookMode.AnnotatedOriginal)
            return;
        if (!IsXlsxXmlWorkbook(input))
            throw Unsupported("AnnotatedOriginal 仅支持 XLSX XML 工作簿");
    }

    /// <summary>
    /// 捕获 ErrorRowsOnly XLSX 输出能够安全重映射的原生校验和名称元数据。
    /// </summary>
    /// <param name="input">已暂存的输入流。</param>
    /// <param name="options">失败工作簿选项。</param>
    /// <param name="capture">失败工作簿捕获器。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void CaptureFailureWorkbookMetadata(Stream input,
        ExcelImportFailureOptions options, ExcelDataReaderFailureWorkbookCapture capture,
        CancellationToken cancellationToken)
    {
        if (capture == null || options?.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly
            || !IsXlsxXmlWorkbook(input))
            return;

        var originalPosition = input.Position;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            input.Position = 0;
            using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
            capture.CaptureWorkbookMetadata(
                ExcelXlsxWorkbookValidation.LoadFailureMetadata(archive, cancellationToken));
        }
        finally
        {
            input.Position = originalPosition;
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
        /// 初始化一个 <see cref="ColumnBinding" /> 类型的实例。
        /// </summary>
        /// <param name="column">对应列的映射配置。</param>
        /// <param name="index">从 0 开始的物理索引。</param>
        /// <param name="header">匹配到的表头文本。</param>
        /// <param name="dynamicColumn">动态列定义；固定列时为 null。</param>
        internal ColumnBinding(IExcelMappingColumn column, int index, string header,
            IExcelDynamicMappingColumn dynamicColumn = null)
        {
            Column = column;
            Index = index;
            Header = header;
            DynamicColumn = dynamicColumn;
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
        /// <summary>
        /// 获取动态列定义；固定列时为 null。
        /// </summary>
        internal IExcelDynamicMappingColumn DynamicColumn { get; }
        /// <summary>
        /// 获取是否为动态列绑定。
        /// </summary>
        internal bool IsDynamic => Column.IsDynamicColumn;
        /// <summary>
        /// 获取用于错误定位和动态字典访问的列键。
        /// </summary>
        internal string Key => DynamicColumn?.Key ?? (IsDynamic ? Header : Column.Name);
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
        /// 初始化一个 <see cref="RowBudget" /> 类型的实例。
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
        /// 初始化一个 <see cref="CellBudget" /> 类型的实例。
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
        /// 初始化一个 <see cref="ErrorCollector" /> 类型的实例。
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

        /// <summary>
        /// 标记后续关系阶段因达到错误预算而截断。
        /// </summary>
        internal void MarkTruncated() => Truncated = true;
    }

    /// <summary>
    /// 用于区分调用方批次回调失败的异常。
    /// </summary>
    private sealed class BatchCallbackException : Exception
    {
        /// <summary>
        /// 初始化一个 <see cref="BatchCallbackException" /> 类型的实例。
        /// </summary>
        /// <param name="innerException">原始异常；没有原始异常时为 null。</param>
        internal BatchCallbackException(Exception innerException)
            : base("分批回调执行失败。", innerException)
        {
        }
    }

    /// <summary>
    /// 保存工作簿资源错误及其物理工作表上下文。
    /// </summary>
    private sealed class WorkbookResourceBudgetError
    {
        /// <summary>
        /// 初始化一个 <see cref="WorkbookResourceBudgetError" /> 类型的实例。
        /// </summary>
        /// <param name="error">资源限制错误。</param>
        /// <param name="sheetIndex">错误工作表的从 0 开始物理索引；未知时为负数。</param>
        internal WorkbookResourceBudgetError(ExcelImportError error, int sheetIndex)
        {
            Error = error ?? throw new ArgumentNullException(nameof(error));
            SheetIndex = sheetIndex;
        }

        /// <summary>
        /// 获取资源限制错误。
        /// </summary>
        internal ExcelImportError Error { get; }

        /// <summary>
        /// 获取错误工作表的从 0 开始物理索引。
        /// </summary>
        internal int SheetIndex { get; }
    }

    /// <summary>
    /// 维护工作簿分批回调使用的全局批次号。
    /// </summary>
    private sealed class BatchNumberState
    {
        /// <summary>
        /// 下一个待交付的批次号。
        /// </summary>
        private int _next = 1;

        /// <summary>
        /// 获取当前批次号并推进计数。
        /// </summary>
        /// <returns>从一开始的批次号。</returns>
        internal int Next() => _next++;
    }
}
