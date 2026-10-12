using System.Collections;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Entities;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Imports;

/// <summary>
/// 基于 ClosedXML 的 XLSX 工作簿导入器。
/// </summary>
public sealed class ClosedXmlExcelImporter : IExcelImporter, IExcelEntityImporter,
    IExcelEntityResourceImporter, IExcelProviderFeatureDescriptor
{
    /// <summary>
    /// ClosedXML Provider 名称。
    /// </summary>
    private const string Provider = "ClosedXML";

    /// <summary>
    /// ClosedXML 映射计划构建器。
    /// </summary>
    private readonly ClosedXmlMappingPlanBuilder _planBuilder;

    /// <summary>
    /// Bing.Offices 异常观察分发器。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 实体布局执行器。
    /// </summary>
    private readonly ClosedXmlEntityLayoutExecutor _entityExecutor;

    /// <summary>
    /// Workbook DOM 准入器。
    /// </summary>
    private readonly IClosedXmlWorkbookAdmission _admission;

    /// <summary>
    /// 失败工作簿的文件提交服务。
    /// </summary>
    private readonly IFileExportCommitter _fileExportCommitter;

    /// <summary>
    /// 实体布局失败输出使用的空 Workbook 请求映射。
    /// </summary>
    private static readonly IReadOnlyDictionary<string, ExcelSheetImportRequest>
        EmptyResolvedSheetRequests = new Dictionary<string, ExcelSheetImportRequest>(
            StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 获取实体布局执行器，供 Provider 内部批次路径复用。
    /// </summary>
    internal ClosedXmlEntityLayoutExecutor EntityExecutor => _entityExecutor;

    /// <summary>
    /// 获取失败工作簿使用的文件提交服务。
    /// </summary>
    internal IFileExportCommitter FileExportCommitter => _fileExportCommitter;

    /// <summary>
    /// 观察 Provider 内部批次路径产生的公共异常。
    /// </summary>
    /// <param name="exception">待观察的异常。</param>
    internal void ObserveException(BingOfficesException exception)
        => _exceptionDispatcher.Observe(exception);

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelImporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 使用进程级默认 Workbook DOM 准入器。
    /// </remarks>
    /// <param name="validationRules">可选的公共校验规则集合。</param>
    /// <param name="valueConverters">可选的值转换器集合。</param>
    /// <param name="namedValidationRules">可选的命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">可选的公共映射计划工厂。</param>
    /// <param name="exceptionObservers">可选的异常观察器集合。</param>
    public ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory,
            exceptionObservers, null, ClosedXmlWorkbookAdmission.SharedDefault)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelImporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 使用指定的文件提交服务输出失败工作簿文件；为空时使用默认实现。
    /// </remarks>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">公共映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    /// <param name="fileExportCommitter">文件提交服务；为空时使用默认原子提交实现。</param>
    public ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers,
        IFileExportCommitter fileExportCommitter)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory,
            exceptionObservers, fileExportCommitter, ClosedXmlWorkbookAdmission.SharedDefault)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelImporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 使用调用方提供的 Workbook DOM 准入器。
    /// </remarks>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">公共映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    /// <param name="admission">Workbook DOM 准入器。</param>
    internal ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers,
        IClosedXmlWorkbookAdmission admission)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory,
            exceptionObservers, null, admission)
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlExcelImporter" /> 类型的实例。
    /// </summary>
    /// <remarks>
    /// 同时配置文件提交服务与内部 Workbook DOM 准入器。
    /// </remarks>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">公共映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    /// <param name="fileExportCommitter">文件提交服务；为空时使用默认原子提交实现。</param>
    /// <param name="admission">Workbook DOM 准入器。</param>
    internal ClosedXmlExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers,
        IFileExportCommitter fileExportCommitter,
        IClosedXmlWorkbookAdmission admission)
    {
        var converters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        var factory = mappingPlanFactory ??
            ExcelMappingPlanFactoryProvider.CreateDefault(valueConverters: converters,
                validationRules: validationRules, namedValidationRules: namedValidationRules);
        _planBuilder = new ClosedXmlMappingPlanBuilder(factory);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
        _entityExecutor = new ClosedXmlEntityLayoutExecutor(factory);
        _fileExportCommitter = fileExportCommitter ?? new DefaultFileExportCommitter();
        _admission = admission ?? throw new ArgumentNullException(nameof(admission));
    }

    /// <inheritdoc />
    public string ProviderName => Provider;

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Entity
        | ExcelProviderCapabilities.Template
        | ExcelProviderCapabilities.Merge | ExcelProviderCapabilities.Async
        | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = new[] { ExcelFormat.Xlsx };
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
        "XLS、XLSB、XLSM、ODS 和加密工作簿在当前 ClosedXML Provider 中明确拒绝。"
    };
    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.TemplateEditing
        | ExcelProviderFeatures.WorkbookEditing | ExcelProviderFeatures.FormulaText
        | ExcelProviderFeatures.FormulaCachedValues | ExcelProviderFeatures.NamedListValidation
        | ExcelProviderFeatures.MultiSheetBatchImport | ExcelProviderFeatures.EntityLayoutBatchImport;
    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;

    /// <inheritdoc />
    public ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
        where TEntity : class, new()
        => ImportEntity(source, layout, new ExcelEntityImportOptions(), cancellationToken);

    /// <inheritdoc />
    public ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityImportOptions options,
        CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        ValidateEntitySource(source, layout, cancellationToken);
        var limits = ValidateEntityOptions(options);
        try
        {
            var buffered = BufferSource(source, limits, cancellationToken);
            ExcelXlsxZipPreflight.Validate(buffered, limits, Provider, requireZip: true,
                cancellationToken: cancellationToken, includePictures: true);
            var isDate1904 = ExcelXlsxZipPreflight.GetDate1904(buffered, cancellationToken);
            using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Import);
            using var workbook = new XLWorkbook(buffered);
            var result = _entityExecutor.Read(workbook, layout, requireTemplateMerges: false,
                isDate1904, limits, options.ValidationFailureMode, options.ValidationMode,
                options.UnsupportedFeaturePolicy, cancellationToken);
            WriteEntityFailureWorkbook(workbook, layout, options.FailureOptions, result.Errors,
                cancellationToken);
            return result;
        }
        catch (OperationCanceledException)
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
            var translated = new BingOfficesImportException("ClosedXML 实体导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <inheritdoc />
    public async Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
        where TEntity : class, new()
        => await ImportEntityAsync(source, layout, new ExcelEntityImportOptions(), cancellationToken)
            .ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityImportOptions options,
        CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        ValidateEntitySource(source, layout, cancellationToken);
        var limits = ValidateEntityOptions(options);
        try
        {
            await using var buffer = await BufferSourceAsync(source, limits,
                cancellationToken).ConfigureAwait(false);
            ExcelXlsxZipPreflight.Validate(buffer, limits, Provider,
                requireZip: true, cancellationToken: cancellationToken, includePictures: true);
            var isDate1904 = ExcelXlsxZipPreflight.GetDate1904(buffer, cancellationToken);
            await using var admission = await _admission
                .AcquireAsync(cancellationToken, BingOfficesOperation.Import)
                .ConfigureAwait(false);
            using var workbook = new XLWorkbook(buffer);
            var result = _entityExecutor.Read(workbook, layout, requireTemplateMerges: false,
                isDate1904, limits, options.ValidationFailureMode, options.ValidationMode,
                options.UnsupportedFeaturePolicy, cancellationToken);
            await WriteEntityFailureWorkbookAsync(workbook, layout, options.FailureOptions,
                result.Errors, cancellationToken).ConfigureAwait(false);
            return result;
        }
        catch (OperationCanceledException)
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
            var translated = new BingOfficesImportException("ClosedXML 异步实体导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <inheritdoc />
    public ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => ImportForTemplate(source, layout, template, new ExcelEntityImportOptions(), cancellationToken);

    /// <inheritdoc />
    public ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var limits = ValidateEntityOptions(options);
        try
        {
            ValidateEntitySource(source, layout, cancellationToken);
            ValidateEntityTemplate(template, layout, limits, cancellationToken);
            var buffered = BufferSource(source, limits, cancellationToken);
            ExcelXlsxZipPreflight.Validate(buffered, limits, Provider, requireZip: true,
                cancellationToken: cancellationToken, includePictures: true);
            var isDate1904 = ExcelXlsxZipPreflight.GetDate1904(buffered, cancellationToken);
            using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Import);
            using var workbook = new XLWorkbook(buffered);
            var result = _entityExecutor.Read(workbook, layout, requireTemplateMerges: false,
                isDate1904, limits, options.ValidationFailureMode, options.ValidationMode,
                options.UnsupportedFeaturePolicy, cancellationToken);
            WriteEntityFailureWorkbook(workbook, layout, options.FailureOptions, result.Errors,
                cancellationToken);
            return result;
        }
        catch (OperationCanceledException)
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
            var translated = new BingOfficesImportException("ClosedXML 模板实体导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (!template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <inheritdoc />
    public async Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => await ImportForTemplateAsync(source, layout, template, new ExcelEntityImportOptions(), cancellationToken)
            .ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityImportOptions options, CancellationToken cancellationToken = default)
        where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var limits = ValidateEntityOptions(options);
        try
        {
            ValidateEntitySource(source, layout, cancellationToken);
            await ValidateEntityTemplateAsync(template, layout, limits, cancellationToken).ConfigureAwait(false);
            await using var buffer = await BufferSourceAsync(source, limits, cancellationToken)
                .ConfigureAwait(false);
            ExcelXlsxZipPreflight.Validate(buffer, limits, Provider, requireZip: true,
                cancellationToken: cancellationToken, includePictures: true);
            var isDate1904 = ExcelXlsxZipPreflight.GetDate1904(buffer, cancellationToken);
            await using var admission = await _admission
                .AcquireAsync(cancellationToken, BingOfficesOperation.Import)
                .ConfigureAwait(false);
            using var workbook = new XLWorkbook(buffer);
            var result = _entityExecutor.Read(workbook, layout, requireTemplateMerges: false,
                isDate1904, limits, options.ValidationFailureMode, options.ValidationMode,
                options.UnsupportedFeaturePolicy, cancellationToken);
            await WriteEntityFailureWorkbookAsync(workbook, layout, options.FailureOptions,
                result.Errors, cancellationToken).ConfigureAwait(false);
            return result;
        }
        catch (OperationCanceledException)
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
            var translated = new BingOfficesImportException("ClosedXML 异步模板实体导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (!template.LeaveOpen)
                template.Template.Dispose();
        }
    }

    /// <summary>
    /// 写出实体导入失败工作簿。
    /// </summary>
    /// <param name="workbook">实体导入使用的原始工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">实体导入错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void WriteEntityFailureWorkbook<TEntity>(XLWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var requiredRows = options != null
            && options.Mode != ExcelImportFailureWorkbookMode.None
            && errors.Count > 0
            ? BuildFailureWorkbookRequiredRows(workbook, layout, cancellationToken)
            : null;
        using var artifact = ClosedXmlFailureWorkbookWriter.Create(workbook, options, errors,
            EmptyResolvedSheetRequests, cancellationToken, requiredRows);
        if (artifact == null)
            return;
        if (!string.IsNullOrWhiteSpace(options.DestinationPath))
            _fileExportCommitter.Commit(options.DestinationPath,
                destination => artifact.CopyTo(destination, cancellationToken), cancellationToken,
                "FailureWorkbook");
        else
            artifact.CopyTo(options.Destination, cancellationToken);
    }

    /// <summary>
    /// 异步写出实体导入失败工作簿。
    /// </summary>
    /// <param name="workbook">实体导入使用的原始工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">实体导入错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>表示输出操作的任务。</returns>
    private async Task WriteEntityFailureWorkbookAsync<TEntity>(XLWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var requiredRows = options != null
            && options.Mode != ExcelImportFailureWorkbookMode.None
            && errors.Count > 0
            ? BuildFailureWorkbookRequiredRows(workbook, layout, cancellationToken)
            : null;
        using var artifact = ClosedXmlFailureWorkbookWriter.Create(workbook, options, errors,
            EmptyResolvedSheetRequests, cancellationToken, requiredRows);
        if (artifact == null)
            return;
        if (!string.IsNullOrWhiteSpace(options.DestinationPath))
            await _fileExportCommitter.CommitAsync(options.DestinationPath,
                (destination, token) => artifact.CopyToAsync(destination, token), cancellationToken,
                "FailureWorkbook").ConfigureAwait(false);
        else
            await artifact.CopyToAsync(options.Destination, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 为前向实体批次写出失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">实体导入错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal void WriteEntityFailureWorkbookForBatch<TEntity>(XLWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        CancellationToken cancellationToken) where TEntity : class, new()
        => WriteEntityFailureWorkbook(workbook, layout, options, errors, cancellationToken);

    /// <summary>
    /// 异步为前向实体批次写出失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="workbook">原始导入工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">失败工作簿输出选项。</param>
    /// <param name="errors">实体导入错误集合。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>表示输出操作的任务。</returns>
    internal Task WriteEntityFailureWorkbookForBatchAsync<TEntity>(XLWorkbook workbook,
        ExcelEntityLayout<TEntity> layout, ExcelImportFailureOptions options,
        IReadOnlyCollection<ExcelImportError> errors,
        CancellationToken cancellationToken) where TEntity : class, new()
        => WriteEntityFailureWorkbookAsync(workbook, layout, options, errors, cancellationToken);

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            var limits = request.ResourceLimits ?? new ExcelResourceLimits();
            var buffered = BufferSource(source, limits, cancellationToken);
            ExcelXlsxZipPreflight.Validate(buffered, limits, Provider, requireZip: true,
                cancellationToken: cancellationToken, includePictures: true);
            ValidateRequest(request);
            ValidateFailureWorkbookInput(buffered, request, cancellationToken);
            using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Import);
            var result = ImportBufferedWorkbook(buffered, request, cancellationToken, out var failureArtifact);
            using (failureArtifact)
            {
                if (failureArtifact != null)
                {
                    var failureOptions = request.FailureOptions;
                    if (failureOptions.DestinationPath != null)
                        _fileExportCommitter.Commit(failureOptions.DestinationPath,
                            destination => failureArtifact.CopyTo(destination, cancellationToken), cancellationToken,
                            "FailureWorkbook");
                    else
                        failureArtifact.CopyTo(failureOptions.Destination, cancellationToken);
                }
            }
            return result;
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
        catch (NotSupportedException exception)
        {
            var translated = new BingOfficesUnsupportedFeatureException(
                "当前 ClosedXML 导入功能不受支持。", exception, Provider,
                BingOfficesOperation.Import, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("ClosedXML Excel 导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <inheritdoc />
    public async Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        cancellationToken.ThrowIfCancellationRequested();
        try
        {
            var limits = request.ResourceLimits ?? new ExcelResourceLimits();
            await using var buffer = await BufferSourceAsync(source,
                limits, cancellationToken)
                .ConfigureAwait(false);
            ExcelXlsxZipPreflight.Validate(buffer, limits, Provider, requireZip: true,
                cancellationToken: cancellationToken, includePictures: true);
            ValidateRequest(request);
            ValidateFailureWorkbookInput(buffer, request, cancellationToken);
            await using var admission = await _admission
                .AcquireAsync(cancellationToken, BingOfficesOperation.Import)
                .ConfigureAwait(false);
            var result = ImportBufferedWorkbook(buffer, request, cancellationToken, out var failureArtifact);
            using (failureArtifact)
            {
                if (failureArtifact != null)
                {
                    var failureOptions = request.FailureOptions;
                    if (failureOptions.DestinationPath != null)
                        await _fileExportCommitter.CommitAsync(failureOptions.DestinationPath,
                            (destination, token) => failureArtifact.CopyToAsync(destination, token),
                            cancellationToken, "FailureWorkbook").ConfigureAwait(false);
                    else
                        await failureArtifact.CopyToAsync(failureOptions.Destination, cancellationToken)
                            .ConfigureAwait(false);
                }
            }
            return result;
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
        catch (NotSupportedException exception)
        {
            var translated = new BingOfficesUnsupportedFeatureException(
                "当前 ClosedXML 异步导入功能不受支持。", exception, Provider,
                BingOfficesOperation.Import, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("ClosedXML 异步 Excel 导入失败。", exception,
                Provider, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <summary>
    /// 从已缓冲并完成预检的流创建 ClosedXML 工作簿并执行导入计划。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="buffered">已缓冲的 XLSX 流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="failureArtifact">已完整序列化的失败工作簿临时产物。</param>
    /// <returns>Workbook 导入结果。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportBufferedWorkbook<TWorkbook>(Stream buffered,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken,
        out ClosedXmlFailureWorkbookArtifact failureArtifact)
        where TWorkbook : class, new()
    {
        failureArtifact = null;
        var limits = request.ResourceLimits ?? new ExcelResourceLimits();
        var rowViolation = ClosedXmlRowBudgetPreflight.FindViolation(buffered,
            request.Sheets, request.SheetNameComparison, limits.MaxRows, cancellationToken);
        if (rowViolation != null)
        {
            var error = new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                $"XLSX 数据行数超过限制: {rowViolation.MaximumRows}。",
                rowViolation.SheetName, rowViolation.RowIndex, 0, string.Empty);
            var sheetResult = new ExcelSheetImportResult(rowViolation.SheetName,
                rowViolation.ItemType, Array.Empty<int>(), new[] { error });
            return new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(),
                new[] { sheetResult }, new[] { error }, false, limits.MaxErrors);
        }

        var isDate1904 = ExcelXlsxZipPreflight.GetDate1904(buffered, cancellationToken);
        using var workbook = new XLWorkbook(buffered);
        var root = new TWorkbook();
        var sheetResults = new List<ExcelSheetImportResult>();
        var workbookErrors = new List<ExcelImportError>();
        var relationKeyBudget = new ClosedXmlRelationKeyBudget(
            request.ResourceLimits?.MaxTrackedRelationKeys);
        var resolvedSheetRequests = new Dictionary<string, ExcelSheetImportRequest>(
            StringComparer.OrdinalIgnoreCase);
        var totalRows = 0;
        foreach (var sheetRequest in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var worksheet = ResolveWorksheet(workbook, sheetRequest, request.SheetNameComparison);
            if (worksheet == null)
            {
                AddError(workbookErrors, request, new ExcelImportError(
                    ExcelImportErrorCode.InvalidHeader, $"找不到工作表: {sheetRequest.Name}", sheetRequest.Name,
                    0, 0, string.Empty));
                sheetResults.Add(new ExcelSheetImportResult(sheetRequest.Name, sheetRequest.ItemType,
                    Array.Empty<int>(), workbookErrors.ToArray()));
                continue;
            }
            resolvedSheetRequests[worksheet.Name] = sheetRequest;
            var result = ImportSheet(root, worksheet, sheetRequest, request, ref totalRows,
                workbookErrors, isDate1904, cancellationToken);
            sheetResults.Add(result);
        }
        var relationErrorsTruncated = ClosedXmlRelationCoordinator.Bind(root, request.Relations,
            workbookErrors, request.ResourceLimits?.MaxErrors, relationKeyBudget, cancellationToken);
        var importResult = new ExcelWorkbookImportResult<TWorkbook>(root, sheetResults,
            workbookErrors.ToArray(), relationErrorsTruncated
                || request.ResourceLimits?.MaxErrors.HasValue == true
                && workbookErrors.Count >= request.ResourceLimits.MaxErrors.Value,
            request.ResourceLimits?.MaxErrors);
        failureArtifact = ClosedXmlFailureWorkbookWriter.Create(workbook, request.FailureOptions,
            workbookErrors, resolvedSheetRequests, cancellationToken);
        return importResult;
    }

    /// <summary>
    /// 同步加载并校验实体模板布局。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="template">实体模板选项。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void ValidateEntityTemplate<TEntity>(ExcelEntityTemplateOptions template,
        ExcelEntityLayout<TEntity> layout, ExcelResourceLimits limits, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (!template.Template.CanRead)
            throw new ArgumentException("实体模板流不可读取。", nameof(template));
        ClosedXmlTemplatePreflight.Validate(template.Template, BingOfficesOperation.Import,
            cancellationToken, limits);
        if (template.Template.CanSeek)
            template.Template.Position = 0;
        using var admission = _admission.Acquire(cancellationToken, BingOfficesOperation.Import);
        using var workbook = new XLWorkbook(template.Template);
        _entityExecutor.ValidateTemplate(workbook, layout, cancellationToken);
    }

    /// <summary>
    /// 异步加载并校验实体模板布局。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="template">实体模板选项。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private async Task ValidateEntityTemplateAsync<TEntity>(ExcelEntityTemplateOptions template,
        ExcelEntityLayout<TEntity> layout, ExcelResourceLimits limits, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (!template.Template.CanRead)
            throw new ArgumentException("实体模板流不可读取。", nameof(template));
        ClosedXmlTemplatePreflight.Validate(template.Template, BingOfficesOperation.Import,
            cancellationToken, limits);
        if (template.Template.CanSeek)
            template.Template.Position = 0;
        await using var admission = await _admission
            .AcquireAsync(cancellationToken, BingOfficesOperation.Import)
            .ConfigureAwait(false);
        using var workbook = new XLWorkbook(template.Template);
        _entityExecutor.ValidateTemplate(workbook, layout, cancellationToken);
    }

    /// <summary>
    /// 校验实体导入的输入流、布局和取消状态。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="source">输入流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidateEntitySource<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 收集实体失败工作簿需要保留的固定和列表结构行。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">已打开的实体工作簿。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按工作表分组的工作表行号结构快照。</returns>
    private static IReadOnlyDictionary<string, IReadOnlyList<int>> BuildFailureWorkbookRequiredRows<TEntity>(
        XLWorkbook workbook, ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        var rows = new Dictionary<string, SortedSet<int>>(StringComparer.OrdinalIgnoreCase);

        void Add(string sheetName, int row)
        {
            if (row < 0)
                return;
            if (!rows.TryGetValue(sheetName, out var sheetRows))
            {
                sheetRows = new SortedSet<int>();
                rows.Add(sheetName, sheetRows);
            }
            sheetRows.Add(row + 1);
        }

        void AddRange(string sheetName, int firstRow, int lastRow)
        {
            for (var row = firstRow; row <= lastRow; row++)
                Add(sheetName, row);
        }

        foreach (var cell in layout.Cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var reference = cell.AnchorName == null
                ? cell.Reference
                : ResolveNamedAnchor(workbook, cell.SheetName, cell.AnchorName);
            Add(cell.SheetName, reference.Row);
        }
        foreach (var region in layout.ListRegions.Where(item => item.IncludeHeader))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var start = region.AnchorName == null
                ? region.Start
                : ResolveNamedAnchor(workbook, region.SheetName, region.AnchorName);
            Add(region.SheetName, start.Row);
        }
        foreach (var merge in layout.Merges)
        {
            cancellationToken.ThrowIfCancellationRequested();
            AddRange(merge.SheetName, merge.Range.First.Row, merge.Range.Last.Row);
        }
        foreach (var region in layout.ListRegions)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!workbook.Worksheets.TryGetWorksheet(region.SheetName, out var sheet))
                continue;
            if (region.PageHeader != null)
            {
                AddRange(region.SheetName, region.PageHeader.SourceRange.First.Row,
                    region.PageHeader.SourceRange.Last.Row);
            }
            if (region.Footer != null || region.GroupSubtotal != null
                || region.PageSubtotal != null || region.PageHeader != null
                || region.PageFooter != null)
            {
                var lastRow = region.End?.Row ?? sheet.LastRowUsed()?.RowNumber() - 1
                    ?? region.Start.Row;
                foreach (var span in ClosedXmlEntityLayoutExecutor.FindFooterSkipSpans(sheet, region,
                    region.Start.Row + (region.IncludeHeader ? 1 : 0), lastRow, cancellationToken))
                    AddRange(region.SheetName, span.StartRow, span.EndRow);
            }
        }

        return rows.ToDictionary(pair => pair.Key,
            pair => (IReadOnlyList<int>)Array.AsReadOnly(pair.Value.ToArray()),
            StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 解析实体失败输出所需的模板命名锚点。
    /// </summary>
    /// <param name="workbook">已打开的实体工作簿。</param>
    /// <param name="sheetName">布局声明的工作表名称。</param>
    /// <param name="anchorName">命名锚点名称。</param>
    /// <returns>命名锚点的零基坐标。</returns>
    private static ExcelEntityCellReference ResolveNamedAnchor(XLWorkbook workbook, string sheetName,
        string anchorName)
    {
        var sheet = workbook.Worksheets.FirstOrDefault(item => string.Equals(item.Name, sheetName,
            StringComparison.OrdinalIgnoreCase));
        if (sheet == null || !ClosedXmlWorkbookValidationPipeline.TryResolveNamedRange(sheet, anchorName,
                out var range) || range == null)
            throw new BingOfficesConfigurationException(
                $"模板缺少或无法解析命名锚点: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        var first = range.RangeAddress.FirstAddress;
        var last = range.RangeAddress.LastAddress;
        if (first.RowNumber != last.RowNumber || first.ColumnNumber != last.ColumnNumber
            || !string.Equals(range.Worksheet.Name, sheet.Name, StringComparison.OrdinalIgnoreCase))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用当前 Sheet 的单个单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        return ExcelEntityCellReference.Parse(first.ToString());
    }

    /// <summary>
    /// 校验并提取实体导入选项中的资源限制。
    /// </summary>
    /// <param name="options">实体导入选项。</param>
    /// <returns>已验证的资源限制。</returns>
    private static ExcelResourceLimits ValidateEntityOptions(ExcelEntityImportOptions options)
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        options.ResourceLimits.Validate();
        return options.ResourceLimits;
    }

    /// <summary>
    /// 按单个 Sheet 请求读取表头、数据行和动态列。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="root">Workbook 根模型。</param>
    /// <param name="worksheet">来源工作表。</param>
    /// <param name="sheetRequest">Sheet 导入请求。</param>
    /// <param name="workbookRequest">Workbook 导入请求。</param>
    /// <param name="totalRows">已处理数据行总数。</param>
    /// <param name="workbookErrors">Workbook 级错误集合。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>该 Sheet 的导入结果。</returns>
    private ExcelSheetImportResult ImportSheet<TWorkbook>(TWorkbook root, IXLWorksheet worksheet,
        ExcelSheetImportRequest sheetRequest, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        ref int totalRows, ICollection<ExcelImportError> workbookErrors, bool isDate1904,
        CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var plan = _planBuilder.Create(sheetRequest.ItemType,
            sheetRequest.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
            sheetRequest.MappingConfiguration, MappingDirection.Import);
        var fixedColumns = plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn).ToArray();
        var dynamicColumns = plan.DynamicColumns.ToArray();
        var headerRow = sheetRequest.HeaderRowIndex + 1;
        var dataRow = sheetRequest.DataRowStartIndex + 1;
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;
        var headers = new List<string>();
        for (var column = 1; column <= lastColumn && column <= sheetRequest.MaxReadColumns; column++)
            headers.Add(NormalizeHeader(ReadCell(worksheet.Cell(headerRow, column)), sheetRequest.HeaderWhitespace));
        ValidateUniqueHeaders(headers, sheetRequest.HeaderWhitespace, sheetRequest.HeaderComparison,
            worksheet.Name, headerRow);
        var bindings = new List<FixedBinding>();
        foreach (var column in fixedColumns)
        {
            var property = sheetRequest.ItemType.GetProperty(column.Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (property == null || !property.CanWrite)
                continue;
            var index = FindHeader(headers, column.Title ?? column.Name, column.Aliases,
                sheetRequest.HeaderWhitespace, sheetRequest.HeaderComparison);
            if (index < 0)
            {
                if (sheetRequest.RequireExpectedHeaders)
                    AddError(workbookErrors, workbookRequest, new ExcelImportError(
                        ExcelImportErrorCode.InvalidHeader, $"缺少表头: {column.Title ?? column.Name}",
                        worksheet.Name, headerRow, 0, column.Name));
                continue;
            }
            bindings.Add(new FixedBinding(column, property, index + 1));
        }
        var dynamicBindings = dynamicColumns.Select(column => new DynamicBinding(column,
            FindHeader(headers, column.Title, column.Aliases, sheetRequest.HeaderWhitespace,
                sheetRequest.HeaderComparison) + 1))
            .Where(binding => binding.ColumnIndex > 0).ToArray();
        var target = sheetRequest.Target(root);
        if (target == null)
            throw new BingOfficesConfigurationException($"Sheet {sheetRequest.Name} 的目标集合为空。",
                stage: BingOfficesStage.Plan);
        var sourceRows = new List<int>();
        var sheetErrors = new List<ExcelImportError>();
        var limits = workbookRequest.ResourceLimits;
        var validationEnabled = workbookRequest.ValidationMode != ExcelImportValidationMode.Disabled;
        var tracksUnique = validationEnabled
            && (bindings.Any(binding => binding.Column.IsUnique)
                || dynamicBindings.Any(binding => binding.Column.IsUnique));
        var duplicateValues = tracksUnique
            ? new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase)
            : null;
        var unique = tracksUnique
            ? new UniqueTracker(duplicateValues, limits?.MaxTrackedUniqueValues,
                CreateStringComparer(limits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase))
            : null;
        var workbookValidationEnabled = workbookRequest.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || workbookRequest.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook;
        for (var row = dataRow; row <= lastRow; row++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (workbookRequest.ResourceLimits?.MaxRows is int maximum && totalRows >= maximum)
            {
                AddError(workbookErrors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.ResourceLimit,
                    $"XLSX 数据行数超过限制: {maximum}。", worksheet.Name, row, 0,
                    string.Empty));
                break;
            }
            totalRows++;
            var item = Activator.CreateInstance(sheetRequest.ItemType);
            var valid = true;
            if (unique != null)
                unique.BeginRow();
            var dynamicValues = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            foreach (var binding in bindings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var cell = worksheet.Cell(row, binding.ColumnIndex);
                var raw = ReadCell(cell, binding.Property.PropertyType);
                if (workbookValidationEnabled)
                {
                    var workbookValidation = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, cell, raw,
                        ClosedXmlValueAdapter.ToText(raw, sheetRequest.Culture ?? CultureInfo.InvariantCulture),
                        sheetRequest.Culture ?? CultureInfo.InvariantCulture, isDate1904, cancellationToken);
                    if (!workbookValidation.IsValid)
                    {
                        AddError(workbookErrors, workbookRequest, new ExcelImportError(
                            ExcelImportErrorCode.WorkbookValidation, workbookValidation.Message,
                            worksheet.Name, row, binding.ColumnIndex, binding.Column.Name,
                            rawValue: raw));
                        var reportUnsupported = workbookValidation.IsUnsupported
                            && workbookRequest.UnsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report;
                        if (!reportUnsupported)
                            valid = false;
                        if (!reportUnsupported && (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                            || !workbookValidation.IsUnsupported
                            || workbookRequest.UnsupportedFeaturePolicy != ExcelUnsupportedFeaturePolicy.Report))
                            break;
                    }
                }
                try
                {
                    var converted = ClosedXmlValueAdapter.ConvertFrom(raw, binding.Column,
                        binding.Property, worksheet.Name, row, binding.ColumnIndex,
                        sheetRequest.Culture ?? CultureInfo.InvariantCulture, isDate1904,
                        ClosedXmlValueAdapter.CreateCell(raw, ClosedXmlValueAdapter.ToText(raw,
                            sheetRequest.Culture ?? CultureInfo.InvariantCulture), isDate1904));
                    if (!Validate(binding.Column, raw, converted, worksheet.Name, row, binding.ColumnIndex,
                            binding.Property, sheetRequest, workbookRequest, unique, workbookErrors, isDate1904))
                    {
                        valid = false;
                        if (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                        continue;
                    }
                    binding.Property.SetValue(item, converted);
                }
                catch (Exception exception) when (exception is not OperationCanceledException)
                {
                    valid = false;
                    AddError(workbookErrors, workbookRequest, new ExcelImportError(
                        ExcelImportErrorCode.ValueConversion, exception.Message, worksheet.Name, row,
                        binding.ColumnIndex, binding.Column.Name, rawValue: raw));
                    if (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                        break;
                }
            }
            foreach (var binding in dynamicBindings)
            {
                if (!valid)
                    break;
                var cell = worksheet.Cell(row, binding.ColumnIndex);
                var raw = ReadCell(cell);
                if (workbookValidationEnabled)
                {
                    var workbookValidation = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, cell, raw,
                        ClosedXmlValueAdapter.ToText(raw, sheetRequest.Culture ?? CultureInfo.InvariantCulture),
                        sheetRequest.Culture ?? CultureInfo.InvariantCulture, isDate1904, cancellationToken);
                    if (!workbookValidation.IsValid)
                    {
                        AddError(workbookErrors, workbookRequest, new ExcelImportError(
                            ExcelImportErrorCode.WorkbookValidation, workbookValidation.Message,
                            worksheet.Name, row, binding.ColumnIndex, binding.Column.Key,
                            rawValue: raw));
                        var reportUnsupported = workbookValidation.IsUnsupported
                            && workbookRequest.UnsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report;
                        if (!reportUnsupported)
                            valid = false;
                        if (!reportUnsupported && (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure
                            || !workbookValidation.IsUnsupported
                            || workbookRequest.UnsupportedFeaturePolicy != ExcelUnsupportedFeaturePolicy.Report))
                            break;
                    }
                }
                try
                {
                    dynamicValues[binding.Column.Key] = ClosedXmlValueAdapter.ConvertDynamicFrom(raw,
                        binding.Column, worksheet.Name, row, binding.ColumnIndex,
                        sheetRequest.Culture ?? CultureInfo.InvariantCulture, isDate1904,
                        ClosedXmlValueAdapter.CreateCell(raw, ClosedXmlValueAdapter.ToText(raw,
                            sheetRequest.Culture ?? CultureInfo.InvariantCulture), isDate1904));
                    var converted = dynamicValues[binding.Column.Key];
                    if (!ValidateDynamic(binding.Column, raw, converted, worksheet.Name, row,
                            binding.ColumnIndex, sheetRequest, workbookRequest, unique, workbookErrors,
                            isDate1904))
                    {
                        valid = false;
                        if (sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                            break;
                    }
                }
                catch (Exception exception) when (exception is not OperationCanceledException)
                {
                    valid = false;
                    AddError(workbookErrors, workbookRequest, new ExcelImportError(
                        ExcelImportErrorCode.ValueConversion, exception.Message, worksheet.Name, row,
                        binding.ColumnIndex, binding.Column.Key, rawValue: raw));
                }
            }
            if (valid)
            {
                SetDynamicValues(item, sheetRequest, dynamicValues);
                AddItem(target, item);
                sourceRows.Add(row - 1);
            }
            if (unique != null)
            {
                if (valid)
                    unique.CommitRow();
                else
                    unique.RollbackRow();
            }
            if (!valid && sheetRequest.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                continue;
        }
        return new ExcelSheetImportResult(worksheet.Name, sheetRequest.ItemType, sourceRows,
            sheetErrors.Count == 0 ? workbookErrors.Where(error => string.Equals(error.SheetName,
                worksheet.Name, StringComparison.OrdinalIgnoreCase)).ToArray() : sheetErrors.ToArray());
    }

    /// <summary>
    /// 执行固定映射列的校验和唯一值检查。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="column">固定映射列。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="sheet">工作表名称。</param>
    /// <param name="row">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="property">目标属性。</param>
    /// <param name="sheetRequest">Sheet 导入请求。</param>
    /// <param name="workbookRequest">Workbook 导入请求。</param>
    /// <param name="unique">唯一值跟踪器。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <returns>校验通过时返回 <see langword="true" />。</returns>
    private static bool Validate<TWorkbook>(IExcelMappingColumn column, object raw, object converted, string sheet,
        int row, int columnIndex, PropertyInfo property, ExcelSheetImportRequest sheetRequest,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, UniqueTracker unique,
        ICollection<ExcelImportError> errors, bool isDate1904)
        where TWorkbook : class, new()
    {
        if (workbookRequest.ValidationMode == ExcelImportValidationMode.Disabled)
            return true;
        var culture = sheetRequest.Culture ?? CultureInfo.InvariantCulture;
        var text = ClosedXmlValueAdapter.ToText(raw, culture);
        var cell = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
        foreach (var binding in column.ValidationBindings ?? Array.Empty<IExcelValidationBinding>())
        {
            var value = binding.IsRaw ? raw : converted;
            try
            {
                if (binding.Validate(new ExcelValidationContext(text, sheet, row, columnIndex,
                    column.Name, value, property.PropertyType, cell, culture)))
                    continue;
                AddError(errors, workbookRequest, new ExcelImportError(
                    ClosedXmlValueAdapter.GetValidationCode(binding), binding.ErrorMessage,
                    sheet, row, columnIndex, column.Name, rawValue: raw));
                return false;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                AddError(errors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.Validation, exception.Message, sheet, row, columnIndex,
                    column.Name, rawValue: raw));
                return false;
            }
        }
        if (!column.IsUnique || (column.UniqueIgnoreEmpty && string.IsNullOrWhiteSpace(text)))
            return true;
        try
        {
            if (unique.TryReserve(column.Name, text, false, column.UniqueIgnoreEmpty, row))
                return true;
            AddError(errors, workbookRequest, new ExcelImportError(
                ExcelImportErrorCode.Validation, "重复数据。", sheet, row, columnIndex,
                column.Name, rawValue: raw));
            return false;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            AddError(errors, workbookRequest, new ExcelImportError(
                ExcelImportErrorCode.ResourceLimit, exception.Message, sheet, row, columnIndex,
                column.Name, rawValue: raw));
            return false;
        }
    }

    /// <summary>
    /// 执行动态映射列的校验和唯一值检查。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="column">动态映射列。</param>
    /// <param name="raw">单元格原始值。</param>
    /// <param name="converted">转换后的值。</param>
    /// <param name="sheet">工作表名称。</param>
    /// <param name="row">工作表行号。</param>
    /// <param name="columnIndex">工作表列号。</param>
    /// <param name="sheetRequest">Sheet 导入请求。</param>
    /// <param name="workbookRequest">Workbook 导入请求。</param>
    /// <param name="unique">唯一值跟踪器。</param>
    /// <param name="errors">错误集合。</param>
    /// <param name="isDate1904">是否使用 1904 日期系统。</param>
    /// <returns>校验通过时返回 <see langword="true" />。</returns>
    private static bool ValidateDynamic<TWorkbook>(IExcelDynamicMappingColumn column, object raw,
        object converted, string sheet, int row, int columnIndex, ExcelSheetImportRequest sheetRequest,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, UniqueTracker unique,
        ICollection<ExcelImportError> errors, bool isDate1904)
        where TWorkbook : class, new()
    {
        if (workbookRequest.ValidationMode == ExcelImportValidationMode.Disabled)
            return true;
        var culture = sheetRequest.Culture ?? CultureInfo.InvariantCulture;
        var text = ClosedXmlValueAdapter.ToText(raw, culture);
        var cell = ClosedXmlValueAdapter.CreateCell(raw, text, isDate1904);
        var propertyType = ClosedXmlValueAdapter.ResolveDynamicType(column.DataTypeName);
        foreach (var binding in column.ValidationBindings ?? Array.Empty<IExcelValidationBinding>())
        {
            var value = binding.IsRaw ? raw : converted;
            try
            {
                if (binding.Validate(new ExcelValidationContext(text, sheet, row, columnIndex,
                    column.Key, value, propertyType, cell, culture)))
                    continue;
                AddError(errors, workbookRequest, new ExcelImportError(
                    ClosedXmlValueAdapter.GetValidationCode(binding), binding.ErrorMessage,
                    sheet, row, columnIndex, column.Key, rawValue: raw));
                return false;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OutOfMemoryException
                && exception is not StackOverflowException)
            {
                AddError(errors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.Validation, exception.Message, sheet, row, columnIndex,
                    column.Key, rawValue: raw));
                return false;
            }
        }
        if (!column.IsUnique || (column.UniqueIgnoreEmpty && string.IsNullOrWhiteSpace(text)))
            return true;
        try
        {
            if (unique.TryReserve(column.Key, text, false, column.UniqueIgnoreEmpty, row))
                return true;
            AddError(errors, workbookRequest, new ExcelImportError(
                ExcelImportErrorCode.Validation, "重复数据。", sheet, row, columnIndex,
                column.Key, rawValue: raw));
            return false;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            AddError(errors, workbookRequest, new ExcelImportError(
                ExcelImportErrorCode.ResourceLimit, exception.Message, sheet, row, columnIndex,
                column.Key, rawValue: raw));
            return false;
        }
    }

    /// <summary>
    /// 将动态列值写入请求指定的动态目标集合。
    /// </summary>
    /// <param name="item">当前数据项目。</param>
    /// <param name="request">Sheet 导入请求。</param>
    /// <param name="values">动态列值。</param>
    private static void SetDynamicValues(object item, ExcelSheetImportRequest request,
        IDictionary<string, object> values)
    {
        if (request.DynamicTarget == null || values.Count == 0)
            return;
        var member = (request.DynamicTarget as LambdaExpression)?.Body as MemberExpression;
        if (member?.Member is PropertyInfo property && property.CanWrite)
        {
            var current = property.GetValue(item) as IDictionary<string, object>;
            if (current == null)
            {
                current = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                property.SetValue(item, current);
            }
            foreach (var pair in values)
                current[pair.Key] = pair.Value;
        }
        else
        {
            var current = request.DynamicTargetGetter?.Invoke(item) as IDictionary<string, object>;
            if (current != null)
                foreach (var pair in values)
                    current[pair.Key] = pair.Value;
        }
    }

    /// <summary>
    /// 将项目加入导入目标集合。
    /// </summary>
    /// <param name="target">目标集合。</param>
    /// <param name="item">待加入项目。</param>
    private static void AddItem(object target, object item)
    {
        if (target is IList list)
        {
            list.Add(item);
            return;
        }
        var method = target.GetType().GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);
        if (method == null)
            throw new BingOfficesConfigurationException("导入目标集合不支持 Add。", stage: BingOfficesStage.Plan);
        method.Invoke(target, new[] { item });
    }

    /// <summary>
    /// 从单元格读取原始值，并按目标类型保留公式文本。
    /// </summary>
    /// <param name="cell">来源单元格。</param>
    /// <param name="targetType">目标类型；字符串类型保留公式文本。</param>
    /// <returns>单元格原始值；空单元格返回 <see langword="null" />。</returns>
    private static object ReadCell(IXLCell cell, Type targetType = null)
    {
        var hasFormula = !string.IsNullOrWhiteSpace(cell.FormulaA1);
        if (hasFormula && targetType == typeof(string))
            return "=" + cell.FormulaA1;

        // CachedValue is deliberately used for formula cells: reading Value may evaluate
        // the formula and would silently replace the public formula-preservation contract.
        var value = hasFormula ? cell.CachedValue : cell.Value;
        if (value.IsBlank)
            return null;
        if (value.IsBoolean)
            return value.GetBoolean();
        if (value.IsNumber)
            return value.GetNumber();
        if (value.IsDateTime)
            return value.GetDateTime();
        if (value.IsTimeSpan)
            return value.GetTimeSpan();
        if (value.IsError)
            return value.GetError().ToString();
        return value.GetText();
    }

    /// <summary>
    /// 按导入请求的名称或索引解析工作表。
    /// </summary>
    /// <param name="workbook">来源工作簿。</param>
    /// <param name="request">Sheet 导入请求。</param>
    /// <param name="comparison">工作表名称比较策略。</param>
    /// <returns>匹配的工作表；未找到时返回 <see langword="null" />。</returns>
    private static IXLWorksheet ResolveWorksheet(XLWorkbook workbook,
        ExcelSheetImportRequest request, ExcelNameComparison comparison)
    {
        if (request.Selector.Kind == ExcelSheetSelectorKind.ByIndex)
        {
            var index = request.Selector.Index.Value;
            return index >= 0 && index < workbook.Worksheets.Count ? workbook.Worksheets.Worksheet(index + 1) : null;
        }
        var comparer = comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        return workbook.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, request.Selector.Name,
            comparer));
    }

    /// <summary>
    /// 在规范化表头中查找标题或别名。
    /// </summary>
    /// <param name="headers">规范化后的表头。</param>
    /// <param name="title">目标标题。</param>
    /// <param name="aliases">目标别名。</param>
    /// <param name="whitespace">空白处理策略。</param>
    /// <param name="comparison">名称比较策略。</param>
    /// <returns>匹配的零基索引；未找到时返回 -1。</returns>
    private static int FindHeader(IReadOnlyList<string> headers, string title,
        IReadOnlyList<string> aliases, ExcelWhitespacePolicy whitespace,
        ExcelNameComparison comparison)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            if (HeaderEquals(headers[index], title, whitespace, comparison))
                return index;
            if (aliases != null && aliases.Any(alias =>
                    HeaderEquals(headers[index], alias, whitespace, comparison)))
                return index;
        }
        return -1;
    }

    /// <summary>
    /// 按表头空白策略和名称比较策略比较两个表头。
    /// </summary>
    /// <param name="left">左侧表头。</param>
    /// <param name="right">右侧表头。</param>
    /// <param name="whitespace">空白处理策略。</param>
    /// <param name="comparison">名称比较策略。</param>
    /// <returns>规范化后相等时返回 <see langword="true" />。</returns>
    private static bool HeaderEquals(string left, string right, ExcelWhitespacePolicy whitespace,
        ExcelNameComparison comparison) =>
        string.Equals(NormalizeHeaderKey(left, whitespace), NormalizeHeaderKey(right, whitespace),
            comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 检查规范化后的非空表头是否重复。
    /// </summary>
    /// <param name="headers">规范化后的表头。</param>
    /// <param name="whitespace">空白处理策略。</param>
    /// <param name="comparison">名称比较策略。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <param name="row">表头行号。</param>
    private static void ValidateUniqueHeaders(IReadOnlyList<string> headers,
        ExcelWhitespacePolicy whitespace, ExcelNameComparison comparison, string sheetName, int row)
    {
        var comparer = comparison == ExcelNameComparison.Ordinal
            ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var seen = new HashSet<string>(comparer);
        for (var index = 0; index < headers.Count; index++)
        {
            var header = headers[index];
            if (string.IsNullOrWhiteSpace(header))
                continue;
            var key = NormalizeHeaderKey(header, whitespace);
            if (!seen.Add(key))
                throw new BingOfficesConfigurationException(
                    $"工作表 {sheetName} 的表头存在重复列: {header}", stage: BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 按空白策略规范化表头键。
    /// </summary>
    /// <param name="value">表头文本。</param>
    /// <param name="policy">空白处理策略。</param>
    /// <returns>规范化后的表头键。</returns>
    private static string NormalizeHeaderKey(string value, ExcelWhitespacePolicy policy)
    {
        if (policy == ExcelWhitespacePolicy.Preserve)
            return value ?? string.Empty;
        if (policy == ExcelWhitespacePolicy.RemoveAll)
            return new string((value ?? string.Empty).Where(character => !char.IsWhiteSpace(character)).ToArray());
        return (value ?? string.Empty).Trim();
    }

    /// <summary>
    /// 将单元格值转换为规范化表头文本。
    /// </summary>
    /// <param name="value">表头单元格值。</param>
    /// <param name="policy">空白处理策略。</param>
    /// <returns>规范化后的表头文本。</returns>
    private static string NormalizeHeader(object value, ExcelWhitespacePolicy policy)
    {
        var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        return NormalizeHeaderKey(text, policy);
    }

    /// <summary>
    /// 同步缓冲输入流并在读取过程中应用字节上限。
    /// </summary>
    /// <param name="source">输入流。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>定位到起始位置的内存流。</returns>
    private static MemoryStream BufferSource(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        limits?.Validate();
        var maximum = limits?.MaxInputBytes;
        if (source.CanSeek && maximum.HasValue && source.Length - source.Position > maximum.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。", provider: Provider,
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
        var buffer = new MemoryStream();
        var bytes = new byte[81920];
        long total = 0;
        int read;
        while ((read = source.Read(bytes, 0, bytes.Length)) != 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            total += read;
            if (maximum.HasValue && total > maximum.Value)
                throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。", provider: Provider,
                    operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
            buffer.Write(bytes, 0, read);
        }
        buffer.Position = 0;
        return buffer;
    }

    /// <summary>
    /// 异步缓冲输入流并在读取过程中应用字节上限。
    /// </summary>
    /// <param name="source">输入流。</param>
    /// <param name="limits">资源限制。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>异步返回定位到起始位置的内存流。</returns>
    private static async Task<MemoryStream> BufferSourceAsync(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        limits?.Validate();
        var maximum = limits?.MaxInputBytes;
        if (source.CanSeek && maximum.HasValue && source.Length - source.Position > maximum.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。", provider: Provider,
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);

        var buffer = new MemoryStream();
        var bytes = new byte[81920];
        long total = 0;
        try
        {
            while (true)
            {
                var requested = bytes.Length;
                if (maximum.HasValue)
                {
                    var remaining = maximum.Value - total;
                    if (remaining < 0)
                        throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。", provider: Provider,
                            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
                    requested = (int)Math.Min(requested, remaining + 1);
                }

                var read = await source.ReadAsync(bytes.AsMemory(0, requested), cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0)
                    break;

                total += read;
                if (maximum.HasValue && total > maximum.Value)
                    throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。", provider: Provider,
                        operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
                await buffer.WriteAsync(bytes.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
            }

            buffer.Position = 0;
            return buffer;
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 校验 Workbook 导入请求的 ClosedXML 支持边界。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="request">Workbook 导入请求。</param>
    private static void ValidateRequest<TWorkbook>(ExcelWorkbookImportRequest<TWorkbook> request)
        where TWorkbook : class, new()
    {
        // Workbook 原生校验在导入循环内按 Fail/Report 策略处理；其他请求校验由公共构建器完成。
    }

    /// <summary>
    /// 在创建 ClosedXML DOM 前校验失败工作簿需要无损保留的输入部件。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="buffered">已缓冲且可定位的 XLSX 流。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void ValidateFailureWorkbookInput<TWorkbook>(Stream buffered,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        if (request.FailureOptions == null
            || request.FailureOptions.Mode == ExcelImportFailureWorkbookMode.None)
            return;
        ClosedXmlTemplatePreflight.Validate(buffered, BingOfficesOperation.Import,
            cancellationToken, request.ResourceLimits);
        buffered.Position = 0;
    }

    /// <summary>
    /// 在未达到错误上限时记录 Workbook 导入错误。
    /// </summary>
    /// <typeparam name="TWorkbook">Workbook 根模型类型。</typeparam>
    /// <param name="errors">错误集合。</param>
    /// <param name="request">Workbook 导入请求。</param>
    /// <param name="error">待记录错误。</param>
    private static void AddError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        if (request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum)
            return;
        errors.Add(error);
    }

    /// <summary>
    /// 将公共字符串比较选项转换为 .NET 比较器。
    /// </summary>
    /// <param name="comparison">字符串比较选项。</param>
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

    /// <summary>
    /// 保存固定映射列、目标属性和实际列号。
    /// </summary>
    /// <param name="Column">固定映射列。</param>
    /// <param name="Property">目标属性。</param>
    /// <param name="ColumnIndex">实际工作表列号。</param>
    private sealed record FixedBinding(IExcelMappingColumn Column, PropertyInfo Property, int ColumnIndex);

    /// <summary>
    /// 保存动态映射列和实际列号。
    /// </summary>
    /// <param name="Column">动态映射列。</param>
    /// <param name="ColumnIndex">实际工作表列号。</param>
    private sealed record DynamicBinding(IExcelDynamicMappingColumn Column, int ColumnIndex);
}
