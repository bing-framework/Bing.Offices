using System.Reflection;
using System.Collections.Concurrent;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;

namespace Bing.Offices.Imports;

/// <summary>
/// 基于 NPOI 的 Excel 导入器。
/// </summary>
/// <remarks>
/// 输入会先复制，并由 NPOI 建立内存中的 Workbook DOM。
/// </remarks>
public sealed class NpoiExcelImporter : IExcelImporter, IExcelEntityImporter,
    IExcelEntityResourceImporter, IExcelProviderFeatureDescriptor
{
    /// <inheritdoc />
    public string ProviderName => "NPOI";

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Entity
        | ExcelProviderCapabilities.Template | ExcelProviderCapabilities.Merge
        | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xls
        | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = new[] { ExcelFormat.Xls, ExcelFormat.Xlsx };
    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> WriteFormats { get; } = Array.Empty<ExcelFormat>();
    /// <inheritdoc />
    public bool SupportsCompleteWorkbookImport => true;
    /// <inheritdoc />
    public bool SupportsBatchImport => false;
    /// <inheritdoc />
    public bool SupportsCompleteWorkbookExport => false;
    /// <inheritdoc />
    public bool SupportsTrueAsyncIo => true;
    /// <inheritdoc />
    public IReadOnlyList<string> Limitations { get; } = new[]
    {
        "XLSM、XLSB 和 ODS 在当前 NPOI Provider 中明确拒绝。"
    };
    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.TemplateEditing
        | ExcelProviderFeatures.WorkbookEditing | ExcelProviderFeatures.FormulaText
        | ExcelProviderFeatures.FormulaCachedValues;
    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;
    /// <summary>
    /// 调用指定工作簿根类型的泛型工作表导入逻辑。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="target">执行工作表导入的导入器。</param>
    /// <param name="sheet">待读取的 NPOI 工作表。</param>
    /// <param name="request">当前工作表导入请求。</param>
    /// <param name="root">接收导入实体的工作簿根对象。</param>
    /// <param name="sheetResults">接收工作表导入结果的集合。</param>
    /// <param name="errors">接收导入错误的收集器。</param>
    /// <param name="validationMode">当前导入验证模式。</param>
    /// <param name="resourceLimits">当前导入资源限制。</param>
    /// <param name="unsupportedFeaturePolicy">不支持功能的处理策略。</param>
    /// <param name="sourceLocations">实体到源工作表位置的映射。</param>
    /// <param name="runtime">当前导入运行时状态。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">导入过程中检查的取消令牌。</param>
    /// <param name="mappingPlan">当前工作表的映射计划。</param>
    private delegate void ImportSheetInvoker<TWorkbook>(NpoiExcelImporter target, ISheet sheet,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> sheetResults,
        ExcelImportErrorCollector errors, ExcelImportValidationMode validationMode,
        ExcelResourceLimits resourceLimits, ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations, ExcelImportRuntime runtime, bool isDate1904,
        CancellationToken cancellationToken, IExcelMappingPlan mappingPlan) where TWorkbook : class, new();

    /// <summary>
    /// 缓存指定工作簿根类型的工作表导入委托。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    private static class ImportSheetInvokerCache<TWorkbook> where TWorkbook : class, new()
    {
        /// <summary>
        /// 按工作表实体运行时类型缓存当前工作簿根类型的导入委托。
        /// </summary>
        internal static readonly ConcurrentDictionary<Type, ImportSheetInvoker<TWorkbook>> Invokers = new();
    }
    /// <summary>
    /// 当前导入器使用的校验规则。
    /// </summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;

    /// <summary>
    /// 当前导入器使用的值转换器。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;

    /// <summary>
    /// 将请求级配置编译为提供程序无关映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 将 Workbook 导入请求展开为按工作表执行的导入计划生成器。
    /// </summary>
    private readonly NpoiImportPlanBuilder _planBuilder;

    /// <summary>
    /// 当前导入器使用的命名配置校验规则。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>
    /// 按行校验并物化导入实体的执行器。
    /// </summary>
    private readonly NpoiImportRowMaterializer _rowMaterializer;
    /// <summary>
    /// 执行单个工作表的表头、资源、校验和逐行导入。
    /// </summary>
    private readonly NpoiImportSheetExecutor _sheetExecutor;
    /// <summary>
    /// 观察并记录公共 Excel 导入异常。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;
    /// <summary>
    /// 保存负责失败工作簿外围异步输出的可替换 staging 策略。
    /// </summary>
    private readonly INpoiAsyncStagingFactory _asyncStagingFactory;
    /// <summary>
    /// 单个实体布局执行器。
    /// </summary>
    private readonly NpoiEntityImportExecutor _entityExecutor;

    /// <summary>
    /// 初始化一个 <see cref="NpoiExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    public NpoiExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
        : this(validationRules, valueConverters, namedValidationRules, mappingPlanFactory, exceptionObservers,
            new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.TempFile))
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="NpoiExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名配置校验规则集合。</param>
    /// <param name="mappingPlanFactory">方向化映射计划工厂。</param>
    /// <param name="exceptionObservers">接收公共运行异常的观察器集合。</param>
    /// <param name="asyncStagingFactory">失败工作簿外围异步输出的 staging 工厂。</param>
    internal NpoiExcelImporter(IEnumerable<IExcelValidationRule> validationRules,
        IEnumerable<IExcelValueConverter> valueConverters,
        IEnumerable<INamedExcelValidationRule> namedValidationRules,
        IExcelMappingPlanFactory mappingPlanFactory,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers,
        INpoiAsyncStagingFactory asyncStagingFactory)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? NpoiMappingPlanFactoryResolver.CreateDefault(
            _valueConverters, _validationRules, _namedValidationRules);
        _planBuilder = new NpoiImportPlanBuilder(_mappingPlanFactory);
        _rowMaterializer = new NpoiImportRowMaterializer();
        _sheetExecutor = new NpoiImportSheetExecutor(_rowMaterializer);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
        _asyncStagingFactory = asyncStagingFactory ?? throw new ArgumentNullException(nameof(asyncStagingFactory));
        _entityExecutor = new NpoiEntityImportExecutor(_mappingPlanFactory, _valueConverters);
    }

    /// <summary>
    /// 初始化一个 <see cref="NpoiExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="asyncStagingFactory">失败工作簿外围异步输出的 staging 工厂。</param>
    internal NpoiExcelImporter(INpoiAsyncStagingFactory asyncStagingFactory)
        : this(null, null, null, null, null, asyncStagingFactory)
    {
    }

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
        ValidateEntitySource(source, layout);
        var limits = ValidateEntityOptions(options);
        try
        {
            using var buffered = new MemoryStream();
            NpoiStreamCopier.Copy(source, buffered, cancellationToken, limits.MaxInputBytes);
            return ImportEntityBuffered(buffered, layout, false, limits, cancellationToken);
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
            var translated = new BingOfficesImportException("Excel 实体导入失败。", exception, "NPOI",
                BingOfficesStage.Read);
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
        ValidateEntitySource(source, layout);
        var limits = ValidateEntityOptions(options);
        try
        {
            using var buffered = new MemoryStream();
            await NpoiStreamCopier.CopyAsync(source, buffered, cancellationToken, limits.MaxInputBytes)
                .ConfigureAwait(false);
            return ImportEntityBuffered(buffered, layout, false, limits, cancellationToken);
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
            var translated = new BingOfficesImportException("Excel 实体异步导入失败。", exception, "NPOI",
                BingOfficesStage.Read);
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
        ValidateEntitySource(source, layout);
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var limits = ValidateEntityOptions(options);
        try
        {
            using var templateBuffer = new MemoryStream();
            NpoiStreamCopier.Copy(template.Template, templateBuffer, cancellationToken, limits.MaxInputBytes);
            using var templateWorkbook = OpenEntityWorkbook(templateBuffer, limits, cancellationToken);
            ValidateEntityTemplate(templateWorkbook, layout);
            using var buffered = new MemoryStream();
            NpoiStreamCopier.Copy(source, buffered, cancellationToken, limits.MaxInputBytes);
            return ImportEntityBuffered(buffered, layout, false, limits, cancellationToken);
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
        ValidateEntitySource(source, layout);
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        var limits = ValidateEntityOptions(options);
        try
        {
            using var templateBuffer = new MemoryStream();
            await NpoiStreamCopier.CopyAsync(template.Template, templateBuffer, cancellationToken,
                limits.MaxInputBytes).ConfigureAwait(false);
            using var templateWorkbook = OpenEntityWorkbook(templateBuffer, limits, cancellationToken);
            ValidateEntityTemplate(templateWorkbook, layout);
            using var buffered = new MemoryStream();
            await NpoiStreamCopier.CopyAsync(source, buffered, cancellationToken, limits.MaxInputBytes)
                .ConfigureAwait(false);
            return ImportEntityBuffered(buffered, layout, false, limits, cancellationToken);
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
            var translated = new BingOfficesImportException("Excel 模板实体异步导入失败。", exception, "NPOI",
                BingOfficesStage.Read);
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
    /// 从已缓冲的输入流创建工作簿并执行实体布局导入。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="buffered">已定位到内存中的工作簿流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="requireTemplateMerges">是否要求模板中的合并区域已存在。</param>
    /// <param name="limits">实体导入使用的资源限制。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <returns>实体导入结果。</returns>
    private ExcelEntityImportResult<TEntity> ImportEntityBuffered<TEntity>(MemoryStream buffered,
        ExcelEntityLayout<TEntity> layout, bool requireTemplateMerges, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        NpoiXlsxZipPreflight.Validate(buffered, limits, cancellationToken);
        buffered.Position = 0;
        using var workbook = WorkbookFactory.Create(buffered);
        return _entityExecutor.Read(workbook, new TEntity(), layout, requireTemplateMerges, limits,
            cancellationToken);
    }

    /// <summary>
    /// 对已缓冲的实体输入执行预检并打开 NPOI 工作簿。
    /// </summary>
    /// <param name="buffered">已缓冲的工作簿流。</param>
    /// <param name="limits">实体导入使用的资源限制。</param>
    /// <param name="cancellationToken">用于取消预检和打开操作的令牌。</param>
    /// <returns>已打开的 NPOI 工作簿。</returns>
    private static IWorkbook OpenEntityWorkbook(MemoryStream buffered, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        NpoiXlsxZipPreflight.Validate(buffered, limits, cancellationToken);
        buffered.Position = 0;
        return WorkbookFactory.Create(buffered);
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
    /// 校验实体模板包含布局声明的工作表和合并区域。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="workbook">已打开的模板工作簿。</param>
    /// <param name="layout">实体布局。</param>
    private static void ValidateEntityTemplate<TEntity>(IWorkbook workbook, ExcelEntityLayout<TEntity> layout)
        where TEntity : class, new()
    {
        foreach (var name in layout.Cells.Select(item => item.SheetName)
                     .Concat(layout.ListRegions.Select(item => item.SheetName))
                     .Concat(layout.Merges.Select(item => item.SheetName))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var sheet = workbook.GetSheet(name);
            if (sheet == null)
                throw new BingOfficesConfigurationException($"模板缺少请求的 Sheet: {name}", stage: BingOfficesStage.Plan);
            NpoiEntityLayoutSupport.PreflightMerges(sheet, layout.Merges, false, true);
        }
    }

    /// <summary>
    /// 校验实体导入源流和布局参数。
    /// </summary>
    /// <typeparam name="TEntity">实体类型。</typeparam>
    /// <param name="source">待读取的工作簿流。</param>
    /// <param name="layout">实体布局。</param>
    private static void ValidateEntitySource<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout)
        where TEntity : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
    }

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
            return ImportCore(source, request, cancellationToken);
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
        catch (BingOfficesResourceLimitException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (ImageResourceLimitException exception)
        {
            var translated = new BingOfficesResourceLimitException("Excel 导入超过图片资源限制。", exception,
                "NPOI", BingOfficesOperation.Import, BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("Excel 导入失败。", exception, "NPOI",
                BingOfficesStage.Read);
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
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();

        using var bufferedSource = new MemoryStream();
        INpoiAsyncStaging failureStaging = null;
        Exception primaryException = null;
        try
        {
            await NpoiStreamCopier.CopyAsync(source, bufferedSource, cancellationToken,
                request.ResourceLimits?.MaxInputBytes).ConfigureAwait(false);
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
        catch (BingOfficesException exception)
        {
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("Excel 输入流异步读取失败。", exception, "NPOI",
                BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }

        bufferedSource.Position = 0;
        try
        {
            var failureOptions = request.FailureOptions;
            var failureDestination = failureOptions?.Destination;
            if (failureOptions != null && failureOptions.Mode != ExcelImportFailureWorkbookMode.None
                && (failureDestination != null || failureOptions.DestinationPath != null))
                failureStaging = _asyncStagingFactory.Create("bing-offices-failure-async-");
            var result = ImportBufferedCore(bufferedSource, request, cancellationToken,
                failureStaging?.WriteStream);
            if (failureStaging != null && result.Errors.Count > 0)
            {
                await failureStaging.FlushAsync(cancellationToken).ConfigureAwait(false);
                if (failureOptions.DestinationPath != null)
                    await AtomicFileCommitter.CommitAsync(failureOptions.DestinationPath,
                        (destination, token) => failureStaging.CopyToAsync(destination, token),
                        cancellationToken, "FailureWorkbook").ConfigureAwait(false);
                else
                    await failureStaging.CopyToAsync(failureDestination, cancellationToken).ConfigureAwait(false);
            }
            return result;
        }
        catch (OperationCanceledException exception) when (cancellationToken.IsCancellationRequested
            && exception.GetType() != typeof(OperationCanceledException))
        {
            primaryException = new OperationCanceledException(cancellationToken);
            throw primaryException;
        }
        catch (OperationCanceledException exception)
        {
            primaryException = exception;
            throw;
        }
        catch (ArgumentException exception)
        {
            primaryException = exception;
            throw;
        }
        catch (BingOfficesResourceLimitException exception)
        {
            primaryException = exception;
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (BingOfficesException exception)
        {
            primaryException = exception;
            _exceptionDispatcher.Observe(exception);
            throw;
        }
        catch (ImageResourceLimitException exception)
        {
            var translated = new BingOfficesResourceLimitException("Excel 导入超过图片资源限制。", exception,
                "NPOI", BingOfficesOperation.Import, BingOfficesStage.Read);
            primaryException = translated;
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            var translated = new BingOfficesImportException("Excel 导入失败。", exception, "NPOI",
                BingOfficesStage.Read);
            primaryException = translated;
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
        finally
        {
            if (failureStaging != null)
            {
                try
                {
                    failureStaging.Dispose();
                }
                catch (Exception cleanupException) when (cleanupException is IOException
                    || cleanupException is UnauthorizedAccessException)
                {
                    if (primaryException != null)
                        primaryException.Data["Bing.Offices.ExcelAsync.FailureStagingCleanupException"] =
                            cleanupException;
                    else
                        throw new IOException("Excel 失败工作簿临时文件清理失败。", cleanupException);
                }
            }
        }
    }

    /// <summary>
    /// 将源流缓冲到内存后执行工作簿导入。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="source">待导入的源流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="cancellationToken">缓冲和导入过程中检查的取消令牌。</param>
    /// <param name="failureDestinationOverride">可选的失败工作簿输出流。</param>
    /// <returns>包含根实体、工作表结果和错误集合的导入结果。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportCore<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken,
        Stream failureDestinationOverride = null)
        where TWorkbook : class, new()
    {
        using var bufferedSource = new MemoryStream();
        NpoiStreamCopier.Copy(source, bufferedSource, cancellationToken, request.ResourceLimits?.MaxInputBytes);
        return ImportBufferedCore(bufferedSource, request, cancellationToken, failureDestinationOverride);
    }

    /// <summary>
    /// 从已缓冲源流创建 NPOI 工作簿并执行各工作表导入。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="bufferedSource">已定位到可读取内容的缓冲源流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="cancellationToken">导入过程中检查的取消令牌。</param>
    /// <param name="failureDestinationOverride">可选的失败工作簿输出流。</param>
    /// <returns>包含根实体、工作表结果和错误集合的导入结果。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportBufferedCore<TWorkbook>(Stream bufferedSource,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken,
        Stream failureDestinationOverride = null)
        where TWorkbook : class, new()
    {
        bufferedSource.Position = 0;
        NpoiXlsxZipPreflight.Validate(bufferedSource, request.ResourceLimits ?? new ExcelResourceLimits(),
            cancellationToken);
        bufferedSource.Position = 0;
        using var workbook = WorkbookFactory.Create(bufferedSource);
        var isDate1904 = workbook.IsDate1904();
        var root = new TWorkbook();
        var sheetResults = new List<ExcelSheetImportResult>();
        var errors = new ExcelImportErrorCollector(request.ResourceLimits?.MaxErrors);
        var sourceLocations = request.Relations.Count == 0
            ? null
            : new Dictionary<object, SourceLocation>(ReferenceObjectComparer.Instance);
        var runtime = new ExcelImportRuntime(request.ResourceLimits);
        var resolvedSheets = request.Sheets.Select(sheet => ResolveSheet(workbook, sheet,
            request.SheetNameComparison)).ToArray();
        var existingSheets = resolvedSheets.Where(sheet => sheet.Exists).ToArray();
        var duplicatePhysicalSheet = existingSheets.GroupBy(item => item.Index)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicatePhysicalSheet != null)
        {
            var physicalIndex = duplicatePhysicalSheet.Key;
            var physicalName = duplicatePhysicalSheet.First().Name;
            var selectors = string.Join(", ", duplicatePhysicalSheet.Select(item =>
                GetSelectorDescription(item.Request.Selector)));
            throw new ArgumentException(
                $"多个 Sheet selector 指向同一物理 Sheet: {selectors}; 实际 Sheet=#{physicalIndex} {physicalName}");
        }
        Dictionary<ExcelSheetImportRequest, IExcelMappingPlan> plans;
        try
        {
            plans = _planBuilder.Create(existingSheets);
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException("Excel 映射配置无效。", exception,
                BingOfficesStage.Plan);
        }
        var resolvedSheetRequests = resolvedSheets.Where(sheet => sheet.Exists &&
                !workbook.IsSheetHidden(sheet.Index) && !workbook.IsSheetVeryHidden(sheet.Index))
            .ToDictionary(sheet => sheet.Name, sheet => sheet.Request, StringComparer.OrdinalIgnoreCase);
        foreach (var resolvedSheet in resolvedSheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (errors.IsLimitReached)
                break;
            var sheetRequest = resolvedSheet.Request;
            if (!resolvedSheet.Exists)
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {GetSelectorDescription(sheetRequest.Selector)}",
                    GetSelectorDescription(sheetRequest.Selector), 0, 0, null));
                continue;
            }
            if (workbook.IsSheetHidden(resolvedSheet.Index) || workbook.IsSheetVeryHidden(resolvedSheet.Index))
            {
                errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"请求的 Sheet 被隐藏: {resolvedSheet.Name}", resolvedSheet.Name,
                    0, 0, null));
                continue;
            }
            var sheet = workbook.GetSheetAt(resolvedSheet.Index);
            try
            {
                ImportTypedSheet(sheet, sheetRequest, root, sheetResults, errors, request.ValidationMode,
                    request.ResourceLimits, request.UnsupportedFeaturePolicy, sourceLocations, runtime,
                    cancellationToken, plans[sheetRequest], isDate1904);
            }
            catch (NpoiSheetStructureException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    sheet.SheetName, sheetRequest.HeaderRowIndex + 1, 0, null);
                errors.Add(error);
                sheetResults.Add(new ExcelSheetImportResult(sheet.SheetName, sheetRequest.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
        }

        if (!runtime.RowLimitExceeded)
        {
            foreach (var relation in request.Relations)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                NpoiRelationBinder.Bind(root, relation, errors, sourceLocations, cancellationToken);
            }
        }
        NpoiFailureWorkbookWriter.Write(workbook, request.FailureOptions, errors.Errors, resolvedSheetRequests,
            cancellationToken, new SystemFailureWorkbookFileSystem(), failureDestinationOverride);
        return new ExcelWorkbookImportResult<TWorkbook>(runtime.RowLimitExceeded ? new TWorkbook() : root,
            sheetResults, errors.Errors,
            errors.IsTruncated, errors.MaxErrors);
    }

    /// <summary>
    /// 根据选择器查找工作表的零基物理索引。
    /// </summary>
    /// <param name="workbook">待查找的工作簿。</param>
    /// <param name="selector">按名称或索引定位工作表的选择器。</param>
    /// <param name="comparison">工作表名称比较规则。</param>
    /// <returns>匹配的零基工作表索引；未找到时为 -1。</returns>
    private static int ResolveSheetIndex(IWorkbook workbook, ExcelSheetSelector selector,
        ExcelNameComparison comparison)
    {
        if (selector.Kind == ExcelSheetSelectorKind.ByIndex)
            return selector.Index.Value < workbook.NumberOfSheets ? selector.Index.Value : -1;
        var comparer = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        for (var index = 0; index < workbook.NumberOfSheets; index++)
        {
            if (string.Equals(workbook.GetSheetName(index), selector.Name, comparer))
                return index;
        }
        return -1;
    }

    /// <summary>
    /// 解析一个工作表请求并保存物理名称，供后续所有导入阶段复用。
    /// </summary>
    /// <param name="workbook">待查找的工作簿。</param>
    /// <param name="request">待解析的工作表请求。</param>
    /// <param name="comparison">工作表名称比较规则。</param>
    /// <returns>已解析的工作表执行描述。</returns>
    private static NpoiResolvedSheet ResolveSheet(IWorkbook workbook, ExcelSheetImportRequest request,
        ExcelNameComparison comparison)
    {
        var index = ResolveSheetIndex(workbook, request.Selector, comparison);
        return new NpoiResolvedSheet(request, index, index < 0 ? null : workbook.GetSheetName(index));
    }

    /// <summary>
    /// 生成用于错误消息的工作表选择器描述。
    /// </summary>
    /// <param name="selector">待描述的工作表选择器。</param>
    /// <returns>索引选择器的井号形式或名称选择器的名称。</returns>
    private static string GetSelectorDescription(ExcelSheetSelector selector) =>
        selector.Kind == ExcelSheetSelectorKind.ByIndex ? $"#{selector.Index.Value}" : selector.Name;

    /// <summary>
    /// 通过一次类型擦除调用执行单个 Sheet 导入计划。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="sheet">待导入的 NPOI 工作表。</param>
    /// <param name="request">当前工作表导入请求。</param>
    /// <param name="root">接收导入实体的工作簿根对象。</param>
    /// <param name="sheetResults">接收工作表导入结果的集合。</param>
    /// <param name="errors">接收导入错误的收集器。</param>
    /// <param name="validationMode">当前导入验证模式。</param>
    /// <param name="resourceLimits">当前导入资源限制。</param>
    /// <param name="unsupportedFeaturePolicy">不支持功能的处理策略。</param>
    /// <param name="sourceLocations">实体到源工作表位置的映射。</param>
    /// <param name="runtime">当前导入运行时状态。</param>
    /// <param name="cancellationToken">导入过程中检查的取消令牌。</param>
    /// <param name="mappingPlan">当前工作表的映射计划。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    private void ImportTypedSheet<TWorkbook>(ISheet sheet, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        CancellationToken cancellationToken, IExcelMappingPlan mappingPlan, bool isDate1904)
        where TWorkbook : class, new()
    {
        ImportSheetInvokerCache<TWorkbook>.Invokers.GetOrAdd(request.ItemType,
            CreateImportSheetInvoker<TWorkbook>)(this, sheet, request, root, sheetResults, errors,
            validationMode, resourceLimits, unsupportedFeaturePolicy, sourceLocations, runtime,
            isDate1904, cancellationToken, mappingPlan);
    }

    /// <summary>
    /// 为运行时实体类型创建泛型工作表导入委托。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <param name="itemType">工作表实体的运行时类型。</param>
    /// <returns>绑定到指定实体类型的工作表导入委托。</returns>
    private static ImportSheetInvoker<TWorkbook> CreateImportSheetInvoker<TWorkbook>(Type itemType)
        where TWorkbook : class, new()
    {
        var method = typeof(NpoiExcelImporter).GetMethod(nameof(ImportTypedSheetCore),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(TWorkbook), itemType);
        return (ImportSheetInvoker<TWorkbook>)method.CreateDelegate(typeof(ImportSheetInvoker<TWorkbook>));
    }

    /// <summary>
    /// 导入一个具体实体类型，并将成功项写入 Workbook 根集合。
    /// </summary>
    /// <typeparam name="TWorkbook">工作簿根实体类型。</typeparam>
    /// <typeparam name="TItem">当前工作表的实体类型。</typeparam>
    /// <param name="sheet">待导入的 NPOI 工作表。</param>
    /// <param name="request">当前工作表导入请求。</param>
    /// <param name="root">接收导入实体的工作簿根对象。</param>
    /// <param name="sheetResults">接收工作表导入结果的集合。</param>
    /// <param name="errors">接收导入错误的收集器。</param>
    /// <param name="validationMode">当前导入验证模式。</param>
    /// <param name="resourceLimits">当前导入资源限制。</param>
    /// <param name="unsupportedFeaturePolicy">不支持功能的处理策略。</param>
    /// <param name="sourceLocations">实体到源工作表位置的映射。</param>
    /// <param name="runtime">当前导入运行时状态。</param>
    /// <param name="isDate1904">工作簿是否使用 1904 日期系统。</param>
    /// <param name="cancellationToken">导入过程中检查的取消令牌。</param>
    /// <param name="mappingPlan">当前工作表的映射计划。</param>
    private void ImportTypedSheetCore<TWorkbook, TItem>(ISheet sheet, ExcelSheetImportRequest request,
        TWorkbook root, ICollection<ExcelSheetImportResult> sheetResults, ExcelImportErrorCollector errors,
        ExcelImportValidationMode validationMode, ExcelResourceLimits resourceLimits,
        ExcelUnsupportedFeaturePolicy unsupportedFeaturePolicy,
        IDictionary<object, SourceLocation> sourceLocations,
        ExcelImportRuntime runtime,
        bool isDate1904, CancellationToken cancellationToken, IExcelMappingPlan mappingPlan)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var options = new ExcelImportExecutionOptions<TItem>
        {
            HeaderRowIndex = request.HeaderRowIndex,
            DataRowIndex = request.DataRowStartIndex,
            MaxReadColumns = request.MaxReadColumns,
            ReadColumnRange = request.ReadColumnRange,
            HeaderComparison = request.HeaderComparison,
            HeaderWhitespace = request.HeaderWhitespace,
            BodyWhitespace = request.BodyWhitespace,
            ValidationMode = validationMode,
            UnsupportedFeaturePolicy = unsupportedFeaturePolicy,
            DynamicTargetGetter = request.DynamicTargetGetter,
            RequireExpectedHeaders = request.RequireExpectedHeaders,
            ValidationFailureMode = request.ValidationFailureMode,
            Culture = request.Culture,
            DynamicColumns = request.DynamicColumns,
            FailOnUnknownDynamicColumns = request.FailOnUnknownDynamicColumns,
            ReportEmptyRows = request.ReportEmptyRows,
            StopAtFirstEmptyRow = request.StopAtFirstEmptyRow,
            MappingConfiguration = request.MappingConfiguration,
            MappingDocument = request.MappingDocument,
            MappingPlan = mappingPlan,
            MaxTrackedUniqueValues = resourceLimits?.MaxTrackedUniqueValues,
            UniqueComparison = resourceLimits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase
        };
        options.IsDate1904 = isDate1904;
        if (mappingPlan == null)
        {
            try
            {
                mappingPlan = _mappingPlanFactory.CreateWorkbook<TItem>(options.MappingDocument
                    ?? new ExcelMappingDocument { UseConventionFallback = true }, options.MappingConfiguration,
                    MappingDirection.Import, new[] { sheet.SheetName }).Sheets[0].Mapping;
                options.MappingPlan = mappingPlan;
            }
            catch (BingOfficesException)
            {
                throw;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                throw new BingOfficesConfigurationException("Excel 映射配置无效。", exception,
                    BingOfficesStage.Plan);
            }
        }
        var dynamicProperties = mappingPlan.Columns.Count(property => property.IsDynamicColumn);
        if (dynamicProperties > 1)
            throw new BingOfficesConfigurationException(
                $"导入模板 {typeof(TItem).FullName} 只能声明一个动态列属性。", stage: BingOfficesStage.Plan);
        var items = new List<TItem>();
        var rows = new List<int>();
        var sheetErrors = errors.CreateChild();
        _sheetExecutor.Execute(sheet, options, items, sheetErrors, runtime, cancellationToken, rows);
        var target = (ICollection<TItem>)request.Target(root);
        if (target == null)
            throw new BingOfficesConfigurationException($"Workbook 导入目标集合不可写入: {request.Name}",
                stage: BingOfficesStage.Plan);
        foreach (var item in items)
            target.Add(item);
        if (sourceLocations != null)
        {
            for (var index = 0; index < items.Count; index++)
                sourceLocations[items[index]] = new SourceLocation(sheet.SheetName, rows[index] + 1);
        }
        var result = new ExcelSheetImportResult(sheet.SheetName, typeof(TItem), rows,
            sheetErrors.Errors);
        sheetResults.Add(result);
    }

    /// <summary>
    /// 读取单元格的原始文本值，并优先读取公式的缓存结果。
    /// </summary>
    /// <param name="cell">待读取的单元格。</param>
    /// <returns>单元格文本；单元格为空时为空字符串。</returns>
    internal static string GetRawStringValue(ICell cell)
    {
        if (cell == null)
            return string.Empty;
        var cellType = cell.CellType == CellType.Formula ? cell.CachedFormulaResultType : cell.CellType;
        return cellType == CellType.String ? cell.StringCellValue ?? string.Empty : cell.GetStringValue();
    }

    /// <summary>
    /// 按照指定空白策略规范化单元格文本。
    /// </summary>
    /// <param name="value">待规范化的文本；为 null 时按空字符串处理。</param>
    /// <param name="policy">空白字符处理策略。</param>
    /// <returns>规范化后的文本。</returns>
    internal static string NormalizeText(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => RemoveWhitespace(value),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    /// <summary>
    /// 移除文本中的所有 Unicode 空白字符。
    /// </summary>
    /// <param name="value">待处理的文本。</param>
    /// <returns>移除空白后的文本；原文本无空白时直接返回原引用。</returns>
    private static string RemoveWhitespace(string value)
    {
        var hasWhitespace = value.Any(char.IsWhiteSpace);
        if (!hasWhitespace)
            return value;
        var buffer = new char[value.Length];
        var length = 0;
        foreach (var character in value)
        {
            if (!char.IsWhiteSpace(character))
                buffer[length++] = character;
        }
        return new string(buffer, 0, length);
    }

    /// <summary>
    /// 读取不依赖 NPOI 的单元格值描述。
    /// </summary>
    /// <param name="cell">待读取的单元格。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <returns>用于转换器和默认转换的单元格值描述。</returns>
    internal static ExcelCellValue ReadCellValue(ICell cell, bool isDate1904 = false)
    {
        if (cell == null)
            return new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty);

        var isFormula = cell.CellType == CellType.Formula;
        var effectiveType = isFormula ? cell.CachedFormulaResultType : cell.CellType;
        var effectiveKind = ResolveCellKind(effectiveType, cell);
        object value = effectiveType switch
        {
            CellType.Boolean => cell.BooleanCellValue,
            // 保留日期单元格的原始 serial，由 Core 统一应用 1900/1904 日期系统规则。
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => cell.NumericCellValue,
            CellType.Numeric => cell.NumericCellValue,
            CellType.Error => null,
            _ => GetRawStringValue(cell)
        };
        return new ExcelCellValue(value, GetRawStringValue(cell), isFormula ? ExcelCellKind.Formula : effectiveKind,
            isFormula ? effectiveKind : null, isFormula ? cell.CellFormula : null,
            effectiveType == CellType.Error ? cell.ErrorCellValue : null, cell.CellStyle?.DataFormat,
            isDate1904);
    }

    /// <summary>
    /// 映射 NPOI 单元格类型到提供程序无关的逻辑类型。
    /// </summary>
    /// <param name="cellType">NPOI 单元格类型。</param>
    /// <param name="cell">用于判断日期格式的单元格。</param>
    /// <returns>提供程序无关的单元格逻辑类型。</returns>
    private static ExcelCellKind ResolveCellKind(CellType cellType, ICell cell) => cellType switch
    {
        CellType.Blank => ExcelCellKind.Empty,
        CellType.String => ExcelCellKind.Text,
        CellType.Boolean => ExcelCellKind.Boolean,
        CellType.Error => ExcelCellKind.Error,
        CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => ExcelCellKind.DateTime,
        CellType.Numeric => ExcelCellKind.Number,
        _ => ExcelCellKind.Text
    };

    /// <summary>
    /// 创建与指定字符串比较规则等效的字符串比较器。
    /// </summary>
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
