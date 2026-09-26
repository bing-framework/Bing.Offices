using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Mappings;
using Bing.Offices.MiniExcel.Internals;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using MiniExcelLibs;
using MiniExcelLibs.OpenXml;
using MiniExcelApi = MiniExcelLibs.MiniExcel;

namespace Bing.Offices.Imports;

/// <summary>
/// 基于 MiniExcel 的 XLSX 导入器。
/// </summary>
/// <remarks>
/// 由 MiniExcel 逐行读取，并由 Core 映射计划负责转换和校验。
/// </remarks>
public sealed class MiniExcelExcelImporter : IExcelImporter, IExcelProviderFeatureDescriptor
{
    /// <inheritdoc />
    public string ProviderName => "MiniExcel";

    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.List
        | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx;

    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;

    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = new[] { ExcelFormat.Xlsx };
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
        "XLS、XLSB、XLSM、ODS 和加密工作簿在当前 MiniExcel Provider 中明确拒绝。"
    };
    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.FormulaCachedValues;
    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;
    /// <summary>
    /// 表示按实体类型异步导入工作表的反射委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="target">执行导入的 MiniExcel 导入器。</param>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="sheetResults">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <param name="mappingPlan">当前工作表的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private delegate Task ImportSheetInvoker<TWorkbook>(MiniExcelExcelImporter target, Stream source,
        string physicalName, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        IExcelMappingPlan mappingPlan, bool isDate1904, WorkbookRowBudget rowBudget) where TWorkbook : class, new();

    /// <summary>
    /// 表示按实体类型同步导入工作表的反射委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="target">执行导入的 MiniExcel 导入器。</param>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="sheetResults">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <param name="mappingPlan">当前工作表的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private delegate void ImportSheetSyncInvoker<TWorkbook>(MiniExcelExcelImporter target, Stream source,
        string physicalName, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        IExcelMappingPlan mappingPlan, bool isDate1904, WorkbookRowBudget rowBudget) where TWorkbook : class, new();

    /// <summary>
    /// 工作簿共享的数据行预算。
    /// </summary>
    /// <remarks>所有工作表共用计数，避免按工作表重复分配预算。</remarks>
    private sealed class WorkbookRowBudget
    {
        /// <summary>
        /// 整个工作簿允许导入的最大数据行数；未指定时不限制。
        /// </summary>
        private readonly int? _maximum;

        /// <summary>
        /// 初始化一个 <see cref="WorkbookRowBudget"/> 类型的实例。
        /// </summary>
        /// <param name="maximum">工作簿最大数据行数；null 表示不限制。</param>
        internal WorkbookRowBudget(int? maximum) => _maximum = maximum;
        /// <summary>
        /// 获取数据行预算是否已超出。
        /// </summary>
        internal bool IsExceeded { get; private set; }

        /// <summary>
        /// 尝试消耗一行数据预算。
        /// </summary>
        /// <returns>成功消耗一行预算时返回 true；已达到上限时返回 false。</returns>
        internal bool TryConsume()
        {
            if (_maximum.HasValue && Count >= _maximum.Value)
            {
                IsExceeded = true;
                return false;
            }
            Count++;
            return true;
        }

        /// <summary>
        /// 获取或设置已消耗的数据行数。
        /// </summary>
        private int Count { get; set; }
    }

    /// <summary>
    /// 按工作簿根实体类型缓存异步工作表导入委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    private static class InvokerCache<TWorkbook> where TWorkbook : class, new()
    {
        /// <summary>
        /// 按行实体运行时类型缓存当前工作簿根类型的异步导入委托。
        /// </summary>
        internal static readonly ConcurrentDictionary<Type, ImportSheetInvoker<TWorkbook>> Values = new();
    }

    /// <summary>
    /// 按工作簿根实体类型缓存同步工作表导入委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    private static class SyncInvokerCache<TWorkbook> where TWorkbook : class, new()
    {
        /// <summary>
        /// 按行实体运行时类型缓存当前工作簿根类型的同步导入委托。
        /// </summary>
        internal static readonly ConcurrentDictionary<Type, ImportSheetSyncInvoker<TWorkbook>> Values = new();
    }

    /// <summary>
    /// 导入时执行的校验规则集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    /// <summary>
    /// 导入时使用的值转换器集合。
    /// </summary>
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    /// <summary>
    /// 按名称解析的导入校验规则集合。
    /// </summary>
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    /// <summary>
    /// 创建导入映射计划的工厂。
    /// </summary>
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    /// <summary>
    /// 按运行时行类型创建 MiniExcel 导入映射计划的构建器。
    /// </summary>
    private readonly MiniExcelMappingPlanBuilder _planBuilder;
    /// <summary>
    /// 向异常观察器分发导入异常的分发器。
    /// </summary>
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>
    /// 初始化一个 <see cref="MiniExcelExcelImporter" /> 类型的实例。
    /// </summary>
    /// <param name="validationRules">校验规则集合。</param>
    /// <param name="valueConverters">值转换器集合。</param>
    /// <param name="namedValidationRules">命名校验规则集合。</param>
    /// <param name="mappingPlanFactory">映射计划工厂。</param>
    /// <param name="exceptionObservers">异常观察器集合。</param>
    public MiniExcelExcelImporter(IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null,
        IExcelMappingPlanFactory mappingPlanFactory = null,
        IEnumerable<IBingOfficesExceptionObserver> exceptionObservers = null)
    {
        _validationRules = validationRules?.ToArray() ?? ExcelValidationRules.CreateDefault();
        _valueConverters = valueConverters?.ToArray() ?? Array.Empty<IExcelValueConverter>();
        _namedValidationRules = namedValidationRules?.ToArray() ?? Array.Empty<INamedExcelValidationRule>();
        _mappingPlanFactory = mappingPlanFactory ?? ExcelMappingPlanFactoryProvider.CreateDefault(
            _valueConverters, _validationRules, _namedValidationRules);
        _planBuilder = new MiniExcelMappingPlanBuilder(_mappingPlanFactory);
        _exceptionDispatcher = new BingOfficesExceptionDispatcher(exceptionObservers);
    }

    /// <inheritdoc />
    public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
        where TWorkbook : class, new()
    {
        ValidateArguments(source, request, cancellationToken);
        try
        {
            using var buffered = CopyToMemory(source, request.ResourceLimits, cancellationToken);
            return ImportBuffered(buffered, request, cancellationToken);
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
            var translated = new BingOfficesImportException("Excel 导入失败。", exception, "MiniExcel",
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
        ValidateArguments(source, request, cancellationToken);
        try
        {
            using var buffered = await CopyToMemoryAsync(source, request.ResourceLimits, cancellationToken)
                .ConfigureAwait(false);
            return await ImportBufferedAsync(buffered, request, cancellationToken).ConfigureAwait(false);
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
            var translated = new BingOfficesImportException("Excel 异步导入失败。", exception, "MiniExcel",
                BingOfficesStage.Read);
            _exceptionDispatcher.Observe(translated);
            throw translated;
        }
    }

    /// <summary>
    /// 从已缓冲的工作簿流同步导入请求的数据。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="source">已缓冲且可定位的工作簿流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <returns>包含根实体、工作表结果和错误的导入结果。</returns>
    private ExcelWorkbookImportResult<TWorkbook> ImportBuffered<TWorkbook>(MemoryStream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        var limits = request.ResourceLimits ?? new ExcelResourceLimits();
        limits.Validate();
        ValidateUnsupportedRequest(request);
        MiniExcelXlsxPreflight.Validate(source, limits, cancellationToken);
        var isDate1904 = MiniExcelXlsxPreflight.GetDate1904(source, cancellationToken);
        var root = new TWorkbook();
        var errors = new List<ExcelImportError>();
        var results = new List<ExcelSheetImportResult>();
        var rowBudget = new WorkbookRowBudget(limits.MaxRows);
        var names = GetSheetNames(source, cancellationToken);
        var plans = BuildPlans(request);
        foreach (var sheet in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsErrorLimitReached(errors, request))
                break;
            var resolvedName = ResolveSheetName(sheet.Selector, names, request.SheetNameComparison);
            if (resolvedName == null)
            {
                AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {DescribeSelector(sheet.Selector)}",
                    DescribeSelector(sheet.Selector), 0, 0, null));
                continue;
            }
            var plan = plans[sheet];
            try
            {
                ImportSheetSync(source, resolvedName, sheet, root, results, errors, request,
                    cancellationToken, plan, isDate1904, rowBudget);
            }
            catch (MiniExcelSheetException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    resolvedName, sheet.HeaderRowIndex + 1, 0, null);
                AddError(errors, request, error);
                results.Add(new ExcelSheetImportResult(resolvedName, sheet.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
            if (rowBudget.IsExceeded)
                break;
        }
        if (rowBudget.IsExceeded)
            return new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(), results, errors,
                IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
        BindRelations(root, request.Relations, errors, request, cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, results, errors,
            IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
    }

    /// <summary>
    /// 从已缓冲的工作簿流异步导入请求的数据。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="source">已缓冲且可定位的工作簿流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <returns>包含根实体、工作表结果和错误的异步导入任务。</returns>
    private async Task<ExcelWorkbookImportResult<TWorkbook>> ImportBufferedAsync<TWorkbook>(
        MemoryStream source, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        var limits = request.ResourceLimits ?? new ExcelResourceLimits();
        limits.Validate();
        ValidateUnsupportedRequest(request);
        MiniExcelXlsxPreflight.Validate(source, limits, cancellationToken);
        var isDate1904 = MiniExcelXlsxPreflight.GetDate1904(source, cancellationToken);
        var root = new TWorkbook();
        var errors = new List<ExcelImportError>();
        var results = new List<ExcelSheetImportResult>();
        var rowBudget = new WorkbookRowBudget(limits.MaxRows);
        var names = GetSheetNames(source, cancellationToken);
        var plans = BuildPlans(request);
        foreach (var sheet in request.Sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsErrorLimitReached(errors, request))
                break;
            var resolvedName = ResolveSheetName(sheet.Selector, names, request.SheetNameComparison);
            if (resolvedName == null)
            {
                AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.InvalidHeader,
                    $"缺少请求的 Sheet: {DescribeSelector(sheet.Selector)}",
                    DescribeSelector(sheet.Selector), 0, 0, null));
                continue;
            }
            var plan = plans[sheet];
            try
            {
                await ImportSheetAsync(source, resolvedName, sheet, root, results, errors, request,
                    cancellationToken, plan, isDate1904, rowBudget).ConfigureAwait(false);
            }
            catch (MiniExcelSheetException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    resolvedName, sheet.HeaderRowIndex + 1, 0, null);
                AddError(errors, request, error);
                results.Add(new ExcelSheetImportResult(resolvedName, sheet.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
            if (rowBudget.IsExceeded)
                break;
        }
        if (rowBudget.IsExceeded)
            return new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(), results, errors,
                IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
        BindRelations(root, request.Relations, errors, request, cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, results, errors,
            IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
    }

    /// <summary>
    /// 按运行时行类型同步处理一个工作表。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="results">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <param name="plan">当前工作表的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private void ImportSheetSync<TWorkbook>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904, WorkbookRowBudget rowBudget) where TWorkbook : class, new()
    {
        SyncInvokerCache<TWorkbook>.Values.GetOrAdd(request.ItemType, CreateSyncInvoker<TWorkbook>)(this,
            source, physicalName, request, root, results, errors, workbookRequest, cancellationToken, plan,
            isDate1904, rowBudget);
    }

    /// <summary>
    /// 按运行时行类型异步处理一个工作表。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="results">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    /// <param name="plan">当前工作表的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private Task ImportSheetAsync<TWorkbook>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904, WorkbookRowBudget rowBudget) where TWorkbook : class, new()
    {
        return InvokerCache<TWorkbook>.Values.GetOrAdd(request.ItemType, CreateInvoker<TWorkbook>)(this, source,
            physicalName, request, root, results, errors, workbookRequest, cancellationToken, plan, isDate1904,
            rowBudget);
    }

    /// <summary>
    /// 为指定行实体类型创建异步导入反射委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="itemType">工作表行实体的运行时类型。</param>
    /// <returns>绑定到指定行实体类型的异步导入委托。</returns>
    private static ImportSheetInvoker<TWorkbook> CreateInvoker<TWorkbook>(Type itemType)
        where TWorkbook : class, new()
    {
        var method = typeof(MiniExcelExcelImporter).GetMethod(nameof(ImportTypedSheet),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(TWorkbook), itemType);
        return (ImportSheetInvoker<TWorkbook>)method.CreateDelegate(typeof(ImportSheetInvoker<TWorkbook>));
    }

    /// <summary>
    /// 为指定行实体类型创建同步导入反射委托。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="itemType">工作表行实体的运行时类型。</param>
    /// <returns>绑定到指定行实体类型的同步导入委托。</returns>
    private static ImportSheetSyncInvoker<TWorkbook> CreateSyncInvoker<TWorkbook>(Type itemType)
        where TWorkbook : class, new()
    {
        var method = typeof(MiniExcelExcelImporter).GetMethod(nameof(ImportTypedSheetSync),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(TWorkbook), itemType);
        return (ImportSheetSyncInvoker<TWorkbook>)method.CreateDelegate(typeof(ImportSheetSyncInvoker<TWorkbook>));
    }

    /// <summary>
    /// 读取指定工作表并按具体行实体类型异步物化数据。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <typeparam name="TItem">当前工作表的行实体类型。</typeparam>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="results">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消读取和物化的令牌。</param>
    /// <param name="plan">当前行类型的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private async Task ImportTypedSheet<TWorkbook, TItem>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904, WorkbookRowBudget rowBudget)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var configuration = CreateConfiguration(request);
        var dateColumns = MiniExcelSheetPlanBuilder.ResolveDateColumns<TItem>(plan, request);
        var rawDateSerials = MiniExcelRawDateSerialReader.Read(source, physicalName, dateColumns,
            request.DataRowStartIndex + 1,
            null,
            cancellationToken);
        source.Position = 0;
        var rowObjects = await MiniExcelApi.QueryAsync(source, true, physicalName, ExcelType.XLSX,
            MiniExcelSheetPlanBuilder.CreateStartCell(request), configuration, cancellationToken).ConfigureAwait(false);
        ImportTypedSheetRows<TWorkbook, TItem>((IEnumerable)rowObjects, physicalName, request, root, results, errors,
            workbookRequest, cancellationToken, plan, isDate1904, rawDateSerials, rowBudget);
    }

    /// <summary>
    /// 读取指定工作表并按具体行实体类型同步物化数据。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <typeparam name="TItem">当前工作表的行实体类型。</typeparam>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="physicalName">工作表实际名称。</param>
    /// <param name="request">工作表导入配置。</param>
    /// <param name="root">接收导入数据的工作簿根实体。</param>
    /// <param name="results">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消读取和物化的令牌。</param>
    /// <param name="plan">当前行类型的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private void ImportTypedSheetSync<TWorkbook, TItem>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904, WorkbookRowBudget rowBudget)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var configuration = CreateConfiguration(request);
        var dateColumns = MiniExcelSheetPlanBuilder.ResolveDateColumns<TItem>(plan, request);
        var rawDateSerials = MiniExcelRawDateSerialReader.Read(source, physicalName, dateColumns,
            request.DataRowStartIndex + 1,
            null,
            cancellationToken);
        source.Position = 0;
        var rowObjects = MiniExcelApi.Query(source, true, physicalName, ExcelType.XLSX,
            MiniExcelSheetPlanBuilder.CreateStartCell(request), configuration);
        ImportTypedSheetRows<TWorkbook, TItem>((IEnumerable)rowObjects, physicalName, request, root, results, errors,
            workbookRequest, cancellationToken, plan, isDate1904, rawDateSerials, rowBudget);
    }

    /// <summary>
    /// 按 MiniExcel 行枚举物化数据，并将当前项映射到真实的一基物理行。
    /// </summary>
    /// <remarks>
    /// 枚举从表头后的首条数据开始，因此跳过正文时仍按表头位置推进行号。
    /// </remarks>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <typeparam name="TItem">当前工作表的行实体类型。</typeparam>
    /// <param name="rowObjects">包含表头和正文对象的 MiniExcel 行枚举。</param>
    /// <param name="physicalName">当前工作表的实际名称。</param>
    /// <param name="request">当前工作表的导入配置。</param>
    /// <param name="root">接收工作表数据的工作簿根实体。</param>
    /// <param name="results">接收工作表结果的集合。</param>
    /// <param name="errors">接收工作簿级错误的集合。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">逐行处理过程中检查的取消令牌。</param>
    /// <param name="plan">当前行类型的映射计划。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rawDateSerials">按物理行列索引的原始日期 serial 集合。</param>
    /// <param name="rowBudget">整个 Workbook 共享的数据行预算。</param>
    private void ImportTypedSheetRows<TWorkbook, TItem>(IEnumerable rowObjects, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904,
        IReadOnlyDictionary<long, double> rawDateSerials, WorkbookRowBudget rowBudget)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var enumerator = rowObjects.GetEnumerator();
        using var enumeratorLifetime = enumerator as IDisposable;
        if (!enumerator.MoveNext())
            throw new MiniExcelSheetException($"Sheet 没有可读取的表头: {physicalName}");
        var first = ToDictionary((object)enumerator.Current);
        var headers = first.Keys.ToArray();
        MiniExcelSheetPlanBuilder.ValidateHeaderCount(headers.Length, request);
        var bindings = MiniExcelSheetPlanBuilder.BuildBindings<TItem>(plan, headers, request);
        MiniExcelSheetPlanBuilder.ValidateUnknownHeaders(headers, bindings, plan, request);
        var items = new List<TItem>();
        var rows = new List<int>();
        var sheetErrors = new List<ExcelImportError>();
        var unique = new UniqueTracker(new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase),
            workbookRequest.ResourceLimits?.MaxTrackedUniqueValues,
            CreateComparer(workbookRequest.ResourceLimits?.UniqueComparison ?? StringComparison.OrdinalIgnoreCase));
        // Query 从表头后一条数据开始枚举，物理行需从表头的下一行（一基）推进。
        var physicalRow = request.HeaderRowIndex + 2;
        var skipped = Math.Max(0, request.DataRowStartIndex - request.HeaderRowIndex - 1);
        var current = first;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (skipped > 0)
                skipped--;
            else
            {
                if (!rowBudget.TryConsume())
                {
                    sheetErrors.Add(new ExcelImportError(
                        ExcelImportErrorCode.ResourceLimit,
                        $"Workbook 数据行数超过限制: {workbookRequest.ResourceLimits?.MaxRows}", physicalName, physicalRow, 0,
                        propertyName: null));
                    break;
                }
                MaterializeRow(current, headers, physicalName, request, plan, bindings, physicalRow, items, rows,
                    sheetErrors, unique, workbookRequest, cancellationToken, typeof(TItem), isDate1904,
                    rawDateSerials);
            }
            if (IsErrorLimitReached(errors, workbookRequest) || IsErrorLimitReached(sheetErrors, workbookRequest))
                break;
            physicalRow++;
            if (!enumerator.MoveNext())
                break;
            current = ToDictionary((object)enumerator.Current);
        }
        AddErrors(errors, sheetErrors, workbookRequest);
        var target = request.Target(root);
        if (target == null)
            throw new BingOfficesConfigurationException($"Workbook 导入目标集合不可写入: {request.Name}",
                stage: BingOfficesStage.Plan);
        var addTarget = CreateCollectionAppender(target.GetType(), typeof(TItem));
        foreach (var item in items)
            addTarget(target, item);
        results.Add(new ExcelSheetImportResult(physicalName, typeof(TItem), rows, sheetErrors));
    }

    /// <summary>
    /// 将一行原始值转换为实体，执行校验并收集行级错误。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="row">当前行的表头和值。</param>
    /// <param name="headers">工作表表头顺序。</param>
    /// <param name="sheetName">工作表实际名称。</param>
    /// <param name="request">当前工作表导入配置。</param>
    /// <param name="plan">当前行类型的映射计划。</param>
    /// <param name="bindings">已解析的列绑定。</param>
    /// <param name="rowNumber">当前行的一基物理行号。</param>
    /// <param name="items">接收成功实体的集合。</param>
    /// <param name="rows">接收成功行索引的集合。</param>
    /// <param name="sheetErrors">接收当前工作表错误的集合。</param>
    /// <param name="unique">当前工作表的唯一性跟踪器。</param>
    /// <param name="workbookRequest">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消行处理的令牌。</param>
    /// <param name="itemType">行实体的运行时类型。</param>
    /// <param name="isDate1904">当前工作簿是否使用 1904 日期系统。</param>
    /// <param name="rawDateSerials">按物理行列索引的原始日期 serial 集合。</param>
    private void MaterializeRow<TWorkbook>(IDictionary<string, object> row, IReadOnlyList<string> headers,
        string sheetName, ExcelSheetImportRequest request, IExcelMappingPlan plan,
        IReadOnlyList<ColumnBinding> bindings, int rowNumber, IList items, ICollection<int> rows,
        ICollection<ExcelImportError> sheetErrors, UniqueTracker unique,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        Type itemType, bool isDate1904, IReadOnlyDictionary<long, double> rawDateSerials)
        where TWorkbook : class, new()
    {
        MiniExcelRowMaterializer.MaterializeRow(row, headers, sheetName, request, plan, bindings,
            rowNumber, items, rows, sheetErrors, unique, workbookRequest, cancellationToken,
            itemType, isDate1904, rawDateSerials);
    }

    /// <summary>
    /// 为工作簿中的所有工作表构建导入映射计划。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="request">包含工作表请求的工作簿导入请求。</param>
    /// <returns>按工作表请求索引的映射计划。</returns>
    private Dictionary<ExcelSheetImportRequest, IExcelMappingPlan> BuildPlans<TWorkbook>(
        ExcelWorkbookImportRequest<TWorkbook> request) where TWorkbook : class, new()
    {
        var plans = new Dictionary<ExcelSheetImportRequest, IExcelMappingPlan>();
        foreach (var sheet in request.Sheets)
        {
            var mappingConfiguration = MiniExcelMappingPlanBuilder.MergeRequestDynamicColumns(
                sheet.MappingConfiguration, sheet.DynamicColumns);
            plans[sheet] = _planBuilder.Create(sheet.ItemType,
                sheet.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
                mappingConfiguration, MappingDirection.Import);
        }
        return plans;
    }

    /// <summary>
    /// 将 MiniExcel 行对象转换为不区分大小写的字典。
    /// </summary>
    /// <param name="value">字典或普通行对象。</param>
    /// <returns>包含行字段和值的字典。</returns>
    private static IDictionary<string, object> ToDictionary(object value)
    {
        if (value is IDictionary<string, object> dictionary)
            return dictionary;
        var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in value.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public))
            result[property.Name] = property.GetValue(value);
        return result;
    }

    /// <summary>
    /// 读取工作簿中的物理工作表名称并恢复流位置。
    /// </summary>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="cancellationToken">用于取消读取的令牌。</param>
    /// <returns>按工作簿顺序排列的工作表名称。</returns>
    private static List<string> GetSheetNames(Stream source, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        source.Position = 0;
        var names = MiniExcelApi.GetSheetNames(source, new OpenXmlConfiguration());
        source.Position = 0;
        return names;
    }

    /// <summary>
    /// 根据工作表请求创建 MiniExcel Open XML 配置。
    /// </summary>
    /// <param name="request">用于读取区域性设置的工作表请求。</param>
    /// <returns>保留表头空白并包含空行的 Open XML 配置。</returns>
    private static OpenXmlConfiguration CreateConfiguration(ExcelSheetImportRequest request) =>
        new OpenXmlConfiguration
        {
            Culture = request.Culture ?? CultureInfo.InvariantCulture,
            TrimColumnNames = false,
            IgnoreEmptyRows = false
        };

    /// <summary>
    /// 校验 MiniExcel 导入请求未启用不受支持的功能。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="request">待校验的工作簿导入请求。</param>
    private static void ValidateUnsupportedRequest<TWorkbook>(ExcelWorkbookImportRequest<TWorkbook> request)
        where TWorkbook : class, new()
    {
        if (request.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || request.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook)
            throw new BingOfficesUnsupportedFeatureException(
                "MiniExcel Provider 暂不支持 Workbook 原生校验。", provider: "MiniExcel",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
        if (request.FailureOptions?.Mode is ExcelImportFailureWorkbookMode mode
            && mode != ExcelImportFailureWorkbookMode.None)
            throw new BingOfficesUnsupportedFeatureException(
                "MiniExcel Provider 暂不支持失败工作簿输出。", provider: "MiniExcel",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
    }

    /// <summary>
    /// 按选择器和名称比较规则解析物理工作表名称。
    /// </summary>
    /// <param name="selector">工作表选择器。</param>
    /// <param name="names">工作簿中的物理工作表名称。</param>
    /// <param name="comparison">工作表名称比较规则。</param>
    /// <returns>匹配的物理名称；未匹配时返回 null。</returns>
    private static string ResolveSheetName(ExcelSheetSelector selector, IReadOnlyList<string> names,
        ExcelNameComparison comparison)
    {
        if (selector.Kind == ExcelSheetSelectorKind.ByIndex)
            return selector.Index.Value < names.Count ? names[selector.Index.Value] : null;
        var comparisonType = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return names.FirstOrDefault(name => string.Equals(name, selector.Name, comparisonType));
    }

    /// <summary>
    /// 为强类型集合编译一次添加器，避免把公开的 ICollection 合同限制为 IList。
    /// </summary>
    /// <param name="collectionType">目标集合的运行时类型。</param>
    /// <param name="itemType">集合元素类型。</param>
    /// <returns>将元素追加到目标集合的委托。</returns>
    private static Action<object, object> CreateCollectionAppender(Type collectionType, Type itemType)
    {
        var contract = typeof(ICollection<>).MakeGenericType(itemType);
        if (!contract.IsAssignableFrom(collectionType))
            throw new BingOfficesConfigurationException(
                $"目标集合 {collectionType.FullName} 未实现 ICollection<{itemType.FullName}>。",
                stage: BingOfficesStage.Plan);
        var collection = Expression.Parameter(typeof(object), "collection");
        var item = Expression.Parameter(typeof(object), "item");
        var add = contract.GetMethod(nameof(ICollection<object>.Add));
        var call = Expression.Call(Expression.Convert(collection, contract), add,
            Expression.Convert(item, itemType));
        return Expression.Lambda<Action<object, object>>(call, collection, item).Compile();
    }

    /// <summary>
    /// 将关系绑定委托给独立协调器，并保留内部职责测试的稳定调用点。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="root">已导入的工作簿根实体。</param>
    /// <param name="relations">待执行的关系绑定请求。</param>
    /// <param name="errors">接收关系绑定错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消关系绑定的令牌。</param>
    private static void BindRelations<TWorkbook>(TWorkbook root, IReadOnlyList<ExcelRelationRequest> relations,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        MiniExcelRelationCoordinator.Bind(root, relations, errors, request, cancellationToken);
    }

    /// <summary>
    /// 将错误集合追加到工作簿错误集合，直到达到资源上限。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="target">接收错误的目标集合。</param>
    /// <param name="source">待追加的错误集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    private static void AddErrors<TWorkbook>(ICollection<ExcelImportError> target,
        IEnumerable<ExcelImportError> source, ExcelWorkbookImportRequest<TWorkbook> request)
        where TWorkbook : class, new()
    {
        foreach (var error in source)
        {
            if (!AddError(target, request, error))
                break;
        }
    }

    /// <summary>
    /// 在未达到错误上限时追加一个导入错误。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="errors">接收错误的集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <param name="error">待追加的错误。</param>
    /// <returns>成功追加时返回 true；达到错误上限时返回 false。</returns>
    private static bool AddError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        if (request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum)
            return false;
        errors.Add(error);
        return true;
    }

    /// <summary>
    /// 判断工作簿错误集合是否已达到配置的上限。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="errors">当前错误集合。</param>
    /// <param name="request">当前工作簿导入请求。</param>
    /// <returns>达到配置上限时返回 true，否则返回 false。</returns>
    private static bool IsErrorLimitReached<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request) where TWorkbook : class, new() =>
        request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum;

    /// <summary>
    /// 根据字符串比较选项创建对应的字符串比较器。
    /// </summary>
    /// <param name="comparison">字符串比较选项。</param>
    /// <returns>与选项对应的字符串比较器。</returns>
    private static IEqualityComparer<string> CreateComparer(StringComparison comparison) => comparison switch
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
    /// 将输入流同步复制到内存并应用输入大小限制。
    /// </summary>
    /// <param name="source">待读取的输入流。</param>
    /// <param name="limits">输入资源限制。</param>
    /// <param name="cancellationToken">用于取消复制的令牌。</param>
    /// <returns>位置重置到开头的内存流。</returns>
    private static MemoryStream CopyToMemory(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        if (limits?.MaxInputBytes is long max && source.CanSeek && source.Length > max)
            throw new BingOfficesResourceLimitException($"Excel 输入流超过最大字节数: {max}",
                provider: "MiniExcel", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Open);
        var destination = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            total += read;
            if (limits?.MaxInputBytes is long maxBytes && total > maxBytes)
                throw new BingOfficesResourceLimitException($"Excel 输入流超过最大字节数: {maxBytes}",
                    provider: "MiniExcel", operation: BingOfficesOperation.Import,
                    stage: BingOfficesStage.Open);
            destination.Write(buffer, 0, read);
        }
        destination.Position = 0;
        return destination;
    }

    /// <summary>
    /// 将输入流异步复制到内存并应用输入大小限制。
    /// </summary>
    /// <param name="source">待读取的输入流。</param>
    /// <param name="limits">输入资源限制。</param>
    /// <param name="cancellationToken">用于取消复制的令牌。</param>
    /// <returns>位置重置到开头的内存流异步任务。</returns>
    private static async Task<MemoryStream> CopyToMemoryAsync(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        if (limits?.MaxInputBytes is long max && source.CanSeek && source.Length > max)
            throw new BingOfficesResourceLimitException($"Excel 输入流超过最大字节数: {max}",
                provider: "MiniExcel", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Open);
        var destination = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
            .ConfigureAwait(false)) > 0)
        {
            total += read;
            if (limits?.MaxInputBytes is long maxBytes && total > maxBytes)
                throw new BingOfficesResourceLimitException($"Excel 输入流超过最大字节数: {maxBytes}",
                    provider: "MiniExcel", operation: BingOfficesOperation.Import,
                    stage: BingOfficesStage.Open);
            await destination.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }
        destination.Position = 0;
        return destination;
    }

    /// <summary>
    /// 校验导入输入流、请求和取消状态。
    /// </summary>
    /// <typeparam name="TWorkbook">导入结果的工作簿根实体类型。</typeparam>
    /// <param name="source">待读取的输入流。</param>
    /// <param name="request">工作簿导入请求。</param>
    /// <param name="cancellationToken">用于取消导入的令牌。</param>
    private static void ValidateArguments<TWorkbook>(Stream source,
        ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken)
        where TWorkbook : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 将工作表选择器转换为错误消息中的文本。
    /// </summary>
    /// <param name="selector">待描述的工作表选择器。</param>
    /// <returns>按索引或名称生成的选择器文本。</returns>
    private static string DescribeSelector(ExcelSheetSelector selector) =>
        selector.Kind == ExcelSheetSelectorKind.ByIndex ? $"#{selector.Index.Value}" : selector.Name;

    /// <summary>
    /// 记录 Excel 列与目标属性之间的绑定关系。
    /// </summary>
    internal sealed class ColumnBinding
    {
        /// <summary>
        /// 初始化一个 <see cref="ColumnBinding" /> 类型的实例。
        /// </summary>
        /// <param name="column">绑定的 Excel 列映射。</param>
        /// <param name="property">绑定的目标属性。</param>
        /// <param name="header">匹配到的表头文本。</param>
        /// <param name="columnIndex">已解析的一基物理列号。</param>
        /// <param name="setter">已编译的属性写入器。</param>
        public ColumnBinding(IExcelMappingColumn column, PropertyInfo property, string header,
            int columnIndex, Action<object, object> setter)
        {
            Column = column;
            Property = property;
            Header = header;
            ColumnIndex = columnIndex;
            Setter = setter;
        }

        /// <summary>
        /// 获取绑定的 Excel 列映射。
        /// </summary>
        public IExcelMappingColumn Column { get; }

        /// <summary>
        /// 获取绑定的目标属性。
        /// </summary>
        public PropertyInfo Property { get; }

        /// <summary>
        /// 获取匹配到的表头文本。
        /// </summary>
        public string Header { get; }

        /// <summary>
        /// 获取已解析的一基物理列号。
        /// </summary>
        public int ColumnIndex { get; }

        /// <summary>
        /// 获取已编译的属性写入器。
        /// </summary>
        public Action<object, object> Setter { get; }
    }

    /// <summary>
    /// 表示工作表级导入处理失败。
    /// </summary>
    internal sealed class MiniExcelSheetException : Exception
    {
        /// <summary>
        /// 初始化一个 <see cref="MiniExcelSheetException" /> 类型的实例。
        /// </summary>
        /// <param name="message">异常消息。</param>
        public MiniExcelSheetException(string message) : base(message) { }
    }

    /// <summary>
    /// 表示单行导入处理失败并携带结构化错误。
    /// </summary>
    internal sealed class MiniExcelRowException : Exception
    {
        /// <summary>
        /// 初始化一个 <see cref="MiniExcelRowException" /> 类型的实例。
        /// </summary>
        /// <param name="error">结构化导入错误。</param>
        public MiniExcelRowException(ExcelImportError error) : base(error.Message) => Error = error;

        /// <summary>
        /// 获取导致当前行失败的结构化导入错误。
        /// </summary>
        public ExcelImportError Error { get; }
    }

}
