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
/// 基于 MiniExcel 的 XLSX 导入器。MiniExcel 负责逐行读取，Core 映射计划负责转换和校验。
/// </summary>
public sealed class MiniExcelExcelImporter : IExcelImporter
{
    private delegate Task ImportSheetInvoker<TWorkbook>(MiniExcelExcelImporter target, Stream source,
        string physicalName, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        IExcelMappingPlan mappingPlan, bool isDate1904) where TWorkbook : class, new();

    private delegate void ImportSheetSyncInvoker<TWorkbook>(MiniExcelExcelImporter target, Stream source,
        string physicalName, ExcelSheetImportRequest request, TWorkbook root,
        ICollection<ExcelSheetImportResult> sheetResults, ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        IExcelMappingPlan mappingPlan, bool isDate1904) where TWorkbook : class, new();

    private static class InvokerCache<TWorkbook> where TWorkbook : class, new()
    {
        internal static readonly ConcurrentDictionary<Type, ImportSheetInvoker<TWorkbook>> Values = new();
    }

    private static class SyncInvokerCache<TWorkbook> where TWorkbook : class, new()
    {
        internal static readonly ConcurrentDictionary<Type, ImportSheetSyncInvoker<TWorkbook>> Values = new();
    }

    private readonly IReadOnlyList<IExcelValidationRule> _validationRules;
    private readonly IReadOnlyList<IExcelValueConverter> _valueConverters;
    private readonly IReadOnlyList<INamedExcelValidationRule> _namedValidationRules;
    private readonly IExcelMappingPlanFactory _mappingPlanFactory;
    private readonly MiniExcelMappingPlanBuilder _planBuilder;
    private readonly BingOfficesExceptionDispatcher _exceptionDispatcher;

    /// <summary>初始化 MiniExcel 导入器。</summary>
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
                    cancellationToken, plan, isDate1904);
            }
            catch (MiniExcelSheetException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    resolvedName, sheet.HeaderRowIndex + 1, 0, null);
                AddError(errors, request, error);
                results.Add(new ExcelSheetImportResult(resolvedName, sheet.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
        }
        BindRelations(root, request.Relations, errors, request, cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, results, errors,
            IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
    }

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
                    cancellationToken, plan, isDate1904).ConfigureAwait(false);
            }
            catch (MiniExcelSheetException exception)
            {
                var error = new ExcelImportError(ExcelImportErrorCode.InvalidHeader, exception.Message,
                    resolvedName, sheet.HeaderRowIndex + 1, 0, null);
                AddError(errors, request, error);
                results.Add(new ExcelSheetImportResult(resolvedName, sheet.ItemType,
                    Array.Empty<int>(), new[] { error }));
            }
        }
        BindRelations(root, request.Relations, errors, request, cancellationToken);
        return new ExcelWorkbookImportResult<TWorkbook>(root, results, errors,
            IsErrorLimitReached(errors, request), request.ResourceLimits?.MaxErrors);
    }

    private void ImportSheetSync<TWorkbook>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904) where TWorkbook : class, new()
    {
        SyncInvokerCache<TWorkbook>.Values.GetOrAdd(request.ItemType, CreateSyncInvoker<TWorkbook>)(this,
            source, physicalName, request, root, results, errors, workbookRequest, cancellationToken, plan,
            isDate1904);
    }

    private Task ImportSheetAsync<TWorkbook>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904) where TWorkbook : class, new()
    {
        return InvokerCache<TWorkbook>.Values.GetOrAdd(request.ItemType, CreateInvoker<TWorkbook>)(this, source,
            physicalName, request, root, results, errors, workbookRequest, cancellationToken, plan, isDate1904);
    }

    private static ImportSheetInvoker<TWorkbook> CreateInvoker<TWorkbook>(Type itemType)
        where TWorkbook : class, new()
    {
        var method = typeof(MiniExcelExcelImporter).GetMethod(nameof(ImportTypedSheet),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(TWorkbook), itemType);
        return (ImportSheetInvoker<TWorkbook>)method.CreateDelegate(typeof(ImportSheetInvoker<TWorkbook>));
    }

    private static ImportSheetSyncInvoker<TWorkbook> CreateSyncInvoker<TWorkbook>(Type itemType)
        where TWorkbook : class, new()
    {
        var method = typeof(MiniExcelExcelImporter).GetMethod(nameof(ImportTypedSheetSync),
            BindingFlags.Instance | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(TWorkbook), itemType);
        return (ImportSheetSyncInvoker<TWorkbook>)method.CreateDelegate(typeof(ImportSheetSyncInvoker<TWorkbook>));
    }

    private async Task ImportTypedSheet<TWorkbook, TItem>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var configuration = CreateConfiguration(request);
        var rawDateSerials = MiniExcelRawDateSerialReader.Read(source, physicalName, cancellationToken);
        source.Position = 0;
        var rowObjects = await MiniExcelApi.QueryAsync(source, true, physicalName, ExcelType.XLSX,
            CreateStartCell(request), configuration, cancellationToken).ConfigureAwait(false);
        ImportTypedSheetRows<TWorkbook, TItem>((IEnumerable)rowObjects, physicalName, request, root, results, errors,
            workbookRequest, cancellationToken, plan, isDate1904, rawDateSerials);
    }

    private void ImportTypedSheetSync<TWorkbook, TItem>(Stream source, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var configuration = CreateConfiguration(request);
        var rawDateSerials = MiniExcelRawDateSerialReader.Read(source, physicalName, cancellationToken);
        source.Position = 0;
        var rowObjects = MiniExcelApi.Query(source, true, physicalName, ExcelType.XLSX,
            CreateStartCell(request), configuration);
        ImportTypedSheetRows<TWorkbook, TItem>((IEnumerable)rowObjects, physicalName, request, root, results, errors,
            workbookRequest, cancellationToken, plan, isDate1904, rawDateSerials);
    }

    /// <summary>
    /// 按 MiniExcel 行枚举物化数据，并将当前项映射到真实的一基物理行。
    /// 枚举从表头后的首条数据开始，因此跳过正文时仍按表头位置推进行号。
    /// </summary>
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
    private void ImportTypedSheetRows<TWorkbook, TItem>(IEnumerable rowObjects, string physicalName,
        ExcelSheetImportRequest request, TWorkbook root, ICollection<ExcelSheetImportResult> results,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> workbookRequest,
        CancellationToken cancellationToken, IExcelMappingPlan plan, bool isDate1904,
        IReadOnlyDictionary<long, double> rawDateSerials)
        where TWorkbook : class, new()
        where TItem : class, new()
    {
        var enumerator = rowObjects.GetEnumerator();
        using var enumeratorLifetime = enumerator as IDisposable;
        if (!enumerator.MoveNext())
            throw new MiniExcelSheetException($"Sheet 没有可读取的表头: {physicalName}");
        var first = ToDictionary((object)enumerator.Current);
        var headers = first.Keys.ToArray();
        ValidateHeaderCount(headers.Length, request, workbookRequest);
        var bindings = BuildBindings<TItem>(plan, headers, request);
        ValidateUnknownHeaders(headers, bindings, plan, request);
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
                MaterializeRow(current, headers, physicalName, request, plan, bindings, physicalRow, items, rows,
                    sheetErrors, unique, workbookRequest, cancellationToken, typeof(TItem), isDate1904,
                    rawDateSerials);
            if (IsErrorLimitReached(errors, workbookRequest) || IsErrorLimitReached(sheetErrors, workbookRequest))
                break;
            physicalRow++;
            if (!enumerator.MoveNext())
                break;
            current = ToDictionary((object)enumerator.Current);
        }
        AddErrors(errors, sheetErrors, workbookRequest);
        var target = request.Target(root) as IList;
        if (target == null)
            throw new BingOfficesConfigurationException($"Workbook 导入目标集合不可写入: {request.Name}",
                stage: BingOfficesStage.Plan);
        foreach (var item in items)
            target.Add(item);
        results.Add(new ExcelSheetImportResult(physicalName, typeof(TItem), rows, sheetErrors));
    }

    private void MaterializeRow<TWorkbook>(IDictionary<string, object> row, IReadOnlyList<string> headers,
        string sheetName,
        ExcelSheetImportRequest request, IExcelMappingPlan plan, IReadOnlyList<ColumnBinding> bindings,
        int rowNumber, IList items, ICollection<int> rows,
        ICollection<ExcelImportError> sheetErrors, UniqueTracker unique,
        ExcelWorkbookImportRequest<TWorkbook> workbookRequest, CancellationToken cancellationToken,
        Type itemType, bool isDate1904, IReadOnlyDictionary<long, double> rawDateSerials)
        where TWorkbook : class, new()
    {
        var item = Activator.CreateInstance(itemType);
        var valid = true;
        var configuredValidationEnabled = IsConfiguredValidationEnabled(workbookRequest.ValidationMode);
        if (configuredValidationEnabled)
            unique.BeginRow();
        var dynamicValues = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < bindings.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var binding = bindings[index];
            row.TryGetValue(binding.Header, out var raw);
            var text = Normalize(MiniExcelValueAdapter.ToText(raw, request.Culture),
                binding.Column.ImportWhitespace ?? request.BodyWhitespace);
            var columnIndex = FindPhysicalColumnIndex(headers, binding.Header, request.ReadColumnRange?.StartIndex ?? 0);
            var cell = CreateRawDateCell(rawDateSerials, rowNumber, columnIndex, text, isDate1904,
                binding.Property.PropertyType);
            try
            {
                if (configuredValidationEnabled)
                    ValidateBindings(binding.Column.ValidationBindings, text, null, sheetName, rowNumber,
                        columnIndex, binding.Column.Name, raw, request.Culture, binding.Property.PropertyType,
                        unique, binding.Column.IsUnique, binding.Column.UniqueIgnoreEmpty,
                        sheetErrors, isDate1904: isDate1904, cell: cell);
                var converted = MiniExcelValueAdapter.ConvertFrom(raw, binding.Column, binding.Property,
                    sheetName, rowNumber, columnIndex, request.Culture, isDate1904, cell);
                if (configuredValidationEnabled)
                    ValidateBindings(binding.Column.ValidationBindings, text, converted, sheetName, rowNumber,
                        columnIndex, binding.Column.Name, raw, request.Culture, binding.Property.PropertyType,
                        unique, binding.Column.IsUnique, binding.Column.UniqueIgnoreEmpty,
                        sheetErrors, rawOnly: false, isDate1904: isDate1904, cell: cell);
                binding.Property.SetValue(item, converted);
            }
            catch (MiniExcelRowException exception)
            {
                AddSheetError(sheetErrors, workbookRequest, exception.Error);
                valid = false;
                if (request.ValidateMode == ValidateMode.StopOnFirstFailure)
                    break;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddSheetError(sheetErrors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.ValueConversion, exception.Message, sheetName, rowNumber,
                    columnIndex, binding.Column.Name, rawValue: raw));
                valid = false;
                if (request.ValidateMode == ValidateMode.StopOnFirstFailure)
                    break;
            }
        }

        foreach (var dynamic in plan.DynamicColumns)
        {
            var header = FindHeader(row.Keys, dynamic.Title, dynamic.Aliases, request.HeaderComparison,
                request.HeaderWhitespace);
            if (header == null)
                continue;
            row.TryGetValue(header, out var raw);
            var columnIndex = FindPhysicalColumnIndex(headers, header, request.ReadColumnRange?.StartIndex ?? 0);
            var cell = CreateRawDateCell(rawDateSerials, rowNumber, columnIndex,
                MiniExcelValueAdapter.ToText(raw, request.Culture), isDate1904,
                MiniExcelValueAdapter.ResolveDynamicType(dynamic.DataTypeName));
            try
            {
                var converted = MiniExcelValueAdapter.ConvertDynamicFrom(raw, dynamic, sheetName,
                    rowNumber, columnIndex, request.Culture, isDate1904, cell);
                dynamicValues[dynamic.Key] = converted;
                if (configuredValidationEnabled)
                    ValidateBindings(dynamic.ValidationBindings, MiniExcelValueAdapter.ToText(raw, request.Culture),
                        converted, sheetName, rowNumber, columnIndex, dynamic.Key, raw, request.Culture,
                        MiniExcelValueAdapter.ResolveDynamicType(dynamic.DataTypeName), unique, dynamic.IsUnique,
                        dynamic.UniqueIgnoreEmpty, sheetErrors, rawOnly: false, isDate1904: isDate1904,
                        cell: cell);
            }
            catch (MiniExcelRowException exception)
            {
                AddSheetError(sheetErrors, workbookRequest, exception.Error);
                valid = false;
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddSheetError(sheetErrors, workbookRequest, new ExcelImportError(
                    ExcelImportErrorCode.ValueConversion, exception.Message, sheetName, rowNumber, columnIndex,
                    dynamic.Key, rawValue: raw));
                valid = false;
            }
        }
        if (valid)
        {
            SetDynamicValues(item, request, dynamicValues);
            items.Add(item);
            rows.Add(rowNumber - 1);
            if (configuredValidationEnabled)
                unique.CommitRow();
        }
        else if (configuredValidationEnabled)
            unique.RollbackRow();
    }

    private static ExcelCellValue CreateRawDateCell(IReadOnlyDictionary<long, double> rawDateSerials,
        int rowNumber, int columnNumber, string text, bool isDate1904, Type propertyType)
    {
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (targetType != typeof(DateTime) && targetType != typeof(DateTimeOffset))
            return null;
        if (rawDateSerials != null
            && rawDateSerials.TryGetValue(MiniExcelRawDateSerialReader.CreateKey(rowNumber, columnNumber),
                out var serial))
            return new ExcelCellValue(serial, text, ExcelCellKind.Number, isDate1904: isDate1904);
        return null;
    }

    private static int FindPhysicalColumnIndex(IReadOnlyList<string> headers, string header, int startIndex)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            if (string.Equals(headers[index], header, StringComparison.Ordinal))
                return startIndex + index + 1;
        }

        throw new MiniExcelSheetException($"动态列表头无法定位物理列: {header}");
    }

    private static void ValidateBindings(IReadOnlyList<IExcelValidationBinding> bindings, string text,
        object converted, string sheetName, int rowNumber, int columnNumber, string propertyName,
        object raw, CultureInfo culture, Type propertyType, UniqueTracker unique,
        bool isUnique, bool ignoreEmpty, ICollection<ExcelImportError> errors,
        bool rawOnly = true, bool isDate1904 = false, ExcelCellValue cell = null)
    {
        if (bindings == null)
            bindings = Array.Empty<IExcelValidationBinding>();
        foreach (var binding in bindings)
        {
            if (!rawOnly && binding.IsRaw)
                continue;
            if (rawOnly && !binding.IsRaw)
                continue;
            if (!rawOnly && binding.Kind == ExcelValidationBindingKind.Unique)
                continue;
            if (!binding.Validate(new ExcelValidationContext(text, sheetName, rowNumber, columnNumber,
                propertyName, converted, propertyType,
                cell ?? MiniExcelValueAdapter.CreateCell(raw, text, isDate1904), culture)))
            {
                throw new MiniExcelRowException(new ExcelImportError(
                    MiniExcelValueAdapter.GetValidationCode(binding), binding.ErrorMessage, sheetName,
                    rowNumber, columnNumber, propertyName, rawValue: raw));
            }
        }
        if (!rawOnly && isUnique && !unique.TryReserve(propertyName, text, false, ignoreEmpty, rowNumber))
        {
            unique.TryGetFirstRowNumber(propertyName, text, out var firstRow);
            throw new MiniExcelRowException(new ExcelImportError(ExcelImportErrorCode.Validation,
                "重复数据。", sheetName, rowNumber, columnNumber, propertyName,
                rawValue: raw, firstRowNumber: firstRow == 0 ? null : firstRow));
        }
    }

    private static List<ColumnBinding> BuildBindings<TItem>(IExcelMappingPlan plan, string[] headers,
        ExcelSheetImportRequest request) where TItem : class, new()
    {
        var bindings = new List<ColumnBinding>();
        foreach (var column in plan.Columns.Where(column => !column.Ignored && !column.IsDynamicColumn))
        {
            var property = typeof(TItem).GetProperty(column.Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (property != null && IsNavigationOrDynamicContainer(property.PropertyType))
                continue;
            var header = FindHeader(headers, column.Title, column.Aliases, request.HeaderComparison,
                request.HeaderWhitespace);
            if (header == null)
            {
                if (request.RequireExpectedHeaders)
                    throw new MiniExcelSheetException($"Sheet {request.Name} 缺少表头: {column.Title}");
                continue;
            }
            if (property == null || !property.CanWrite)
            {
                throw new MiniExcelSheetException($"属性不可写入: {column.Name}");
            }
            bindings.Add(new ColumnBinding(column, property, header));
        }
        return bindings;
    }

    private static bool IsNavigationOrDynamicContainer(Type propertyType) =>
        typeof(IDictionary<string, object>).IsAssignableFrom(propertyType)
        || (propertyType != typeof(string) && typeof(IEnumerable).IsAssignableFrom(propertyType));

    private static void ValidateUnknownHeaders(string[] headers, IReadOnlyList<ColumnBinding> bindings,
        IExcelMappingPlan plan, ExcelSheetImportRequest request)
    {
        if (!request.FailOnUnknownDynamicColumns)
            return;
        var known = new HashSet<string>(bindings.Select(binding => binding.Header),
            request.HeaderComparison == ExcelNameComparison.Ordinal
                ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase);
        foreach (var column in plan.DynamicColumns)
        {
            foreach (var header in headers)
            {
                if (FindHeader(new[] { header }, column.Title, column.Aliases,
                        request.HeaderComparison, request.HeaderWhitespace) != null)
                    known.Add(header);
            }
        }
        var unknown = headers.FirstOrDefault(header => !known.Contains(header));
        if (unknown != null)
            throw new MiniExcelSheetException($"Sheet {request.Name} 包含未声明动态列: {unknown}");
    }

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

    private static Dictionary<string, object> ToDictionary(object value)
    {
        if (value is IDictionary<string, object> dictionary)
            return new Dictionary<string, object>(dictionary, StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in value.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public))
            result[property.Name] = property.GetValue(value);
        return result;
    }

    private static List<string> GetSheetNames(Stream source, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        source.Position = 0;
        var names = MiniExcelApi.GetSheetNames(source, new OpenXmlConfiguration());
        source.Position = 0;
        return names;
    }

    private static OpenXmlConfiguration CreateConfiguration(ExcelSheetImportRequest request) =>
        new OpenXmlConfiguration
        {
            Culture = request.Culture ?? CultureInfo.InvariantCulture,
            TrimColumnNames = false,
            IgnoreEmptyRows = false
        };

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

    private static string ResolveSheetName(ExcelSheetSelector selector, IReadOnlyList<string> names,
        ExcelNameComparison comparison)
    {
        if (selector.Kind == ExcelSheetSelectorKind.ByIndex)
            return selector.Index.Value < names.Count ? names[selector.Index.Value] : null;
        var comparisonType = comparison == ExcelNameComparison.Ordinal
            ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        return names.FirstOrDefault(name => string.Equals(name, selector.Name, comparisonType));
    }

    private static string FindHeader(IEnumerable<string> headers, string title, IReadOnlyList<string> aliases,
        ExcelNameComparison comparison, ExcelWhitespacePolicy whitespace)
    {
        var expected = new[] { title }.Concat(aliases ?? Array.Empty<string>());
        foreach (var header in headers)
        {
            var normalized = Normalize(header, whitespace);
            foreach (var candidate in expected)
            {
                if (string.Equals(normalized, Normalize(candidate, whitespace),
                    comparison == ExcelNameComparison.Ordinal ? StringComparison.Ordinal
                        : StringComparison.OrdinalIgnoreCase))
                    return header;
            }
        }
        return null;
    }

    private static string Normalize(string value, ExcelWhitespacePolicy policy)
    {
        value ??= string.Empty;
        return policy switch
        {
            ExcelWhitespacePolicy.Preserve => value,
            ExcelWhitespacePolicy.Trim => value.Trim(),
            ExcelWhitespacePolicy.RemoveAll => new string(value.Where(character => !char.IsWhiteSpace(character)).ToArray()),
            _ => throw new ArgumentOutOfRangeException(nameof(policy))
        };
    }

    private static string CreateStartCell(ExcelSheetImportRequest request)
    {
        if (request.HeaderRowIndex == 0 && request.ReadColumnRange == null)
            return "A1";
        var column = request.ReadColumnRange?.StartIndex ?? 0;
        var letters = string.Empty;
        do
        {
            letters = (char)('A' + column % 26) + letters;
            column = column / 26 - 1;
        } while (column >= 0);
        return letters + (request.HeaderRowIndex + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static void ValidateHeaderCount(int count, ExcelSheetImportRequest request,
        object workbookRequest)
    {
        if (count > request.MaxReadColumns)
            throw new MiniExcelSheetException($"Sheet {request.Name} 的表头列数超过限制: {request.MaxReadColumns}");
    }

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
            return;
        }
        var target = request.DynamicTargetGetter?.Invoke(item) as IDictionary<string, object>;
        if (target != null)
            foreach (var pair in values)
                target[pair.Key] = pair.Value;
    }

    private static void BindRelations<TWorkbook>(TWorkbook root, IReadOnlyList<ExcelRelationRequest> relations,
        ICollection<ExcelImportError> errors, ExcelWorkbookImportRequest<TWorkbook> request,
        CancellationToken cancellationToken) where TWorkbook : class, new()
    {
        foreach (var relation in relations ?? Array.Empty<ExcelRelationRequest>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                var parents = (relation.Parents(root) as IEnumerable)?.Cast<object>().ToArray()
                    ?? Array.Empty<object>();
                var children = (relation.Children(root) as IEnumerable)?.Cast<object>().ToArray()
                    ?? Array.Empty<object>();
                foreach (var child in children)
                {
                    var key = relation.ChildKey.DynamicInvoke(child);
                    var parent = parents.FirstOrDefault(candidate => RelationKeysEqual(
                        relation.ParentKey.DynamicInvoke(candidate), key, relation.Comparer));
                    if (parent == null)
                    {
                        AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.Relationship,
                            $"未找到关联父实体: {key}", null, 0, 0, null, rawValue: key));
                        continue;
                    }
                    var navigation = relation.Navigation(parent) as IList;
                    navigation?.Add(child);
                }
            }
            catch (Exception exception) when (exception is not OperationCanceledException
                && exception is not OutOfMemoryException && exception is not StackOverflowException)
            {
                AddError(errors, request, new ExcelImportError(ExcelImportErrorCode.Relationship,
                    exception.Message, null, 0, 0, null));
            }
        }
    }

    private static bool RelationKeysEqual(object left, object right, object comparer)
    {
        if (comparer is System.Collections.IEqualityComparer nonGeneric)
            return nonGeneric.Equals(left, right);
        if (comparer is IEqualityComparer<object> objectComparer)
            return objectComparer.Equals(left, right);
        if (comparer != null)
        {
            var leftType = left?.GetType() ?? right?.GetType();
            if (leftType != null)
            {
                var equals = comparer.GetType().GetMethod(nameof(object.Equals),
                    BindingFlags.Instance | BindingFlags.Public, binder: null,
                    types: new[] { leftType, leftType }, modifiers: null);
                if (equals != null)
                    return equals.Invoke(comparer, new[] { left, right }) is true;
            }
        }
        return Equals(left, right);
    }

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

    private static bool AddError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        if (request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum)
            return false;
        errors.Add(error);
        return true;
    }

    private static void AddSheetError<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request, ExcelImportError error)
        where TWorkbook : class, new()
    {
        AddError(errors, request, error);
    }

    private static bool IsErrorLimitReached<TWorkbook>(ICollection<ExcelImportError> errors,
        ExcelWorkbookImportRequest<TWorkbook> request) where TWorkbook : class, new() =>
        request.ResourceLimits?.MaxErrors is int maximum && errors.Count >= maximum;

    private static bool IsErrorLimitReached(ICollection<ExcelImportError> errors,
        object request) => false;

    private static bool IsConfiguredValidationEnabled(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

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

    private static string DescribeSelector(ExcelSheetSelector selector) =>
        selector.Kind == ExcelSheetSelectorKind.ByIndex ? $"#{selector.Index.Value}" : selector.Name;

    private sealed class ColumnBinding
    {
        public ColumnBinding(IExcelMappingColumn column, PropertyInfo property, string header)
        {
            Column = column;
            Property = property;
            Header = header;
        }

        public IExcelMappingColumn Column { get; }
        public PropertyInfo Property { get; }
        public string Header { get; }
    }

    private sealed class MiniExcelSheetException : Exception
    {
        public MiniExcelSheetException(string message) : base(message) { }
    }

    private sealed class MiniExcelRowException : Exception
    {
        public MiniExcelRowException(ExcelImportError error) : base(error.Message) => Error = error;
        public ExcelImportError Error { get; }
    }

}
