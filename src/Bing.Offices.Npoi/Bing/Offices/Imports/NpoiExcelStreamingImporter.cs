using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Providers;
using Bing.Offices.Streaming;
using Bing.Offices.Validations;
using NPOI.XSSF.UserModel;

namespace Bing.Offices.Imports;

/// <summary>逐行解析 XLSX 并复用现有列绑定、转换和校验；不累计实体。</summary>
internal sealed class NpoiExcelStreamingImporter : IExcelStreamingImporter
{
    private readonly IExcelMappingPlanFactory _mapping;
    private readonly BingOfficesExceptionDispatcher _exceptions;
    internal NpoiExcelStreamingImporter(IEnumerable<IExcelValidationRule> rules = null,
        IEnumerable<IExcelValueConverter> converters = null, IEnumerable<INamedExcelValidationRule> namedRules = null,
        IExcelMappingPlanFactory mapping = null, IEnumerable<IBingOfficesExceptionObserver> observers = null)
    {
        _mapping = mapping ?? NpoiMappingPlanFactoryResolver.CreateDefault(converters, rules, namedRules);
        _exceptions = new BingOfficesExceptionDispatcher(observers);
    }

    public ExcelStreamingImportSummary Import<T>(ExcelStreamingImportRequest<T> request, Stream source,
        Func<ExcelImportRow<T>, ExcelRowReadDecision> consumer, ExcelStreamingImportOptions options = null,
        CancellationToken cancellationToken = default) where T : class, new()
    {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        return StreamingXml.Completed(Run(request, source, row => new ValueTask<ExcelRowReadDecision>(consumer(row)),
            options ?? new(), false, cancellationToken));
    }

    public async Task<ExcelStreamingImportSummary> ImportAsync<T>(ExcelStreamingImportRequest<T> request, Stream source,
        Func<ExcelImportRow<T>, CancellationToken, Task<ExcelRowReadDecision>> consumer,
        ExcelStreamingImportOptions options = null, CancellationToken cancellationToken = default) where T : class, new()
    {
        if (consumer == null) throw new ArgumentNullException(nameof(consumer));
        return await Run(request, source, row => new ValueTask<ExcelRowReadDecision>(consumer(row, cancellationToken)),
            options ?? new(), true, cancellationToken).ConfigureAwait(false);
    }

    private async ValueTask<ExcelStreamingImportSummary> Run<T>(ExcelStreamingImportRequest<T> request, Stream source,
        Func<ExcelImportRow<T>, ValueTask<ExcelRowReadDecision>> consumer, ExcelStreamingImportOptions options,
        bool asynchronous, CancellationToken token) where T : class, new()
    {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("输入流不可读。", nameof(source));
        options.Validate(); token.ThrowIfCancellationRequested();
        if (request.Sheet.MaxReadColumns > 16384) throw new ArgumentOutOfRangeException(nameof(request));
        var temporary = new ExcelStreamingTempSession(options.TempDirectory, options.MaxTemporaryBytes);
        Exception failure = null;
        XlsxPackageSession package = null;
        try
        {
            var input = temporary.CreateFile();
            await CopyInput(source, input, options.ResourceLimits.MaxInputBytes.Value, asynchronous, token).ConfigureAwait(false);
            input.Position = 0;
            package = new XlsxPackageSession(input, options, temporary, token);
            await package.Initialize(asynchronous).ConfigureAwait(false);
            return await Consume(request, package, options, consumer, asynchronous, token).ConfigureAwait(false);
        }
        catch (Exception e)
        {
            failure = e;
            if (e is BingOfficesException office) _exceptions.Observe(office);
            throw;
        }
        finally { StreamingCleanup.RunAll(failure, () => package?.Dispose(), temporary.Dispose); }
    }

    private static async ValueTask<bool> CopyInput(Stream source, Stream target, long maxBytes, bool asynchronous, CancellationToken token)
    {
        var buffer = new byte[65536];
        long bytes = 0;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            var read = asynchronous ? await source.ReadAsync(buffer, 0, buffer.Length, token).ConfigureAwait(false)
                : source.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            if (read > maxBytes - bytes) throw new StreamingResourceLimitException("Excel 输入字节数超过限制。");
            bytes += read;
            if (asynchronous) await target.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
            else target.Write(buffer, 0, read);
        }
        if (asynchronous) await target.FlushAsync(token).ConfigureAwait(false); else target.Flush();
        return true;
    }

    private async ValueTask<ExcelStreamingImportSummary> Consume<T>(ExcelStreamingImportRequest<T> request,
        XlsxPackageSession package, ExcelStreamingImportOptions options,
        Func<ExcelImportRow<T>, ValueTask<ExcelRowReadDecision>> consumer, bool asynchronous, CancellationToken token) where T : class, new()
    {
        var config = request.Sheet;
        var state = new SummaryState(options.MaxErrorSamples);
        var selected = config.Selector.Index.HasValue ? config.Selector.Index.Value
            : package.Sheets.FindIndex(sheet => string.Equals(sheet.Name, config.Selector.Name, StringComparison.OrdinalIgnoreCase));
        if (selected < 0 || selected >= package.Sheets.Count)
        {
            state.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader, "未找到请求的工作表。", config.Name, 0, 0, null));
            return state.Result(ImportStopReason.InvalidHeader);
        }
        var (name, path) = package.Sheets[selected];
        var mapping = _mapping.CreateWorkbook<T>(config.MappingDocument ?? new ExcelMappingDocument { UseConventionFallback = true },
            config.MappingConfiguration, MappingDirection.Import, new[] { name }).Sheets[0].Mapping;
        var execution = CreateOptions<T>(config, mapping, package.Date1904);
        using var reader = new XlsxRowReader(package, path, config.MaxReadColumns, options, asynchronous, token);
        XlsxRow header;
        do { header = await reader.Next().ConfigureAwait(false); }
        while (header != null && header.Index < config.HeaderRowIndex);
        if (header == null || header.Index != config.HeaderRowIndex || header.LastColumn >= config.MaxReadColumns || header.Errors.Count != 0
            || header.Cells.Values.Any(value => value.Kind == ExcelCellKind.Error || value.CachedKind == ExcelCellKind.Error))
        {
            state.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader, "表头缺失、无效或超过最大列数。", name, config.HeaderRowIndex + 1, 0, null));
            return state.Result(ImportStopReason.InvalidHeader);
        }
        IReadOnlyDictionary<int, ExcelColumnPlan> columns;
        // 仅为现有绑定器创建一行受限表头；正文从不进入 NPOI DOM。
        using (var headerBook = new XSSFWorkbook())
        {
            var headerRow = headerBook.CreateSheet(name).CreateRow(config.HeaderRowIndex);
            foreach (var pair in header.Cells) headerRow.CreateCell(pair.Key).SetCellValue(pair.Value.Text);
            try { columns = NpoiImportSheetExecutor.CreateColumns(headerRow, execution); }
            catch (NpoiSheetStructureException e)
            {
                state.Add(new ExcelImportError(ExcelImportErrorCode.InvalidHeader, e.Message, name, header.Index + 1, 0, null));
                return state.Result(ImportStopReason.InvalidHeader);
            }
        }
        if (columns.Values.Any(NpoiImportRowMaterializer.IsImageColumn))
            throw new NotSupportedException("流式导入不支持图片列。");
        if (columns.Values.Any(column => column.IsDynamic && column.DynamicDefinition == null))
            throw new NotSupportedException("流式导入要求预先声明动态列。");
        var duplicateValues = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var unique = new UniqueTracker(duplicateValues,
            Math.Min(options.MaxTrackedUniqueValues, options.ResourceLimits.MaxTrackedUniqueValues ?? int.MaxValue),
            Comparer(options.ResourceLimits.UniqueComparison), options.MaxTrackedUniqueCharacters);
        var materializer = new NpoiImportRowMaterializer();
        var maxErrors = Math.Min(options.MaxErrors, options.ResourceLimits.MaxErrors ?? int.MaxValue);
        var expected = config.DataRowStartIndex;
        XlsxRow pending = null;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            XlsxRow row;
            try
            {
                pending ??= await reader.Next().ConfigureAwait(false);
                if (pending == null) return state.Result(ImportStopReason.None);
                if (pending.Index < config.DataRowStartIndex) { pending = null; continue; }
                if (pending.Index > expected && config.StopAtFirstEmptyRow) return state.Result(ImportStopReason.None);
                if (pending.Index > expected && config.ReportEmptyRows) row = new XlsxRow { Index = expected };
                else { row = pending; pending = null; }
                expected = row.Index + 1;
            }
            catch (StreamingResourceLimitException e)
            {
                state.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, e.Message, name, reader.CurrentRowIndex + 1, 0, null));
                return state.Result(ImportStopReason.ResourceLimit);
            }
            if (options.ResourceLimits.MaxRows.HasValue && row.Index - config.DataRowStartIndex >= options.ResourceLimits.MaxRows.Value)
            {
                state.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit, "Sheet 数据行数超过限制。", name, row.Index + 1, 0, null));
                return state.Result(ImportStopReason.ResourceLimit);
            }
            var empty = row.Errors.Count == 0 && row.Cells.Values.All(value => string.IsNullOrWhiteSpace(value.Text));
            if (empty && config.StopAtFirstEmptyRow) return state.Result(ImportStopReason.None);
            if (empty && !config.ReportEmptyRows) continue;
            state.Rows++;
            var errors = new ExcelImportErrorCollector((int)(maxErrors - state.ErrorCount));
            foreach (var pair in row.Errors)
                if (columns.TryGetValue(pair.Key, out var column))
                    errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion, pair.Value.Message, name,
                        row.Index + 1, pair.Key + 1, column.Property.Name,
                        column.IsDynamic ? column.Key : column.Property.Name, column.HeaderName, pair.Value.RawValue));
            if (empty) errors.Add(new ExcelImportError(ExcelImportErrorCode.InvalidInput, "导入数据存在空行", name, row.Index + 1, 0, null));
            T item = null;
            unique.BeginRow();
            if (errors.Count == 0 && materializer.ValidateRawValues(null, columns, duplicateValues, name, row.Index,
                    config.ValidateMode, config.Culture, config.BodyWhitespace, errors, package.Date1904, row.Get))
                materializer.TryCreateItem(null, columns, duplicateValues, unique, name, row.Index, config.ValidateMode, true,
                    errors, config.Culture, config.BodyWhitespace, config.DynamicTargetGetter, null, package.Date1904, out item, row.Get);
            if (item != null && errors.Count == 0) { unique.CommitRow(); state.Success++; }
            else { unique.RollbackRow(); item = null; state.Failed++; }
            foreach (var error in errors.Errors) state.Add(error);
            // 消费者异常不进入资源错误转换分支，保留其原始类型与对象。
            var decision = await consumer(new ExcelImportRow<T>(row.Index, item, errors.Errors)).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            state.Delivered++;
            if (decision is not ExcelRowReadDecision.Continue and not ExcelRowReadDecision.Stop)
                throw new ArgumentOutOfRangeException(nameof(consumer), "消费者返回未知决定。");
            if (decision == ExcelRowReadDecision.Stop) return state.Result(ImportStopReason.ConsumerStopped);
            if (errors.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit)) return state.Result(ImportStopReason.ResourceLimit);
            if (state.ErrorCount >= maxErrors) return state.Result(ImportStopReason.ErrorLimit);
        }
    }

    private static ExcelImportExecutionOptions<T> CreateOptions<T>(ExcelSheetImportRequest request, IExcelMappingPlan mapping,
        bool date1904) where T : class, new() => new()
    {
        HeaderRowIndex = request.HeaderRowIndex, DataRowIndex = request.DataRowStartIndex,
        MaxReadColumns = request.MaxReadColumns, ReadColumnRange = request.ReadColumnRange,
        HeaderComparison = request.HeaderComparison, HeaderWhitespace = request.HeaderWhitespace,
        BodyWhitespace = request.BodyWhitespace, DynamicTargetGetter = request.DynamicTargetGetter,
        RequireExpectedHeaders = request.RequireExpectedHeaders, ValidateMode = request.ValidateMode,
        Culture = request.Culture, DynamicColumns = request.DynamicColumns,
        FailOnUnknownDynamicColumns = request.FailOnUnknownDynamicColumns,
        ReportEmptyRows = request.ReportEmptyRows, StopAtFirstEmptyRow = request.StopAtFirstEmptyRow,
        MappingConfiguration = request.MappingConfiguration, MappingDocument = request.MappingDocument,
        MappingPlan = mapping, IsDate1904 = date1904
    };

    private static StringComparer Comparer(StringComparison comparison) => comparison switch
    {
        StringComparison.Ordinal => StringComparer.Ordinal,
        StringComparison.OrdinalIgnoreCase => StringComparer.OrdinalIgnoreCase,
        StringComparison.InvariantCulture => StringComparer.InvariantCulture,
        StringComparison.InvariantCultureIgnoreCase => StringComparer.InvariantCultureIgnoreCase,
        StringComparison.CurrentCulture => StringComparer.CurrentCulture,
        StringComparison.CurrentCultureIgnoreCase => StringComparer.CurrentCultureIgnoreCase,
        _ => throw new ArgumentOutOfRangeException(nameof(comparison))
    };

    private sealed class SummaryState
    {
        private readonly int _samples;
        private readonly List<ExcelImportError> _errors = new();
        internal long Rows, Success, Failed, Delivered, ErrorCount;
        internal SummaryState(int samples) => _samples = samples;
        internal void Add(ExcelImportError error)
        { ErrorCount++; if (_errors.Count < _samples) _errors.Add(error); }
        internal ExcelStreamingImportSummary Result(ImportStopReason reason)
            => new(Rows, Success, Failed, Delivered, ErrorCount, _errors, reason);
    }
}
