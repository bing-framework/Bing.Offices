using System.Collections;
using System.Globalization;
using System.IO.Compression;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 执行受能力边界约束的 NPOI XLSX 实体批次前向导入。
/// </summary>
/// <remarks>
/// 该实现不创建 NPOI Workbook DOM，而是按 XML 行顺序解析；支持单格固定单元格和列表区域命名锚点，
/// 可定位且位于起始位置的输入流会直接读取，
/// 其他输入暂存到带自动清理的临时文件。摘要模式的最终实体只保留空聚合对象，固定字段值仍提供给批次回调。
/// SummaryOnly 与 ErrorRowsOnly 组合直接捕获表头和错误行；其他失败模式仍在解析完成后按原有路径生成失败工作簿。
/// </remarks>
internal sealed class NpoiForwardXlsxEntityBatchImporter
{
    /// <summary>
    /// 实体布局失败输出使用的空 Workbook 请求映射。
    /// </summary>
    private static readonly IReadOnlyDictionary<string, ExcelSheetImportRequest>
        EmptyResolvedSheetRequests = new Dictionary<string, ExcelSheetImportRequest>(
            StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// 负责创建映射计划和物化实体的 NPOI 导入器。
    /// </summary>
    private readonly NpoiExcelImporter _importer;

    /// <summary>
    /// 初始化一个 <see cref="NpoiForwardXlsxEntityBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="importer">NPOI 导入器。</param>
    internal NpoiForwardXlsxEntityBatchImporter(NpoiExcelImporter importer)
    {
        _importer = importer ?? throw new ArgumentNullException(nameof(importer));
    }

    /// <summary>
    /// 同步执行前向实体批次导入。
    /// </summary>
    internal ExcelEntityBatchImportSummary<TEntity> Import<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken,
        bool writeFailureWorkbook = true,
        ExcelDataReaderFailureWorkbookCapture failureCaptureOverride = null,
        bool captureFixedCellFailures = true, ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        ValidateRequest(source, layout, options, onBatch);
        temporaryStorage ??= new ExcelTemporaryStorageBudget(options.MaxTemporaryStorageBytes, "NPOI");
        using var input = OpenSource(source, options.ImportOptions.ResourceLimits.MaxInputBytes,
            cancellationToken, temporaryStorage);
        var failureCapture = writeFailureWorkbook
            ? CreateFailureCapture(options)
            : failureCaptureOverride;
        var result = ImportBuffered(input.Stream, layout, options, onBatch, cancellationToken,
            failureCapture, captureFixedCellFailures, temporaryStorage);
        if (writeFailureWorkbook)
            WriteFailureWorkbook(input.Stream, options, result, cancellationToken, failureCapture);
        return result;
    }

    /// <summary>
    /// 异步执行前向实体批次导入。
    /// </summary>
    internal async Task<ExcelEntityBatchImportSummary<TEntity>> ImportAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken, bool writeFailureWorkbook = true,
        ExcelDataReaderFailureWorkbookCapture failureCaptureOverride = null,
        bool captureFixedCellFailures = true, ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateBatchLayout(layout, options);
        ValidateResultMode(options);
        temporaryStorage ??= new ExcelTemporaryStorageBudget(options.MaxTemporaryStorageBytes, "NPOI");
        await using var input = await OpenSourceAsync(source,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken, temporaryStorage)
            .ConfigureAwait(false);
        var failureCapture = writeFailureWorkbook
            ? CreateFailureCapture(options)
            : failureCaptureOverride;
        var result = await ImportBufferedAsync(input.Stream, layout, options, onBatch,
            cancellationToken, failureCapture, captureFixedCellFailures, temporaryStorage).ConfigureAwait(false);
        if (writeFailureWorkbook)
            await WriteFailureWorkbookAsync(input.Stream, options, result, cancellationToken,
                failureCapture)
                .ConfigureAwait(false);
        return result;
    }

    /// <summary>
    /// 为前向导入结果生成错误行失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="buffered">可定位的原始 XLSX 流。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="result">前向导入结果。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="failureCapture">低保留错误行捕获器。</param>
    /// <param name="requiredRowsBySheet">失败工作簿需要保留的固定结构行。</param>
    internal void WriteFailureWorkbook<TEntity>(Stream buffered,
        ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        IReadOnlyDictionary<string, IReadOnlyList<int>> requiredRowsBySheet = null)
        where TEntity : class, new()
    {
        var failureOptions = options.ImportOptions.FailureOptions;
        if (failureOptions == null || failureOptions.Mode == ExcelImportFailureWorkbookMode.None
            || result.Result.Errors.Count == 0)
            return;
        try
        {
            if (IsLowRetentionFailureOutput(options, failureCapture))
            {
                ExcelForwardXlsxFailureWorkbookWriter.Write(failureOptions,
                    result.Result.Errors, failureCapture, cancellationToken,
                    _importer.FileExportCommitter);
                return;
            }
            buffered.Position = 0;
            using var workbook = WorkbookFactory.Create(new NpoiNonDisposingStream(buffered,
                cancellationToken));
            NpoiFailureWorkbookWriter.Write(workbook, failureOptions, result.Result.Errors,
                EmptyResolvedSheetRequests, cancellationToken, new SystemFailureWorkbookFileSystem(),
                null, _importer.FileExportCommitter, requiredRowsBySheet);
        }
        catch (BingOfficesException exception)
        {
            _importer.ObserveException(exception);
            throw;
        }
    }

    /// <summary>
    /// 异步为前向导入结果生成错误行失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="buffered">可定位的原始 XLSX 流。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="result">前向导入结果。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>代表失败工作簿输出的任务。</returns>
    /// <param name="failureCapture">低保留错误行捕获器。</param>
    /// <param name="requiredRowsBySheet">失败工作簿需要保留的固定结构行。</param>
    internal async Task WriteFailureWorkbookAsync<TEntity>(Stream buffered,
        ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        IReadOnlyDictionary<string, IReadOnlyList<int>> requiredRowsBySheet = null)
        where TEntity : class, new()
    {
        var failureOptions = options.ImportOptions.FailureOptions;
        if (failureOptions == null || failureOptions.Mode == ExcelImportFailureWorkbookMode.None
            || result.Result.Errors.Count == 0)
            return;
        try
        {
            if (IsLowRetentionFailureOutput(options, failureCapture))
            {
                await ExcelForwardXlsxFailureWorkbookWriter.WriteAsync(failureOptions,
                    result.Result.Errors, failureCapture, cancellationToken,
                    _importer.FileExportCommitter).ConfigureAwait(false);
                return;
            }
            buffered.Position = 0;
            using var workbook = WorkbookFactory.Create(new NpoiNonDisposingStream(buffered,
                cancellationToken));
            await NpoiFailureWorkbookWriter.WriteAsync(workbook, failureOptions,
                result.Result.Errors, EmptyResolvedSheetRequests, cancellationToken,
                _importer.FileExportCommitter, requiredRowsBySheet).ConfigureAwait(false);
        }
        catch (BingOfficesException exception)
        {
            _importer.ObserveException(exception);
            throw;
        }
    }

    /// <summary>
    /// 为 HSSF 前向导入结果生成低保留或完整失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">实体聚合对象类型。</typeparam>
    /// <param name="source">原始 HSSF 输入流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="result">前向导入结果。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="failureCapture">低保留错误行捕获器。</param>
    internal void WriteHssfFailureWorkbook<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null)
        where TEntity : class, new()
    {
        if (layout.Cells.Count == 0
            && IsLowRetentionFailureOutput(options, failureCapture)
            && NpoiHssfForwardFailureWorkbookWriter.CanWrite(failureCapture,
                result.Result.Errors))
        {
            try
            {
                NpoiHssfForwardFailureWorkbookWriter.Write(
                    options.ImportOptions.FailureOptions, result.Result.Errors, failureCapture,
                    cancellationToken, _importer.FileExportCommitter);
                return;
            }
            catch (BingOfficesException exception)
            {
                _importer.ObserveException(exception);
                throw;
            }
        }

        RejectUnsafeConditionalFormattingFailureOutput(options, failureCapture);
        source.Position = 0;
        WriteFailureWorkbook(source, options, result, cancellationToken, failureCapture: null,
            requiredRowsBySheet: GetRequiredRowsBySheet(layout));
    }

    /// <summary>
    /// 异步为 HSSF 前向导入结果生成低保留或完整失败工作簿。
    /// </summary>
    /// <typeparam name="TEntity">实体聚合对象类型。</typeparam>
    /// <param name="source">原始 HSSF 输入流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="result">前向导入结果。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <param name="failureCapture">低保留错误行捕获器。</param>
    internal async Task WriteHssfFailureWorkbookAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null)
        where TEntity : class, new()
    {
        if (layout.Cells.Count == 0
            && IsLowRetentionFailureOutput(options, failureCapture)
            && NpoiHssfForwardFailureWorkbookWriter.CanWrite(failureCapture,
                result.Result.Errors))
        {
            try
            {
                await NpoiHssfForwardFailureWorkbookWriter.WriteAsync(
                    options.ImportOptions.FailureOptions, result.Result.Errors, failureCapture,
                    cancellationToken, _importer.FileExportCommitter).ConfigureAwait(false);
                return;
            }
            catch (BingOfficesException exception)
            {
                _importer.ObserveException(exception);
                throw;
            }
        }

        RejectUnsafeConditionalFormattingFailureOutput(options, failureCapture);
        source.Position = 0;
        await WriteFailureWorkbookAsync(source, options, result, cancellationToken,
            failureCapture: null, requiredRowsBySheet: GetRequiredRowsBySheet(layout))
            .ConfigureAwait(false);
    }

    /// <summary>
    /// 拒绝无法在错误行工作簿中安全保留的 HSSF 条件格式。
    /// </summary>
    /// <param name="options">实体批次导入选项。</param>
    /// <param name="failureCapture">失败行捕获快照。</param>
    private static void RejectUnsafeConditionalFormattingFailureOutput(
        ExcelEntityBatchImportOptions options,
        ExcelDataReaderFailureWorkbookCapture failureCapture)
    {
        if (!IsLowRetentionFailureOutput(options, failureCapture)
            || failureCapture?.Sheets.Any(sheet => sheet?.HasConditionalFormatting == true) != true)
            return;
        throw new BingOfficesUnsupportedFeatureException(
            "NPOI HSSF Forward 错误行工作簿无法安全保留复杂或跨省略行的条件格式。",
            provider: "NPOI", operation: BingOfficesOperation.Import,
            stage: BingOfficesStage.Write);
    }

    /// <summary>
    /// 同步预检前向实体批次使用的模板工作簿。
    /// </summary>
    internal void ValidateTemplate<TEntity>(ExcelEntityTemplateOptions template,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateBatchLayout(layout, options);
        using var input = OpenSource(template.Template,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken);
        ValidateTemplateStream(input.Stream, layout, options.ImportOptions.ResourceLimits,
            cancellationToken);
    }

    /// <summary>
    /// 异步预检前向实体批次使用的模板工作簿。
    /// </summary>
    internal async Task ValidateTemplateAsync<TEntity>(ExcelEntityTemplateOptions template,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (template == null)
            throw new ArgumentNullException(nameof(template));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateBatchLayout(layout, options);
        await using var input = await OpenSourceAsync(template.Template,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken)
            .ConfigureAwait(false);
        ValidateTemplateStream(input.Stream, layout, options.ImportOptions.ResourceLimits,
            cancellationToken);
    }

    /// <summary>
    /// 校验前向模板的 XLSX 结构和目标工作表。
    /// </summary>
    private static void ValidateTemplateStream<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelResourceLimits limits,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        EnsureXlsxSignature(buffered);
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "NPOI", requireZip: true,
            cancellationToken);
        buffered.Position = 0;
        using var archive = new ZipArchive(buffered, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ReadWorkbook(archive, cancellationToken);
        ValidateForwardMerges(archive, workbook, layout, cancellationToken);
        foreach (var region in layout.ListRegions)
        {
            ValidateWorkbook(workbook, region.SheetName);
            if (region.AnchorName != null)
                ResolveNamedAnchor(workbook, region.SheetName, region.AnchorName);
            var sheet = workbook.FindSheet(region.SheetName);
            if (sheet == null || archive.GetEntry(sheet.Path) == null)
                throw new BingOfficesImportException("XLSX 模板工作表部件不存在。", provider: "NPOI",
                    stage: BingOfficesStage.Read, sheetName: region.SheetName);
        }
        foreach (var cell in layout.Cells)
        {
            ValidateWorkbook(workbook, cell.SheetName);
            if (cell.AnchorName != null)
                ResolveNamedAnchor(workbook, cell.SheetName, cell.AnchorName);
            var sheet = workbook.FindSheet(cell.SheetName);
            if (sheet == null || archive.GetEntry(sheet.Path) == null)
                throw new BingOfficesImportException("XLSX 模板工作表部件不存在。", provider: "NPOI",
                    stage: BingOfficesStage.Read, sheetName: cell.SheetName);
        }
    }

    /// <summary>
    /// 从已缓冲的 XLSX 流同步导入实体。
    /// </summary>
    private ExcelEntityBatchImportSummary<TEntity> ImportBuffered<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        bool captureFixedCellFailures = true, ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        var limits = options.ImportOptions.ResourceLimits;
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        EnsureXlsxSignature(buffered);
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "NPOI", requireZip: true,
            cancellationToken);
        var workbookValidation = UsesWorkbookValidation(options.ImportOptions.ValidationMode)
            ? ExcelXlsxWorkbookValidation.Load(buffered, cancellationToken)
            : null;
        buffered.Position = 0;
        using var archive = new ZipArchive(buffered, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ReadWorkbook(archive, cancellationToken);
        failureCapture?.CaptureWorkbookMetadata(
            ExcelXlsxWorkbookValidation.LoadFailureMetadata(archive, cancellationToken));
        ValidateForwardMerges(archive, workbook, layout, cancellationToken);
        using var sharedStrings = ReadSharedStrings(archive, cancellationToken,
            options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
                || options.MaxTemporaryStorageBytes.HasValue, temporaryStorage);
        var styles = ReadStyles(archive, cancellationToken);
        var entity = new TEntity();
        var rowRuntime = new ExcelImportRuntime(limits);
        var globalErrors = new ExcelImportErrorCollector(limits.MaxErrors);
        var uniqueBudget = new NpoiForwardUniqueBudget(limits.MaxTrackedUniqueValues);
        var relationKeyBudget = new NpoiRelationKeyBudget(limits.MaxTrackedRelationKeys);
        var relationTracker = options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && layout.Relations.Count > 0
            ? ExcelEntityForwardRelationTracker.Create(layout, relationKeyBudget.TryReserve,
                "NPOI")
            : null;
        var fixedCells = ApplyFixedCells(archive, workbook, sharedStrings, styles, layout,
            options, entity, workbookValidation, cancellationToken,
            captureFixedCellFailures ? failureCapture : null, globalErrors);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var sheets = new List<ExcelSheetImportResult>(layout.ListRegions.Count);
        var errors = fixedCells.Errors.ToList();
        var errorsTruncated = fixedCells.IsTruncated;
        var needsSourceLocations = options.ResultMode == ExcelEntityBatchResultMode.Complete
            && layout.Relations.Count > 0;
        var sourceRows = needsSourceLocations
            ? new Dictionary<object, ExcelSheetImportResult>(ReferenceObjectComparer.Instance)
            : null;
        var globalBatchNumber = 1;
        if (fixedCells.Errors.Count > 0)
            onBatch(new ExcelEntityImportBatch<TEntity>(entity, null, null, typeof(TEntity), -1,
                globalBatchNumber++, 0, Array.Empty<object>(), fixedCells.Errors, true));
        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            var region = layout.ListRegions[index];
            var resolvedRegion = region.AnchorName == null
                ? region
                : region.WithResolvedStart(ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName));
            ValidateWorkbook(workbook, resolvedRegion.SheetName);
            var sheet = workbook.FindSheet(resolvedRegion.SheetName);
            var sheetEntry = sheet == null ? null : archive.GetEntry(sheet.Path);
            if (sheetEntry == null)
                throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "NPOI",
                    stage: BingOfficesStage.Read, sheetName: resolvedRegion.SheetName);
            var regionResult = InvokeImportRows(sheetEntry, sharedStrings, styles, workbook.Date1904,
                layout, resolvedRegion, options, workbookValidation, batch =>
                {
                    onBatch(new ExcelEntityImportBatch<TEntity>(entity, batch.SheetName,
                        batch.PropertyName, batch.ItemType, index, globalBatchNumber++,
                        batch.FirstRowIndex, batch.Items, batch.Errors, batch.IsDiagnostic));
                }, cancellationToken, entity, failureCapture, rowRuntime,
                globalErrors, uniqueBudget, relationTracker);
            if (needsSourceLocations)
                CaptureSourceRows(entity, resolvedRegion, regionResult, sourceRows);
            AppendRegionResult(regionResult, index, summaries, sheets, errors, ref errorsTruncated);
            if (relationTracker?.IsTruncated == true)
                break;
        }
        if (relationTracker != null)
            errorsTruncated |= relationTracker.IsTruncated;
        else if (needsSourceLocations)
            BindRelations(entity, layout.Relations, errors, ref errorsTruncated, sourceRows,
                options.ImportOptions.ResourceLimits.MaxErrors, relationKeyBudget, cancellationToken);
        return CreateCombinedSummary(entity, summaries, sheets, errors, errorsTruncated, options);
    }

    /// <summary>
    /// 从已缓冲的 XLSX 流异步导入实体。
    /// </summary>
    private async Task<ExcelEntityBatchImportSummary<TEntity>> ImportBufferedAsync<TEntity>(
        Stream buffered, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        bool captureFixedCellFailures = true, ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        // ZIP/XML 解析器本身是前向同步读取器；外围复制和批次回调保持真实异步。
        var limits = options.ImportOptions.ResourceLimits;
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        EnsureXlsxSignature(buffered);
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "NPOI", requireZip: true,
            cancellationToken);
        var workbookValidation = UsesWorkbookValidation(options.ImportOptions.ValidationMode)
            ? ExcelXlsxWorkbookValidation.Load(buffered, cancellationToken)
            : null;
        buffered.Position = 0;
        using var archive = new ZipArchive(buffered, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ReadWorkbook(archive, cancellationToken);
        failureCapture?.CaptureWorkbookMetadata(
            ExcelXlsxWorkbookValidation.LoadFailureMetadata(archive, cancellationToken));
        ValidateForwardMerges(archive, workbook, layout, cancellationToken);
        using var sharedStrings = ReadSharedStrings(archive, cancellationToken,
            options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
                || options.MaxTemporaryStorageBytes.HasValue, temporaryStorage);
        var styles = ReadStyles(archive, cancellationToken);
        var entity = new TEntity();
        var rowRuntime = new ExcelImportRuntime(limits);
        var globalErrors = new ExcelImportErrorCollector(limits.MaxErrors);
        var uniqueBudget = new NpoiForwardUniqueBudget(limits.MaxTrackedUniqueValues);
        var relationKeyBudget = new NpoiRelationKeyBudget(limits.MaxTrackedRelationKeys);
        var relationTracker = options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && layout.Relations.Count > 0
            ? ExcelEntityForwardRelationTracker.Create(layout, relationKeyBudget.TryReserve,
                "NPOI")
            : null;
        var fixedCells = ApplyFixedCells(archive, workbook, sharedStrings, styles, layout,
            options, entity, workbookValidation, cancellationToken,
            captureFixedCellFailures ? failureCapture : null, globalErrors);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var sheets = new List<ExcelSheetImportResult>(layout.ListRegions.Count);
        var errors = fixedCells.Errors.ToList();
        var errorsTruncated = fixedCells.IsTruncated;
        var needsSourceLocations = options.ResultMode == ExcelEntityBatchResultMode.Complete
            && layout.Relations.Count > 0;
        var sourceRows = needsSourceLocations
            ? new Dictionary<object, ExcelSheetImportResult>(ReferenceObjectComparer.Instance)
            : null;
        var globalBatchNumber = 1;
        if (fixedCells.Errors.Count > 0)
        {
            await onBatch(new ExcelEntityImportBatch<TEntity>(entity, null, null, typeof(TEntity), -1,
                globalBatchNumber++, 0, Array.Empty<object>(), fixedCells.Errors, true),
                cancellationToken).ConfigureAwait(false);
        }
        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            var region = layout.ListRegions[index];
            var resolvedRegion = region.AnchorName == null
                ? region
                : region.WithResolvedStart(ResolveNamedAnchor(workbook, region.SheetName,
                    region.AnchorName));
            ValidateWorkbook(workbook, resolvedRegion.SheetName);
            var sheet = workbook.FindSheet(resolvedRegion.SheetName);
            var sheetEntry = sheet == null ? null : archive.GetEntry(sheet.Path);
            if (sheetEntry == null)
                throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "NPOI",
                    stage: BingOfficesStage.Read, sheetName: resolvedRegion.SheetName);
            var regionResult = await InvokeImportRowsAsync(sheetEntry, sharedStrings, styles,
                workbook.Date1904, layout, resolvedRegion, options, workbookValidation, async (batch, _) =>
                {
                    await onBatch(new ExcelEntityImportBatch<TEntity>(entity, batch.SheetName,
                        batch.PropertyName, batch.ItemType, index, globalBatchNumber++,
                        batch.FirstRowIndex, batch.Items, batch.Errors, batch.IsDiagnostic),
                        cancellationToken).ConfigureAwait(false);
                }, cancellationToken, entity, failureCapture, rowRuntime,
                globalErrors, uniqueBudget, relationTracker).ConfigureAwait(false);
            if (needsSourceLocations)
                CaptureSourceRows(entity, resolvedRegion, regionResult, sourceRows);
            AppendRegionResult(regionResult, index, summaries, sheets, errors, ref errorsTruncated);
            if (relationTracker?.IsTruncated == true)
                break;
        }
        if (relationTracker != null)
            errorsTruncated |= relationTracker.IsTruncated;
        else if (needsSourceLocations)
            BindRelations(entity, layout.Relations, errors, ref errorsTruncated, sourceRows,
                options.ImportOptions.ResourceLimits.MaxErrors, relationKeyBudget, cancellationToken);
        return CreateCombinedSummary(entity, summaries, sheets, errors, errorsTruncated, options);
    }

    /// <summary>
    /// 保存前向读取实体与源工作表行号的对应关系。
    /// </summary>
    private static void CaptureSourceRows<TEntity>(TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelEntityBatchImportSummary<TEntity> result,
        IDictionary<object, ExcelSheetImportResult> sourceRows) where TEntity : class, new()
    {
        var sheet = result.Result.Sheets.SingleOrDefault(item =>
            string.Equals(item.Name, region.SheetName, StringComparison.OrdinalIgnoreCase)
            && item.ItemType == region.ItemType);
        if (sheet == null)
            return;
        var items = region.Getter(entity);
        if (items == null)
            return;
        sourceRows[items] = sheet;
    }

    /// <summary>
    /// 在所有列表完成后执行前向关系绑定并合并关系错误。
    /// </summary>
    private static void BindRelations<TEntity>(TEntity entity,
        IReadOnlyList<ExcelRelationRequest> relations, IList<ExcelImportError> errors,
        ref bool errorsTruncated,
        IReadOnlyDictionary<object, ExcelSheetImportResult> sourceRows,
        int? maxErrors, NpoiRelationKeyBudget relationKeyBudget,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        if (relations == null || relations.Count == 0)
            return;
        var remainingErrors = maxErrors.HasValue
            ? Math.Max(0, maxErrors.Value - errors.Count)
            : (int?)null;
        var collector = new ExcelImportErrorCollector(remainingErrors);
        foreach (var relation in relations)
        {
            cancellationToken.ThrowIfCancellationRequested();
            NpoiRelationBinder.BindWithSourceRows(entity, relation, collector, sourceRows,
                relationKeyBudget, cancellationToken);
        }
        foreach (var error in collector.Errors)
            errors.Add(error);
        errorsTruncated |= collector.IsTruncated;
    }

    /// <summary>
    /// 解析当前 Sheet，并同步交付批次。
    /// </summary>
    private ExcelEntityBatchImportSummary<TEntity> InvokeImportRows<TEntity>(ZipArchiveEntry sheetEntry,
        IReadOnlyList<string> sharedStrings, XlsxStyleTable styles, bool date1904,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region,
        ExcelEntityBatchImportOptions options, ExcelXlsxWorkbookValidation workbookValidation,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken, TEntity entity = null,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelImportRuntime runtime = null,
        ExcelImportErrorCollector globalErrors = null,
        NpoiForwardUniqueBudget uniqueBudget = null,
        ExcelEntityForwardRelationTracker relationTracker = null)
        where TEntity : class, new()
    {
        var method = typeof(NpoiForwardXlsxEntityBatchImporter).GetMethod(nameof(ImportRowsTyped),
            BindingFlags.Instance | BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        try
        {
            return (ExcelEntityBatchImportSummary<TEntity>)method.Invoke(this, new object[]
            {
                sheetEntry, sharedStrings, styles, date1904, layout, region, options,
                workbookValidation, onBatch,
                cancellationToken, entity, failureCapture, runtime, globalErrors, uniqueBudget,
                relationTracker
            });
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 解析当前 Sheet，并同步交付批次。
    /// </summary>
    private ExcelEntityBatchImportSummary<TEntity> ImportRowsTyped<TEntity, TItem>(
        ZipArchiveEntry sheetEntry, IReadOnlyList<string> sharedStrings, XlsxStyleTable styles,
        bool date1904,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region,
        ExcelEntityBatchImportOptions options, ExcelXlsxWorkbookValidation workbookValidation,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken, TEntity targetEntity,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        ExcelImportRuntime runtime,
        ExcelImportErrorCollector globalErrors,
        NpoiForwardUniqueBudget uniqueBudget,
        ExcelEntityForwardRelationTracker relationTracker)
        where TEntity : class, new() where TItem : class, new()
    {
        var context = CreateContext<TItem, TEntity>(layout, region, options, date1904,
            sharedStrings, workbookValidation, relationTracker);
        var entity = targetEntity ?? new TEntity();
        var errors = globalErrors?.CreateChild()
            ?? new ExcelImportErrorCollector(options.ImportOptions.ResourceLimits.MaxErrors);
        var retainResult = options.ResultMode == ExcelEntityBatchResultMode.Complete;
        if (errors.IsTruncated)
            return CreateSummary<TEntity, TItem>(entity, region, Array.Empty<int>(), retainResult,
                errors, 0, 0, 0);
        var entityItems = retainResult
            ? PrepareEntityItems<TEntity, TItem>(entity, region)
            : null;
        var sourceRows = retainResult ? new List<int>() : null;
        var rowsRead = 0;
        var succeededRows = 0;
        var failedRows = 0;
        runtime ??= new ExcelImportRuntime(options.ImportOptions.ResourceLimits);
        IDictionary<string, HashSet<string>> duplicateValues = null;
        UniqueTracker uniqueTracker = null;
        var dispatcher = new SyncBatchDispatcher<TEntity>(entity, region, options.BatchSize, onBatch);
        var headerRead = false;
        var dataRow = region.Start.Row + 1;
        var headerRow = region.Start.Row;

        if (failureCapture != null)
            failureCapture.CaptureSheetColumnMetadata(region.SheetName,
                ExcelForwardXlsxReader.ReadColumnMetadata(sheetEntry, cancellationToken));
        using var stream = sheetEntry.Open();
        var rows = ReadRows(stream, sharedStrings, styles, date1904, cancellationToken);
        foreach (var row in ExcelEntityForwardFooterBoundary.Filter(rows, region,
                     item => item.RowIndex,
                     (item, column) => item.Cells.TryGetValue(column, out var cell) ? cell : null,
                     cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.RowIndex < headerRow)
                continue;
            if (!headerRead)
            {
                if (row.RowIndex != headerRow)
                    continue;
                context.Columns = CreateColumns(context, row.Cells);
                if (UsesConfiguredValidation(options.ImportOptions.ValidationMode)
                    && context.Columns.Values.Any(column => column.IsUnique))
                {
                    duplicateValues = new Dictionary<string, HashSet<string>>(
                        StringComparer.OrdinalIgnoreCase);
                    uniqueTracker = new UniqueTracker(duplicateValues,
                        uniqueBudget?.Remaining
                            ?? options.ImportOptions.ResourceLimits.MaxTrackedUniqueValues,
                        GetStringComparer(options.ImportOptions.ResourceLimits.UniqueComparison));
                }
                failureCapture?.CaptureSheet(region.SheetName, ToFailureValues(row.Cells),
                    headerRow, ToFailureNumberFormats(row.Cells, row.StyleIndexes, styles),
                    ToFailureBoldColumns(row.StyleIndexes, styles),
                    ToFailureFontStyles(row.StyleIndexes, styles));
                headerRead = true;
                continue;
            }
            if (row.RowIndex < dataRow)
                continue;
            if (region.End.HasValue && row.RowIndex > region.End.Value.Row)
                break;
            var rowResult = ProcessRow(row, context, entityItems, sourceRows, retainResult, errors, runtime, duplicateValues,
                uniqueTracker, ref rowsRead, ref succeededRows, ref failedRows,
                options.ImportOptions, cancellationToken);
            if (rowResult.Errors.Count > 0)
                failureCapture?.CaptureErrorRow(region.SheetName, row.RowIndex + 1,
                    ToFailureValues(row.Cells), rowResult.Errors,
                    ToFailureNumberFormats(row.Cells, row.StyleIndexes, styles),
                    ToFailureBoldColumns(row.StyleIndexes, styles),
                    ToFailureFontStyles(row.StyleIndexes, styles));
            if (rowResult.HasBatch)
                dispatcher.Add(rowResult.RowIndex, rowResult.Item, rowResult.Errors);
            if (errors.IsTruncated)
            {
                errors.MarkTruncated();
                break;
            }
        }
        if (!headerRead)
            throw new BingOfficesConfigurationException("前向实体导入未匹配列表表头。",
                stage: BingOfficesStage.Plan);
        dispatcher.Flush();
        uniqueBudget?.Consume(uniqueTracker?.TrackedValueCount ?? 0);
        return CreateSummary<TEntity, TItem>(entity, region, sourceRows, retainResult, errors,
            rowsRead, succeededRows, failedRows);
    }

    /// <summary>
    /// 解析当前 Sheet，并异步交付批次。
    /// </summary>
    private async Task<ExcelEntityBatchImportSummary<TEntity>> InvokeImportRowsAsync<TEntity>(
        ZipArchiveEntry sheetEntry, IReadOnlyList<string> sharedStrings, XlsxStyleTable styles,
        bool date1904,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region,
        ExcelEntityBatchImportOptions options, ExcelXlsxWorkbookValidation workbookValidation,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken, TEntity entity = null,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelImportRuntime runtime = null,
        ExcelImportErrorCollector globalErrors = null,
        NpoiForwardUniqueBudget uniqueBudget = null,
        ExcelEntityForwardRelationTracker relationTracker = null)
        where TEntity : class, new()
    {
        var method = typeof(NpoiForwardXlsxEntityBatchImporter).GetMethod(nameof(ImportRowsAsyncTyped),
            BindingFlags.Instance | BindingFlags.NonPublic).MakeGenericMethod(typeof(TEntity), region.ItemType);
        try
        {
            var task = (Task<ExcelEntityBatchImportSummary<TEntity>>)method.Invoke(this, new object[]
            {
                sheetEntry, sharedStrings, styles, date1904, layout, region, options,
                workbookValidation, onBatch,
                cancellationToken, entity, failureCapture, runtime, globalErrors, uniqueBudget,
                relationTracker
            });
            return await task.ConfigureAwait(false);
        }
        catch (TargetInvocationException exception) when (exception.InnerException != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(exception.InnerException).Throw();
            throw;
        }
    }

    /// <summary>
    /// 解析当前 Sheet，并异步交付批次。
    /// </summary>
    private async Task<ExcelEntityBatchImportSummary<TEntity>> ImportRowsAsyncTyped<TEntity, TItem>(
        ZipArchiveEntry sheetEntry, IReadOnlyList<string> sharedStrings, XlsxStyleTable styles,
        bool date1904,
        ExcelEntityLayout<TEntity> layout, ExcelEntityListRegion<TEntity> region,
        ExcelEntityBatchImportOptions options, ExcelXlsxWorkbookValidation workbookValidation,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken, TEntity targetEntity,
        ExcelDataReaderFailureWorkbookCapture failureCapture,
        ExcelImportRuntime runtime,
        ExcelImportErrorCollector globalErrors,
        NpoiForwardUniqueBudget uniqueBudget,
        ExcelEntityForwardRelationTracker relationTracker)
        where TEntity : class, new() where TItem : class, new()
    {
        var context = CreateContext<TItem, TEntity>(layout, region, options, date1904,
            sharedStrings, workbookValidation, relationTracker);
        var entity = targetEntity ?? new TEntity();
        var errors = globalErrors?.CreateChild()
            ?? new ExcelImportErrorCollector(options.ImportOptions.ResourceLimits.MaxErrors);
        var retainResult = options.ResultMode == ExcelEntityBatchResultMode.Complete;
        if (errors.IsTruncated)
            return CreateSummary<TEntity, TItem>(entity, region, Array.Empty<int>(), retainResult,
                errors, 0, 0, 0);
        var entityItems = retainResult
            ? PrepareEntityItems<TEntity, TItem>(entity, region)
            : null;
        var sourceRows = retainResult ? new List<int>() : null;
        var rowsRead = 0;
        var succeededRows = 0;
        var failedRows = 0;
        runtime ??= new ExcelImportRuntime(options.ImportOptions.ResourceLimits);
        IDictionary<string, HashSet<string>> duplicateValues = null;
        UniqueTracker uniqueTracker = null;
        var dispatcher = new AsyncBatchDispatcher<TEntity>(entity, region, options.BatchSize, onBatch);
        var headerRead = false;
        var dataRow = region.Start.Row + 1;
        var headerRow = region.Start.Row;

        if (failureCapture != null)
            failureCapture.CaptureSheetColumnMetadata(region.SheetName,
                ExcelForwardXlsxReader.ReadColumnMetadata(sheetEntry, cancellationToken));
        using var stream = sheetEntry.Open();
        var rows = ReadRows(stream, sharedStrings, styles, date1904, cancellationToken);
        foreach (var row in ExcelEntityForwardFooterBoundary.Filter(rows, region,
                     item => item.RowIndex,
                     (item, column) => item.Cells.TryGetValue(column, out var cell) ? cell : null,
                     cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.RowIndex < headerRow)
                continue;
            if (!headerRead)
            {
                if (row.RowIndex != headerRow)
                    continue;
                context.Columns = CreateColumns(context, row.Cells);
                if (UsesConfiguredValidation(options.ImportOptions.ValidationMode)
                    && context.Columns.Values.Any(column => column.IsUnique))
                {
                    duplicateValues = new Dictionary<string, HashSet<string>>(
                        StringComparer.OrdinalIgnoreCase);
                    uniqueTracker = new UniqueTracker(duplicateValues,
                        uniqueBudget?.Remaining
                            ?? options.ImportOptions.ResourceLimits.MaxTrackedUniqueValues,
                        GetStringComparer(options.ImportOptions.ResourceLimits.UniqueComparison));
                }
                failureCapture?.CaptureSheet(region.SheetName, ToFailureValues(row.Cells),
                    headerRow, ToFailureNumberFormats(row.Cells, row.StyleIndexes, styles),
                    ToFailureBoldColumns(row.StyleIndexes, styles),
                    ToFailureFontStyles(row.StyleIndexes, styles));
                headerRead = true;
                continue;
            }
            if (row.RowIndex < dataRow)
                continue;
            if (region.End.HasValue && row.RowIndex > region.End.Value.Row)
                break;
            var rowResult = ProcessRow(row, context, entityItems, sourceRows, retainResult, errors, runtime,
                duplicateValues, uniqueTracker, ref rowsRead, ref succeededRows, ref failedRows,
                options.ImportOptions, cancellationToken);
            if (rowResult.Errors.Count > 0)
                failureCapture?.CaptureErrorRow(region.SheetName, row.RowIndex + 1,
                    ToFailureValues(row.Cells), rowResult.Errors,
                    ToFailureNumberFormats(row.Cells, row.StyleIndexes, styles),
                    ToFailureBoldColumns(row.StyleIndexes, styles),
                    ToFailureFontStyles(row.StyleIndexes, styles));
            if (rowResult.HasBatch)
                await dispatcher.AddAsync(rowResult.RowIndex, rowResult.Item, rowResult.Errors,
                    cancellationToken).ConfigureAwait(false);
            if (errors.IsTruncated)
            {
                errors.MarkTruncated();
                break;
            }
        }
        if (!headerRead)
            throw new BingOfficesConfigurationException("前向实体导入未匹配列表表头。",
                stage: BingOfficesStage.Plan);
        await dispatcher.FlushAsync(cancellationToken).ConfigureAwait(false);
        uniqueBudget?.Consume(uniqueTracker?.TrackedValueCount ?? 0);
        return CreateSummary<TEntity, TItem>(entity, region, sourceRows, retainResult, errors,
            rowsRead, succeededRows, failedRows);
    }

    /// <summary>
    /// 根据前向读取的表头创建列执行计划。
    /// </summary>
    private static IReadOnlyDictionary<int, ExcelColumnPlan> CreateColumns<TItem>(
        ForwardContext<TItem> context, IReadOnlyDictionary<int, ExcelCellValue> cells)
        where TItem : class, new()
    {
        var scopedCells = context.Options.ReadColumnRange == null
            ? cells
            : cells.Where(pair => context.Options.ReadColumnRange.Contains(pair.Key))
                .ToDictionary(pair => pair.Key, pair => pair.Value);
        if (scopedCells.Keys.Any(column => column >= context.Options.MaxReadColumns))
            throw new BingOfficesConfigurationException(
                $"Forward 实体导入表头超过最大列长度: {context.Options.MaxReadColumns}。",
                stage: BingOfficesStage.Plan);
        var headers = scopedCells.ToDictionary(pair => pair.Key,
            pair => pair.Value?.Text ?? string.Empty);
        return NpoiImportSheetExecutor.CreateColumns(headers, context.Options);
    }

    /// <summary>
    /// 在读取列表前应用直接坐标固定单元格，并收集实体级诊断错误。
    /// </summary>
    private FixedCellResult ApplyFixedCells<TEntity>(ZipArchive archive, WorkbookInfo workbook,
        IReadOnlyList<string> sharedStrings, XlsxStyleTable styles, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options, TEntity entity,
        ExcelXlsxWorkbookValidation workbookValidation, CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelImportErrorCollector globalErrors = null)
        where TEntity : class, new()
    {
        if (layout.Cells.Count == 0)
            return FixedCellResult.Empty;

        var errors = globalErrors?.CreateChild()
            ?? new ExcelImportErrorCollector(options.ImportOptions.ResourceLimits.MaxErrors);
        var resolvedCells = layout.Cells.Select(cell => cell.AnchorName == null
            ? cell
            : cell.WithResolvedReference(ResolveNamedAnchor(workbook, cell.SheetName,
                cell.AnchorName))).ToArray();
        foreach (var group in resolvedCells.GroupBy(cell => cell.SheetName,
                     StringComparer.OrdinalIgnoreCase))
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateWorkbook(workbook, group.Key);
            var sheet = workbook.FindSheet(group.Key);
            var sheetEntry = sheet == null ? null : archive.GetEntry(sheet.Path);
            if (sheetEntry == null)
                throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "NPOI",
                    stage: BingOfficesStage.Read, sheetName: group.Key);

            var bindings = group.ToArray();
            var values = ReadFixedCellValues(sheetEntry, sharedStrings, styles, workbook.Date1904,
                bindings, cancellationToken);
            var rowErrors = failureCapture == null
                ? null
                : new Dictionary<int, List<ExcelImportError>>();
            foreach (var binding in bindings)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (errors.IsLimitReached)
                {
                    errors.MarkTruncated();
                    break;
                }
                if (binding.Setter == null)
                    throw new BingOfficesConfigurationException(
                        $"实体属性不可写入: {binding.PropertyName}", stage: BingOfficesStage.Plan);

                var reference = binding.Reference;
                var cellValue = values.TryGetValue((reference.Row, reference.Column), out var value)
                    ? value
                    : new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty,
                        isDate1904: workbook.Date1904);
                var errorCount = errors.Errors.Count;
                try
                {
                    var column = NpoiEntityLayoutSupport.CreateCellPlan(
                        _importer.PlanBuilder.MappingPlanFactory, binding, MappingDirection.Import);
                    var text = NpoiExcelImporter.NormalizeText(cellValue.Text,
                        column.Property.ImportWhitespace ?? ExcelWhitespacePolicy.Trim);
                    var normalized = new ExcelCellValue(cellValue.Value, text, cellValue.Kind,
                        cellValue.CachedKind, cellValue.Formula, cellValue.ErrorCode,
                        cellValue.FormatIndex, cellValue.IsDate1904);
                    if (workbookValidation != null && !ValidateWorkbookCell(workbookValidation,
                            group.Key, reference.Row, reference.Column, cellValue.Value, text,
                            binding.PropertyName, reference, errors, options.ImportOptions,
                            cancellationToken))
                        continue;
                    if (UsesConfiguredValidation(options.ImportOptions.ValidationMode)
                        && !NpoiEntityLayoutSupport.ValidateCellImportRaw(column, normalized, text,
                            group.Key, reference, errors))
                        continue;
                    var converted = column.ConvertFrom(text, normalized, group.Key,
                        reference.Row + 1, reference.Column + 1, CultureInfo.InvariantCulture);
                    if (UsesConfiguredValidation(options.ImportOptions.ValidationMode)
                        && !NpoiEntityLayoutSupport.ValidateCellImportConverted(column, normalized,
                            text, converted, group.Key, reference, errors))
                        continue;
                    binding.Setter(entity, converted);
                }
                catch (BingOfficesException)
                {
                    throw;
                }
                catch (Exception exception) when (exception is not OperationCanceledException
                    && exception is not OutOfMemoryException && exception is not StackOverflowException)
                {
                        errors.Add(new ExcelImportError(ExcelImportErrorCode.ValueConversion,
                        exception.Message, group.Key, reference.Row + 1, reference.Column + 1,
                        binding.PropertyName, binding.PropertyName, null, cellValue.Value ?? cellValue.Text));
                }
                finally
                {
                    if (rowErrors != null && errors.Errors.Count > errorCount)
                    {
                        if (!rowErrors.TryGetValue(reference.Row, out var current))
                        {
                            current = new List<ExcelImportError>();
                            rowErrors[reference.Row] = current;
                        }
                        current.AddRange(errors.Errors.Skip(errorCount));
                    }
                }
            }
            if (rowErrors != null)
            {
                foreach (var pair in rowErrors)
                {
                    var rowCells = new Dictionary<int, ExcelCellValue>();
                    foreach (var binding in bindings.Where(item => item.Reference.Row == pair.Key))
                    {
                        rowCells[binding.Reference.Column] = values.TryGetValue(
                            (pair.Key, binding.Reference.Column), out var cell)
                            ? cell
                            : new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty,
                                isDate1904: workbook.Date1904);
                    }
                    failureCapture.CaptureErrorRow(group.Key, pair.Key + 1,
                        ToFailureValues(rowCells), pair.Value);
                }
            }
        }
        return new FixedCellResult(errors.Errors, errors.IsTruncated);
    }

    /// <summary>
    /// 校验前向固定单元格的工作簿原生规则。
    /// </summary>
    private static bool ValidateWorkbookCell(ExcelXlsxWorkbookValidation workbookValidation,
        string sheetName, int rowIndex, int columnIndex, object raw, string text,
        string propertyName, ExcelEntityCellReference reference, ExcelImportErrorCollector errors,
        ExcelEntityImportOptions importOptions, CancellationToken cancellationToken)
    {
        var valid = true;
        foreach (var result in workbookValidation.ValidateCell(sheetName, rowIndex, columnIndex,
                     raw, text, CultureInfo.InvariantCulture, cancellationToken))
        {
            if (result.IsValid)
                continue;
            errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation,
                result.Message, sheetName, reference.Row + 1, reference.Column + 1,
                propertyName, propertyName, null, raw));
            if (result.IsUnsupported
                && importOptions.UnsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report)
                continue;
            valid = false;
        }
        return valid;
    }

    /// <summary>
    /// 从指定工作表中读取固定单元格所需的最小坐标集合。
    /// </summary>
    private static IReadOnlyDictionary<(int Row, int Column), ExcelCellValue> ReadFixedCellValues<TEntity>(
        ZipArchiveEntry sheetEntry, IReadOnlyList<string> sharedStrings, XlsxStyleTable styles,
        bool date1904, IReadOnlyList<ExcelEntityCellBinding<TEntity>> bindings,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        var requested = bindings.Select(binding => (binding.Reference.Row, binding.Reference.Column))
            .ToHashSet();
        var values = new Dictionary<(int Row, int Column), ExcelCellValue>();
        var maxRow = requested.Count == 0 ? -1 : requested.Max(item => item.Row);
        using var stream = sheetEntry.Open();
        foreach (var row in ReadRows(stream, sharedStrings, styles, date1904, cancellationToken))
        {
            if (row.RowIndex > maxRow)
                break;
            foreach (var pair in row.Cells)
            {
                var key = (row.RowIndex, pair.Key);
                if (requested.Contains(key))
                    values[key] = pair.Value;
            }
            if (values.Count == requested.Count)
                break;
        }
        return values;
    }

    /// <summary>
    /// 创建前向解析上下文。
    /// </summary>
    private ForwardContext<TItem> CreateContext<TItem, TEntity>(ExcelEntityLayout<TEntity> layout,
        ExcelEntityListRegion<TEntity> region, ExcelEntityBatchImportOptions options, bool date1904,
        IReadOnlyList<string> sharedStrings, ExcelXlsxWorkbookValidation workbookValidation,
        ExcelEntityForwardRelationTracker relationTracker = null)
        where TItem : class, new() where TEntity : class, new()
    {
        var groups = region.DynamicColumnGroups ?? Array.Empty<IExcelEntityDynamicColumnGroup>();
        var groupDefinitions = groups.SelectMany(group => group.Definitions).ToArray();
        var request = ExcelImport.Workbook<ForwardRoot<TItem>>(builder => builder.Sheet(region.SheetName,
            root => root.Items, configure =>
            {
                configure.HeaderRowIndex(0).DataRowStartIndex(1);
                if (region.MappingDocument != null)
                    configure.Mapping(region.MappingDocument);
            else if (region.MappingConfiguration != null)
                    configure.Mapping(region.MappingConfiguration);
                if (groupDefinitions.Length > 0)
                    configure.DynamicColumns(CreateDynamicTargetGetter<TItem>(groups[0]), groupDefinitions);
                configure.FailOnUnknownDynamicColumns(
                    region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail);
            }));
        var resolved = new NpoiResolvedSheet(request.Sheets[0], 0, region.SheetName);
        var mapping = _importer.PlanBuilder.Create(new[] { resolved })[request.Sheets[0]];
        var dynamicColumns = mapping.Columns.Where(column => column.IsDynamicColumn).ToArray();
        var implicitDynamicProperties = groups.Count == 0
            ? CreateImplicitDynamicPropertyMap<TItem>(mapping, dynamicColumns)
            : new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        PropertyInfo dynamicProperty = null;
        if (groups.Count == 0 && dynamicColumns.Length == 1)
        {
            dynamicProperty = typeof(TItem).GetProperty(dynamicColumns[0].Name,
                BindingFlags.Instance | BindingFlags.Public);
            if (dynamicProperty == null || !dynamicProperty.CanRead
                || !typeof(IDictionary<string, object>).IsAssignableFrom(dynamicProperty.PropertyType))
                throw new BingOfficesConfigurationException(
                    $"实体动态列属性必须是可读的 IDictionary<string, object>: {dynamicColumns[0].Name}",
                    stage: BingOfficesStage.Plan);
        }
        else if (groups.Count == 0 && dynamicColumns.Length == 0 && mapping.DynamicColumns.Count > 0)
            throw new BingOfficesConfigurationException(
                "实体动态列定义没有对应的动态字典属性。", stage: BingOfficesStage.Plan);
        Func<object, object> dynamicTargetGetter = groups.Count == 0 && dynamicProperty != null
            ? item => dynamicProperty.GetValue(item) : null;
        var importOptions = new ExcelImportExecutionOptions<TItem>
        {
            MappingPlan = mapping,
            HeaderComparison = ExcelNameComparison.OrdinalIgnoreCase,
            HeaderWhitespace = ExcelWhitespacePolicy.Trim,
            BodyWhitespace = ExcelWhitespacePolicy.Trim,
            Culture = CultureInfo.InvariantCulture,
            ValidationMode = UsesConfiguredValidation(options.ImportOptions.ValidationMode)
                ? ExcelImportValidationMode.ConfiguredRules : ExcelImportValidationMode.Disabled,
            ValidationFailureMode = options.ImportOptions.ValidationFailureMode,
            ErrorSheetName = region.SheetName,
            IsDate1904 = date1904,
            MaxReadColumns = options.ImportOptions.ResourceLimits.MaxColumnsPerSheet ?? int.MaxValue,
            ReadColumnRange = region.End.HasValue
                ? ExcelReadColumnRange.Create(region.Start.Column,
                    region.End.Value.Column - region.Start.Column + 1)
                : null,
            MaxTrackedUniqueValues = options.ImportOptions.ResourceLimits.MaxTrackedUniqueValues,
            UniqueComparison = options.ImportOptions.ResourceLimits.UniqueComparison,
            DynamicTargetGetter = dynamicTargetGetter,
            DynamicTargetSetter = groups.Count > 0
                ? (item, values) => ApplyDynamicGroups(item, groups, values)
                : implicitDynamicProperties.Count == 0 ? null
                : (item, values) => ApplyImplicitDynamicValues(item, implicitDynamicProperties, values),
            DynamicPropertyNames = groups.Count == 0 ? null
                : new HashSet<string>(groups.Select(group => group.Property.Name),
                    StringComparer.OrdinalIgnoreCase),
            DynamicColumnPropertyNames = groups.Count == 0
                ? implicitDynamicProperties.ToDictionary(pair => pair.Key, pair => pair.Value.Name,
                    StringComparer.OrdinalIgnoreCase)
                : groupDefinitions.ToDictionary(definition => definition.Key,
                    definition => groups.Single(group => group.Definitions.Any(candidate =>
                        string.Equals(candidate.Key, definition.Key, StringComparison.OrdinalIgnoreCase)))
                        .Property.Name, StringComparer.OrdinalIgnoreCase),
            FailOnUnknownDynamicColumns = region.UnknownDynamicValues == ExcelUnknownDynamicValuePolicy.Fail,
            RequireExpectedHeaders = true
        };
        return new ForwardContext<TItem>(importOptions, _importer.RowMaterializer, sharedStrings,
            workbookValidation, relationTracker);
    }

    /// <summary>
    /// 创建显式动态组使用的目标属性表达式。
    /// </summary>
    private static Expression<Func<TItem, IDictionary<string, object>>> CreateDynamicTargetGetter<TItem>(
        IExcelEntityDynamicColumnGroup group) where TItem : class, new()
    {
        var parameter = Expression.Parameter(typeof(TItem), "item");
        var property = Expression.Property(parameter, group.Property);
        var body = Expression.Convert(property, typeof(IDictionary<string, object>));
        return Expression.Lambda<Func<TItem, IDictionary<string, object>>>(body, parameter);
    }

    /// <summary>
    /// 获取显式动态组的目标字典，并在需要时按属性类型创建实例。
    /// </summary>
    private static IDictionary<string, object> GetOrCreateDynamicValues(object item,
        IExcelEntityDynamicColumnGroup group)
    {
        var current = group.Getter(item);
        if (current != null)
            return current;
        if (group.Setter == null)
            throw new BingOfficesConfigurationException(
                $"实体动态列属性为空且不可写: {group.Property.Name}", stage: BingOfficesStage.Plan);
        try
        {
            object instance;
            if (group.Property.PropertyType.IsInterface || group.Property.PropertyType.IsAbstract)
            {
                if (!group.Property.PropertyType.IsAssignableFrom(typeof(Dictionary<string, object>)))
                    throw new BingOfficesConfigurationException(
                        $"实体动态列属性无法创建字典实例: {group.Property.Name}",
                        stage: BingOfficesStage.Plan);
                instance = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            }
            else
                instance = Activator.CreateInstance(group.Property.PropertyType);
            if (instance is not IDictionary<string, object> created)
                throw new BingOfficesConfigurationException(
                    $"实体动态列属性类型不受支持: {group.Property.Name}",
                    stage: BingOfficesStage.Plan);
            group.Setter(item, created);
            return created;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException(
                $"实体动态列属性无法创建字典实例: {group.Property.Name}", exception,
                BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 将前向解析得到的动态值分发到对应的实体字典组。
    /// </summary>
    private static void ApplyDynamicGroups(object item,
        IReadOnlyList<IExcelEntityDynamicColumnGroup> groups,
        IReadOnlyDictionary<string, object> values)
    {
        foreach (var group in groups)
        {
            var groupValues = values.Where(pair => group.Definitions.Any(definition =>
                    string.Equals(definition.Key, pair.Key, StringComparison.OrdinalIgnoreCase)))
                .ToArray();
            if (groupValues.Length == 0)
                continue;
            var target = GetOrCreateDynamicValues(item, group);
            foreach (var pair in groupValues)
                target[pair.Key] = pair.Value;
        }
    }

    /// <summary>
    /// 创建隐式动态列定义到实体字典属性的归属映射。
    /// </summary>
    /// <typeparam name="TItem">列表项类型。</typeparam>
    /// <param name="mapping">列表项映射计划。</param>
    /// <param name="dynamicColumns">实体声明的动态字典属性。</param>
    /// <returns>动态列 Key 到目标属性的映射。</returns>
    private static IReadOnlyDictionary<string, PropertyInfo> CreateImplicitDynamicPropertyMap<TItem>(
        IExcelMappingPlan mapping, IReadOnlyList<IExcelMappingColumn> dynamicColumns)
        where TItem : class, new()
    {
        var properties = (dynamicColumns ?? Array.Empty<IExcelMappingColumn>())
            .ToDictionary(column => column.Name,
                column => typeof(TItem).GetProperty(column.Name, BindingFlags.Instance | BindingFlags.Public),
                StringComparer.OrdinalIgnoreCase);
        foreach (var pair in properties)
        {
            var property = pair.Value;
            if (property == null || !property.CanRead || !property.CanWrite
                || !typeof(IDictionary<string, object>).IsAssignableFrom(property.PropertyType))
                throw new BingOfficesConfigurationException(
                    $"实体动态列属性必须是可读写的 IDictionary<string, object>: {pair.Key}",
                    stage: BingOfficesStage.Plan);
        }
        if (properties.Count == 0)
            return new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        if (properties.Count == 1)
        {
            var propertyName = properties.Keys.Single();
            var mismatch = (mapping.DynamicColumns ?? Array.Empty<IExcelDynamicMappingColumn>())
                .FirstOrDefault(column => !string.IsNullOrWhiteSpace(column.TargetPropertyName)
                    && !string.Equals(column.TargetPropertyName, propertyName,
                        StringComparison.OrdinalIgnoreCase));
            if (mismatch != null)
                throw new BingOfficesConfigurationException(
                    $"动态列 {mismatch.Key} 的目标属性不是实体动态字典: {mismatch.TargetPropertyName}",
                    stage: BingOfficesStage.Plan);
            return new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        }
        if (mapping.DynamicColumns == null || mapping.DynamicColumns.Count == 0)
            Unsupported("多个隐式动态字典（缺少 TargetPropertyName 归属定义）");
        var result = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in mapping.DynamicColumns)
        {
            if (string.IsNullOrWhiteSpace(column.TargetPropertyName))
                throw new BingOfficesConfigurationException(
                    $"多个动态字典属性时，动态列 {column.Key} 必须指定 TargetPropertyName。",
                    stage: BingOfficesStage.Plan);
            if (!properties.TryGetValue(column.TargetPropertyName, out var property))
                throw new BingOfficesConfigurationException(
                    $"动态列 {column.Key} 的目标属性不是实体动态字典: {column.TargetPropertyName}",
                    stage: BingOfficesStage.Plan);
            if (!result.TryAdd(column.Key, property))
                throw new BingOfficesConfigurationException(
                    $"动态列 Key 重复或归属不明确: {column.Key}", stage: BingOfficesStage.Plan);
        }
        return result;
    }

    /// <summary>
    /// 将隐式动态列值按映射计划写回多个实体字典。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="properties">动态列 Key 到目标属性的映射。</param>
    /// <param name="values">当前行转换后的动态列值。</param>
    private static void ApplyImplicitDynamicValues(object item,
        IReadOnlyDictionary<string, PropertyInfo> properties,
        IReadOnlyDictionary<string, object> values)
    {
        foreach (var pair in values)
        {
            if (!properties.TryGetValue(pair.Key, out var property))
                throw new BingOfficesConfigurationException(
                    $"动态列未找到目标字典属性: {pair.Key}", stage: BingOfficesStage.Plan);
            var target = property.GetValue(item) as IDictionary<string, object>;
            if (target == null)
            {
                if (!property.CanWrite)
                    throw new BingOfficesConfigurationException(
                        $"实体动态列属性为空且不可写: {property.Name}", stage: BingOfficesStage.Plan);
                target = CreateDynamicDictionary(item, property);
            }
            target[pair.Key] = pair.Value;
        }
    }

    /// <summary>
    /// 创建隐式动态列目标字典并写回实体属性。
    /// </summary>
    /// <param name="item">当前列表项。</param>
    /// <param name="property">目标字典属性。</param>
    /// <returns>创建的目标字典。</returns>
    private static IDictionary<string, object> CreateDynamicDictionary(object item, PropertyInfo property)
    {
        try
        {
            object instance = property.PropertyType.IsInterface || property.PropertyType.IsAbstract
                ? new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                : Activator.CreateInstance(property.PropertyType);
            if (instance is not IDictionary<string, object> created)
                throw new InvalidOperationException();
            property.SetValue(item, created);
            return created;
        }
        catch (Exception exception) when (exception is not OperationCanceledException
            && exception is not OutOfMemoryException && exception is not StackOverflowException)
        {
            throw new BingOfficesConfigurationException(
                $"实体动态列属性无法创建字典实例: {property.Name}", exception, BingOfficesStage.Plan);
        }
    }

    /// <summary>
    /// 处理一行数据并返回批次行结果。
    /// </summary>
    private static RowResult ProcessRow<TItem>(ForwardRow row, ForwardContext<TItem> context,
        ICollection<TItem> entityItems, ICollection<int> sourceRows, bool retainResult,
        ExcelImportErrorCollector errors,
        ExcelImportRuntime runtime, IDictionary<string, HashSet<string>> duplicateValues,
        UniqueTracker uniqueTracker, ref int rowsRead, ref int succeededRows, ref int failedRows,
        ExcelEntityImportOptions importOptions, CancellationToken cancellationToken)
        where TItem : class, new()
    {
        if (!runtime.TryConsumeRow())
        {
            if (runtime.TryMarkRowLimitReported())
                errors.Add(new ExcelImportError(ExcelImportErrorCode.ResourceLimit,
                    $"Workbook 数据行数超过限制: {runtime.MaxRows}", context.Options.ErrorSheetName,
                    row.RowIndex + 1, 0, null));
            return new RowResult(row.RowIndex, null, errors.Errors, false);
        }
        if (IsEmpty(row, context.Options.BodyWhitespace))
            return new RowResult(row.RowIndex, null, Array.Empty<ExcelImportError>(), false);
        rowsRead++;
        var child = errors.CreateChild();
        var configured = UsesConfiguredValidation(importOptions.ValidationMode);
        if (uniqueTracker != null)
            uniqueTracker.BeginRow();
        TItem item = null;
        var valid = true;
        if (context.WorkbookValidation != null)
            valid = ValidateWorkbookRow(row, context, importOptions, child, cancellationToken);
        if (configured)
            valid = valid && context.Materializer.ValidateRawValues(row.Cells, context.Columns, duplicateValues,
                context.Options.ErrorSheetName, row.RowIndex, importOptions.ValidationFailureMode,
                context.Options.Culture, context.Options.BodyWhitespace, child);
        if (valid)
            valid = context.Materializer.TryCreateItem(row.Cells, context.Columns, duplicateValues,
                uniqueTracker, context.Options.ErrorSheetName, row.RowIndex,
                importOptions.ValidationFailureMode, configured, child, context.Options.Culture,
                context.Options.BodyWhitespace, context.Options.DynamicTargetGetter,
                context.Options.DynamicTargetSetter, out item);
        if (valid && item != null)
        {
            if (context.RelationTracker != null)
            {
                foreach (var relationError in context.RelationTracker.Track(typeof(TItem), item,
                    context.Options.ErrorSheetName, row.RowIndex))
                    child.Add(relationError);
                if (context.RelationTracker.IsTruncated)
                    errors.MarkTruncated();
            }
            if (retainResult)
            {
                entityItems.Add(item);
                sourceRows.Add(row.RowIndex);
            }
            succeededRows++;
            if (uniqueTracker != null)
                uniqueTracker.CommitRow();
        }
        else
        {
            if (uniqueTracker != null)
                uniqueTracker.RollbackRow();
            if (child.Errors.Count > 0)
                failedRows++;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new RowResult(row.RowIndex, item, child.Errors, true);
    }

    /// <summary>
    /// 校验前向实体行的工作簿原生规则。
    /// </summary>
    private static bool ValidateWorkbookRow<TItem>(ForwardRow row, ForwardContext<TItem> context,
        ExcelEntityImportOptions importOptions, ExcelImportErrorCollector errors,
        CancellationToken cancellationToken) where TItem : class, new()
    {
        var valid = true;
        foreach (var column in context.Columns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var cell = row.Cells.TryGetValue(column.Key, out var value)
                ? value
                : new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty);
            var text = NpoiExcelImporter.NormalizeText(cell.Text, context.Options.BodyWhitespace);
            var raw = cell.Value ?? cell.Text;
            foreach (var result in context.WorkbookValidation.ValidateCell(
                         context.Options.ErrorSheetName, row.RowIndex, column.Key, raw, text,
                         context.Options.Culture, cancellationToken))
            {
                if (result.IsValid)
                    continue;
                errors.Add(new ExcelImportError(ExcelImportErrorCode.WorkbookValidation,
                    result.Message, context.Options.ErrorSheetName, row.RowIndex + 1,
                    column.Key + 1, column.Value.Property.Name,
                    column.Value.DynamicDefinition?.Key ?? column.Value.Property.Name,
                    column.Value.HeaderName, raw));
                if (result.IsUnsupported
                    && importOptions.UnsupportedFeaturePolicy == ExcelUnsupportedFeaturePolicy.Report)
                    continue;
                valid = false;
                if (importOptions.ValidationFailureMode == ExcelValidationFailureMode.StopOnFirstFailure)
                    return false;
            }
        }
        return valid;
    }

    /// <summary>
    /// 创建实体批次摘要。
    /// </summary>
    private static ExcelEntityBatchImportSummary<TEntity> CreateSummary<TEntity, TItem>(TEntity entity,
        ExcelEntityListRegion<TEntity> region, IReadOnlyList<int> sourceRows, bool hasCompleteResult,
        ExcelImportErrorCollector errors, int rowsRead, int succeededRows, int failedRows)
        where TEntity : class, new() where TItem : class, new()
    {
        var sheet = new ExcelSheetImportResult(region.SheetName, typeof(TItem),
            sourceRows ?? Array.Empty<int>(),
            errors.Errors);
        var result = new ExcelEntityImportResult<TEntity>(entity, errors.Errors, new[] { sheet },
            errors.IsTruncated, errors.MaxErrors);
        var summary = new ExcelBatchImportSummary(rowsRead, succeededRows, failedRows,
            errors.Errors.Count, errors.IsTruncated,
            errors.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit));
        var regionSummary = new ExcelEntityBatchRegionSummary(0, region.SheetName,
            region.PropertyName, typeof(TItem), summary);
        return new ExcelEntityBatchImportSummary<TEntity>(result, new[] { regionSummary },
            Array.Empty<ExcelImportError>(), hasCompleteResult);
    }

    /// <summary>
    /// 合并单个列表区域的前向结果。
    /// </summary>
    private static void AppendRegionResult<TEntity>(ExcelEntityBatchImportSummary<TEntity> result,
        int regionIndex, ICollection<ExcelEntityBatchRegionSummary> summaries,
        ICollection<ExcelSheetImportResult> sheets, ICollection<ExcelImportError> errors,
        ref bool errorsTruncated) where TEntity : class, new()
    {
        var region = result.Regions.Single();
        summaries.Add(new ExcelEntityBatchRegionSummary(regionIndex, region.SheetName,
            region.PropertyName, region.ItemType, region.Summary));
        foreach (var sheet in result.Result.Sheets)
            sheets.Add(sheet);
        foreach (var error in result.Result.Errors)
            errors.Add(error);
        errorsTruncated |= result.Result.ErrorsTruncated;
    }

    /// <summary>
    /// 创建多个工作表列表区域的聚合结果。
    /// </summary>
    private static ExcelEntityBatchImportSummary<TEntity> CreateCombinedSummary<TEntity>(TEntity entity,
        IReadOnlyList<ExcelEntityBatchRegionSummary> summaries,
        IReadOnlyList<ExcelSheetImportResult> sheets, IReadOnlyList<ExcelImportError> errors,
        bool errorsTruncated, ExcelEntityBatchImportOptions options) where TEntity : class, new()
    {
        var resultEntity = options.ResultMode == ExcelEntityBatchResultMode.Complete
            ? entity
            : new TEntity();
        var result = new ExcelEntityImportResult<TEntity>(resultEntity, errors, sheets, errorsTruncated,
            options.ImportOptions.ResourceLimits.MaxErrors);
        return new ExcelEntityBatchImportSummary<TEntity>(result, summaries,
            Array.Empty<ExcelImportError>(), options.ResultMode == ExcelEntityBatchResultMode.Complete);
    }

    /// <summary>
    /// 将批次实体集合绑定到聚合对象，并返回可追加的集合。
    /// </summary>
    private static ICollection<TItem> PrepareEntityItems<TEntity, TItem>(TEntity entity,
        ExcelEntityListRegion<TEntity> region)
        where TEntity : class, new() where TItem : class, new()
    {
        if (region.Setter != null)
        {
            var items = new List<TItem>();
            region.Setter(entity, items);
            return items;
        }
        var current = region.Getter(entity) as ICollection<TItem>;
        if (current == null)
            throw new BingOfficesConfigurationException($"实体集合属性不可写入: {region.PropertyName}",
                stage: BingOfficesStage.Plan);
        current.Clear();
        return current;
    }

    /// <summary>
    /// 校验前向解析请求的公共参数。
    /// </summary>
    private static void ValidateRequest<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options, Action<ExcelEntityImportBatch<TEntity>> onBatch)
        where TEntity : class, new()
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanRead)
            throw new ArgumentException("输入流不可读取。", nameof(source));
        if (layout == null)
            throw new ArgumentNullException(nameof(layout));
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        ValidateBatchLayout(layout, options);
        ValidateResultMode(options);
    }

    /// <summary>
    /// 校验当前前向路径支持的实体布局范围。
    /// </summary>
    internal static void ValidateLayout<TEntity>(ExcelEntityLayout<TEntity> layout,
        ExcelEntityImportOptions importOptions, bool allowMerges = true,
        bool allowRelations = true)
        where TEntity : class, new()
    {
        if (layout.ListRegions.Count == 0)
            Unsupported("空列表布局");
        foreach (var group in layout.ListRegions.GroupBy(region => region.SheetName,
                     StringComparer.OrdinalIgnoreCase))
        {
            if (group.Count() > 1 && group.Any(region => !region.End.HasValue))
                Unsupported("同一工作表中的多个列表区域必须声明结束地址");
        }
        if (!allowMerges && layout.Merges.Count > 0)
            Unsupported("合并区域");
        if (!allowRelations && layout.Relations.Count > 0)
            Unsupported("摘要结果中的关系绑定");
        if (layout.LookupLists.Count > 0 || layout.CascadeLists.Count > 0
            || layout.DataValidations.Count > 0 || layout.Protections.Count > 0)
            Unsupported("关系、模板规则或工作表元数据");
        foreach (var region in layout.ListRegions)
        {
            if (region.PageHeader != null)
                ExcelEntityPageHeaderImportValidation.ValidateStart(region);
            if (!region.IncludeHeader)
                Unsupported("无表头列表");
            if (region.GroupSubtotal != null && (region.PageSubtotal != null
                    || region.PageBreakRows.HasValue && region.PageFooter == null)
                || region.PageHeader != null && (!region.PageBreakRows.HasValue
                    || !region.End.HasValue
                    || region.Footer != null && region.PageFooter == null)
                || region.PageFooter != null && !region.PageBreakRows.HasValue
                || region.FooterAnchorName != null || region.PageHeaderAnchorPrefix != null
                || region.PageFooterAnchorPrefix != null)
                Unsupported("分页分组组合、受限重复页眉或命名锚点");
            if (region.PageSubtotal != null && !region.PageBreakRows.HasValue)
                Unsupported("缺少 PageBreak 的分页小计");
        }
    }

    /// <summary>
    /// 校验批次结果保留模式与前向解析模式的组合。
    /// </summary>
    internal static void ValidateResultMode(ExcelEntityBatchImportOptions options)
    {
        if (options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && options.ParsingMode != ExcelEntityBatchParsingMode.Forward)
            Unsupported("Buffered 模式的摘要结果");
    }

    /// <summary>
    /// 校验批次布局，并为摘要关系建立无实体保留的顺序约束。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">实体批次选项。</param>
    internal static void ValidateBatchLayout<TEntity>(ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options) where TEntity : class, new()
    {
        ValidateLayout(layout, options.ImportOptions, allowRelations: true);
        if (options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly)
            ExcelEntityForwardRelationTracker.ValidateSummaryLayout(layout);
    }

    /// <summary>
    /// 创建 HSSF 失败工作簿需要保留的固定单元格行映射。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="layout">实体布局。</param>
    /// <returns>按工作表分组的零基固定单元格行号。</returns>
    private static IReadOnlyDictionary<string, IReadOnlyList<int>> GetRequiredRowsBySheet<TEntity>(
        ExcelEntityLayout<TEntity> layout) where TEntity : class, new()
    {
        if (layout?.Cells == null || layout.Cells.Count == 0)
            return null;
        return layout.Cells
            .GroupBy(cell => cell.SheetName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key,
                group => (IReadOnlyList<int>)group.Select(cell => cell.Reference.Row)
                    .Distinct().OrderBy(row => row).ToArray(),
                StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 创建低保留失败行捕获器。
    /// </summary>
    internal static ExcelDataReaderFailureWorkbookCapture CreateFailureCapture(
        ExcelEntityBatchImportOptions options)
    {
        return options?.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && options.ImportOptions?.FailureOptions?.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly
            ? new ExcelDataReaderFailureWorkbookCapture(options.ImportOptions.FailureOptions, "NPOI",
                captureErrors: false)
            : null;
    }

    /// <summary>
    /// 判断是否使用只保留错误行的失败工作簿路径。
    /// </summary>
    internal static bool IsLowRetentionFailureOutput(ExcelEntityBatchImportOptions options,
        ExcelDataReaderFailureWorkbookCapture failureCapture)
    {
        return options?.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && options.ImportOptions?.FailureOptions?.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly
            && failureCapture != null;
    }

    /// <summary>
    /// 将前向行转换为失败工作簿需要的连续原始值数组。
    /// </summary>
    private static IReadOnlyList<object> ToFailureValues(
        IReadOnlyDictionary<int, ExcelCellValue> cells)
    {
        if (cells == null || cells.Count == 0)
            return Array.Empty<object>();
        var values = new object[cells.Keys.Max() + 1];
        foreach (var pair in cells)
            values[pair.Key] = pair.Value?.Value ?? pair.Value?.Text;
        return values;
    }

    /// <summary>
    /// 提取前向失败工作簿需要保留的数字格式。
    /// </summary>
    private static IReadOnlyDictionary<int, string> ToFailureNumberFormats(
        IReadOnlyDictionary<int, ExcelCellValue> cells,
        IReadOnlyDictionary<int, int> styleIndexes, XlsxStyleTable styles)
    {
        if (cells == null || cells.Count == 0 || styleIndexes == null
            || styleIndexes.Count == 0 || styles == null)
            return new Dictionary<int, string>();
        var formats = new Dictionary<int, string>();
        foreach (var pair in styleIndexes)
        {
            if (!cells.ContainsKey(pair.Key))
                continue;
            var format = styles.GetFormatCode(pair.Value);
            if (!string.IsNullOrWhiteSpace(format)
                && !string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
                formats[pair.Key] = format;
        }
        return formats;
    }

    /// <summary>
    /// 提取前向失败工作簿需要保留的加粗列。
    /// </summary>
    private static IReadOnlyDictionary<int, bool> ToFailureBoldColumns(
        IReadOnlyDictionary<int, int> styleIndexes, XlsxStyleTable styles)
    {
        if (styleIndexes == null || styleIndexes.Count == 0 || styles == null)
            return new Dictionary<int, bool>();
        var boldColumns = new Dictionary<int, bool>();
        foreach (var pair in styleIndexes)
        {
            if (pair.Key >= 0 && styles.IsBold(pair.Value))
                boldColumns[pair.Key] = true;
        }
        return boldColumns;
    }

    /// <summary>
    /// 提取前向失败工作簿需要保留的字体状态。
    /// </summary>
    private static IReadOnlyDictionary<int, ExcelDataReaderFailureFontStyle> ToFailureFontStyles(
        IReadOnlyDictionary<int, int> styleIndexes, XlsxStyleTable styles)
    {
        if (styleIndexes == null || styleIndexes.Count == 0 || styles == null)
            return new Dictionary<int, ExcelDataReaderFailureFontStyle>();
        var fontStyles = new Dictionary<int, ExcelDataReaderFailureFontStyle>();
        foreach (var pair in styleIndexes)
        {
            if (pair.Key < 0 || !styles.TryGetFontStyle(pair.Value, out var style))
                continue;
            if (style.IsBold || style.IsItalic || style.IsStrikeout || style.IsUnderline
                || !string.IsNullOrWhiteSpace(style.FontName) || style.FontHeightInPoints > 0
                || !string.IsNullOrWhiteSpace(style.HorizontalAlignment)
                || !string.IsNullOrWhiteSpace(style.VerticalAlignment)
                || style.WrapText || style.ShrinkToFit || style.TextRotation.HasValue
                || style.Indentation > 0 || !string.IsNullOrWhiteSpace(style.FillPattern)
                || style.BorderLeft != null || style.BorderRight != null
                || style.BorderTop != null || style.BorderBottom != null
                || style.IsLocked.HasValue || style.IsHidden.HasValue)
                fontStyles[pair.Key] = style;
        }
        return fontStyles;
    }

    /// <summary>
    /// 抛出明确的前向路径能力边界错误。
    /// </summary>
    private static void Unsupported(string feature) =>
        throw new BingOfficesConfigurationException($"NPOI Forward 实体批次暂不支持{feature}。",
            stage: BingOfficesStage.Plan);

    /// <summary>
    /// 校验前向路径的输入为 ZIP 格式 XLSX。
    /// </summary>
    private static void EnsureXlsxSignature(Stream stream)
    {
        var signature = new byte[4];
        var count = stream.Read(signature, 0, signature.Length);
        stream.Position = 0;
        if (count != signature.Length || signature[0] != 0x50 || signature[1] != 0x4B)
            Unsupported("XLS/HSSF 或非 XLSX 文件");
    }

    /// <summary>
    /// 打开可直接读取或需要暂存的输入流。
    /// </summary>
    private static ForwardInput OpenSource(Stream source, long? maxBytes,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        if (CanReadInPlace(source, maxBytes))
            return new ForwardInput(source, ownsStream: false);
        var buffer = CreateTempBuffer(temporaryStorage);
        try
        {
            NpoiStreamCopier.Copy(source, buffer, cancellationToken, maxBytes);
            buffer.Flush();
            return new ForwardInput(buffer, ownsStream: true);
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 异步打开可直接读取或需要暂存的输入流。
    /// </summary>
    private static async Task<ForwardInput> OpenSourceAsync(Stream source, long? maxBytes,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        if (CanReadInPlace(source, maxBytes))
            return new ForwardInput(source, ownsStream: false);
        var buffer = CreateTempBuffer(temporaryStorage);
        try
        {
            await NpoiStreamCopier.CopyAsync(source, buffer, cancellationToken, maxBytes)
                .ConfigureAwait(false);
            await buffer.FlushAsync(cancellationToken).ConfigureAwait(false);
            return new ForwardInput(buffer, ownsStream: true);
        }
        catch
        {
            buffer.Dispose();
            throw;
        }
    }

    /// <summary>
    /// 判断是否可以直接使用调用方的可定位源流。
    /// </summary>
    private static bool CanReadInPlace(Stream source, long? maxBytes)
    {
        if (!source.CanSeek || source.Position != 0)
            return false;
        if (maxBytes.HasValue && source.Length > maxBytes.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                provider: "NPOI", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Open);
        return true;
    }

    /// <summary>
    /// 创建由流所有权自动清理的 Forward 输入暂存文件。
    /// </summary>
    private static Stream CreateTempBuffer(ExcelTemporaryStorageBudget temporaryStorage)
    {
        var path = Path.Combine(Path.GetTempPath(),
            $"bing-offices-forward-{Guid.NewGuid():N}.xlsx");
        return (temporaryStorage ?? new ExcelTemporaryStorageBudget(null, "NPOI")).CreateFile(path,
            FileOptions.Asynchronous | FileOptions.SequentialScan | FileOptions.DeleteOnClose);
    }

    /// <summary>
    /// 保存 Forward 读取使用的输入流及其所有权。
    /// </summary>
    private sealed class ForwardInput : IDisposable, IAsyncDisposable
    {
        private readonly bool _ownsStream;

        internal ForwardInput(Stream stream, bool ownsStream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _ownsStream = ownsStream;
        }

        internal Stream Stream { get; }

        public void Dispose()
        {
            if (_ownsStream)
                Stream.Dispose();
        }

        public ValueTask DisposeAsync()
        {
            if (_ownsStream)
                return Stream.DisposeAsync();
            return ValueTask.CompletedTask;
        }
    }

    /// <summary>
    /// 判断是否执行配置校验。
    /// </summary>
    private static bool UsesConfiguredValidation(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.ConfiguredRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 判断是否执行工作簿原生校验。
    /// </summary>
    private static bool UsesWorkbookValidation(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.WorkbookRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 判断前向解析行是否为空。
    /// </summary>
    private static bool IsEmpty(ForwardRow row, ExcelWhitespacePolicy whitespace) =>
        row.Cells.Count == 0 || row.Cells.Values.All(cell =>
            string.IsNullOrWhiteSpace(NpoiExcelImporter.NormalizeText(cell?.Text, whitespace)));

    /// <summary>
    /// 将公共唯一值比较规则转换为字符串比较器。
    /// </summary>
    private static IEqualityComparer<string> GetStringComparer(StringComparison comparison)
    {
        switch (comparison)
        {
            case StringComparison.Ordinal:
                return StringComparer.Ordinal;
            case StringComparison.OrdinalIgnoreCase:
                return StringComparer.OrdinalIgnoreCase;
            case StringComparison.InvariantCulture:
                return StringComparer.InvariantCulture;
            case StringComparison.InvariantCultureIgnoreCase:
                return StringComparer.InvariantCultureIgnoreCase;
            case StringComparison.CurrentCulture:
                return StringComparer.CurrentCulture;
            case StringComparison.CurrentCultureIgnoreCase:
                return StringComparer.CurrentCultureIgnoreCase;
            default:
                throw new ArgumentOutOfRangeException(nameof(comparison));
        }
    }

    /// <summary>
    /// 读取工作簿关系和 Sheet 信息。
    /// </summary>
    private static WorkbookInfo ReadWorkbook(ZipArchive archive, CancellationToken cancellationToken)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new BingOfficesImportException("XLSX 缺少 workbook.xml。", provider: "NPOI",
                stage: BingOfficesStage.Preflight);
        var relationEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new BingOfficesImportException("XLSX 缺少工作簿关系部件。", provider: "NPOI",
                stage: BingOfficesStage.Preflight);
        var sheets = new List<(string Name, string RelationId)>();
        var definedNames = new List<WorkbookDefinedName>();
        var date1904 = false;
        using (var stream = workbookEntry.Open())
        using (var reader = CreateReader(stream))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element)
                    continue;
                if (string.Equals(reader.LocalName, "workbookPr", StringComparison.Ordinal))
                {
                    var value = reader.GetAttribute("date1904");
                    date1904 = string.Equals(value, "1", StringComparison.Ordinal)
                        || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
                }
                else if (string.Equals(reader.LocalName, "sheet", StringComparison.Ordinal))
                {
                    var id = reader.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                        ?? reader.GetAttribute("r:id");
                    sheets.Add((reader.GetAttribute("name"), id));
                }
                else if (string.Equals(reader.LocalName, "definedName", StringComparison.Ordinal))
                {
                    var name = reader.GetAttribute("name");
                    var localSheetId = -1;
                    var localSheetText = reader.GetAttribute("localSheetId");
                    if (!string.IsNullOrWhiteSpace(localSheetText)
                        && (!int.TryParse(localSheetText, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out localSheetId)
                            || localSheetId < 0))
                        localSheetId = -2;
                    var formula = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(name))
                        definedNames.Add(new WorkbookDefinedName(name, localSheetId, formula));
                }
            }
        }
        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        using (var stream = relationEntry.Open())
        using (var reader = CreateReader(stream))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element
                    || !string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal))
                    continue;
                var id = reader.GetAttribute("Id");
                var target = reader.GetAttribute("Target");
                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                    relationships[id] = NormalizePartPath(target);
            }
        }
        if (sheets.Count == 0)
            throw new BingOfficesImportException("XLSX 未声明工作表。", provider: "NPOI",
                stage: BingOfficesStage.Preflight);
        var workbookSheets = new List<WorkbookSheetInfo>(sheets.Count);
        foreach (var sheet in sheets)
        {
            if (string.IsNullOrWhiteSpace(sheet.Name) || string.IsNullOrWhiteSpace(sheet.RelationId)
                || !relationships.TryGetValue(sheet.RelationId, out var path))
                throw new BingOfficesImportException("XLSX 工作表关系无效。", provider: "NPOI",
                    stage: BingOfficesStage.Preflight, sheetName: sheet.Name);
            workbookSheets.Add(new WorkbookSheetInfo(sheet.Name, path));
        }
        return new WorkbookInfo(workbookSheets, date1904, definedNames);
    }

    /// <summary>
    /// 读取当前元素的文本内容并保持父级读取器在当前元素上。
    /// </summary>
    private static string ReadElementText(XmlReader reader)
    {
        var builder = new StringBuilder();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Text || subtree.NodeType == XmlNodeType.CDATA)
                builder.Append(subtree.Value);
        }
        return builder.ToString();
    }

    /// <summary>
    /// 解析前向工作簿中的单格命名锚点。
    /// </summary>
    private static ExcelEntityCellReference ResolveNamedAnchor(WorkbookInfo workbook,
        string sheetName, string anchorName)
    {
        var sheetIndex = workbook.GetSheetIndex(sheetName);
        if (sheetIndex < 0)
            throw new BingOfficesConfigurationException(
                $"Forward 实体布局要求 Sheet {sheetName}，工作簿中不存在该工作表。",
                stage: BingOfficesStage.Plan);
        var candidates = workbook.DefinedNames.Where(item =>
            string.Equals(item.Name, anchorName, StringComparison.OrdinalIgnoreCase)).ToArray();
        var scoped = candidates.Where(item => item.LocalSheetId == sheetIndex).ToArray();
        var global = candidates.Where(item => item.LocalSheetId == -1).ToArray();
        var matches = scoped.Length > 0 ? scoped : global;
        if (matches.Length == 0)
            throw new BingOfficesConfigurationException(
                $"模板缺少命名锚点: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        if (matches.Length > 1)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点不唯一: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);

        var match = ForwardNamedAnchorFormula.Match(matches[0].Formula ?? string.Empty);
        if (!match.Success)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个有界单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        var targetSheet = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'")
            : match.Groups["plain"].Success ? match.Groups["plain"].Value : sheetName;
        if (!match.Groups["quoted"].Success && !match.Groups["plain"].Success
            && matches[0].LocalSheetId != sheetIndex)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点缺少 Sheet 引用: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (!string.Equals(targetSheet, sheetName, StringComparison.OrdinalIgnoreCase))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点引用了其他 Sheet: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (match.Groups["endColumn"].Success
            && (!string.Equals(match.Groups["startColumn"].Value,
                    match.Groups["endColumn"].Value, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(match.Groups["startRow"].Value,
                    match.Groups["endRow"].Value, StringComparison.Ordinal)))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        try
        {
            return ExcelEntityCellReference.Parse(match.Groups["startColumn"].Value
                + match.Groups["startRow"].Value);
        }
        catch (ArgumentException exception)
        {
            throw new BingOfficesConfigurationException(
                $"模板命名锚点地址超出工作表边界: {sheetName}!{anchorName}", exception,
                BingOfficesStage.Plan);
        }
    }

    private static readonly Regex ForwardNamedAnchorFormula = new Regex(
        @"^\s*=?\s*(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!]+))?!\$?(?<startColumn>[A-Za-z]{1,3})\$?(?<startRow>[1-9][0-9]*)(?::\$?(?<endColumn>[A-Za-z]{1,3})\$?(?<endRow>[1-9][0-9]*))?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 校验工作簿 Sheet 与布局目标一致。
    /// </summary>
    private static void ValidateWorkbook(WorkbookInfo workbook, string sheetName)
    {
        if (workbook.FindSheet(sheetName) == null)
            throw new BingOfficesConfigurationException(
                $"Forward 实体布局要求 Sheet {sheetName}，工作簿中不存在该工作表。",
                stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 校验 Forward 布局声明的合并区域存在于源 worksheet。
    /// </summary>
    private static void ValidateForwardMerges<TEntity>(ZipArchive archive, WorkbookInfo workbook,
        ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        if (layout.Merges.Count == 0)
            return;

        foreach (var group in layout.Merges.GroupBy(item => item.SheetName,
                     StringComparer.OrdinalIgnoreCase))
        {
            var sheet = workbook.FindSheet(group.Key);
            if (sheet == null)
                throw new BingOfficesConfigurationException(
                    $"Forward 实体布局要求合并区域所在 Sheet {group.Key}，工作簿中不存在该工作表。",
                    stage: BingOfficesStage.Plan);
            var entry = archive.GetEntry(sheet.Path);
            if (entry == null)
                throw new BingOfficesImportException("XLSX 合并区域所在工作表部件不存在。",
                    provider: "NPOI", stage: BingOfficesStage.Read, sheetName: sheet.Name);
            var actual = ExcelForwardXlsxReader.ReadMergeRanges(entry, cancellationToken);
            foreach (var requested in group)
            {
                if (actual.Contains(requested.Range))
                    continue;
                throw new BingOfficesConfigurationException(
                    $"Forward 输入缺少布局声明的合并区域: {requested.SheetName}!{requested.Range.Address}",
                    stage: BingOfficesStage.Plan, sheetName: requested.SheetName,
                    rowIndex: requested.Range.First.Row + 1,
                    columnIndex: requested.Range.First.Column + 1);
            }
        }
    }

    /// <summary>
    /// 读取共享字符串表。
    /// </summary>
    private static ExcelForwardXlsxSharedStrings ReadSharedStrings(ZipArchive archive,
        CancellationToken cancellationToken, bool diskBacked,
        ExcelTemporaryStorageBudget temporaryStorage)
        => ExcelForwardXlsxReader.ReadSharedStrings(archive, cancellationToken, diskBacked, temporaryStorage);

    /// <summary>
    /// 读取 XLSX 单元格样式与数字格式。
    /// </summary>
    private static XlsxStyleTable ReadStyles(ZipArchive archive, CancellationToken cancellationToken)
    {
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null)
            return XlsxStyleTable.Empty;

        var customFormats = new Dictionary<int, string>();
        var cellFormatIds = new List<int>();
        var cellFontIds = new List<int>();
        var cellFillIds = new List<int>();
        var cellBorderIds = new List<int>();
        var cellAlignments = new List<XlsxAlignment>();
        var cellProtections = new List<XlsxProtection>();
        var fontBold = new List<bool>();
        var fontItalic = new List<bool>();
        var fontStrikeout = new List<bool>();
        var fontUnderline = new List<bool>();
        var fontNames = new List<string>();
        var fontHeights = new List<double>();
        var fontColorIndexes = new List<short?>();
        var fills = new List<XlsxFill>();
        var borders = new List<XlsxBorder>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        var numberFormatsDepth = -1;
        var cellFormatsDepth = -1;
        var fillsDepth = -1;
        var fillDepth = -1;
        var patternFillDepth = -1;
        var bordersDepth = -1;
        var borderDepth = -1;
        var borderSideDepth = -1;
        var fontsDepth = -1;
        var fontDepth = -1;
        var currentFontBold = false;
        var currentFontItalic = false;
        var currentFontStrikeout = false;
        var currentFontUnderline = false;
        var currentFontName = (string)null;
        var currentFontHeight = 0d;
        short? currentFontColorIndex = null;
        string currentFillPattern = null;
        short? currentFillForeground = null;
        short? currentFillBackground = null;
        var currentBorderLeft = (XlsxBorderSide)null;
        var currentBorderRight = (XlsxBorderSide)null;
        var currentBorderTop = (XlsxBorderSide)null;
        var currentBorderBottom = (XlsxBorderSide)null;
        string currentBorderSideName = null;
        string currentBorderSideStyle = null;
        short? currentBorderSideColor = null;
        var cellFormatDepth = -1;
        XlsxAlignment currentAlignment = null;
        XlsxProtection currentProtection = null;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (string.Equals(reader.LocalName, "numFmts", StringComparison.Ordinal))
                {
                    numberFormatsDepth = reader.Depth;
                }
                else if (string.Equals(reader.LocalName, "fonts", StringComparison.Ordinal))
                {
                    fontsDepth = reader.Depth;
                }
                else if (string.Equals(reader.LocalName, "borders", StringComparison.Ordinal))
                {
                    bordersDepth = reader.Depth;
                }
                else if (bordersDepth >= 0 && reader.Depth == bordersDepth + 1
                    && string.Equals(reader.LocalName, "border", StringComparison.Ordinal))
                {
                    borderDepth = reader.Depth;
                    currentBorderLeft = null;
                    currentBorderRight = null;
                    currentBorderTop = null;
                    currentBorderBottom = null;
                    if (reader.IsEmptyElement)
                    {
                        borders.Add(new XlsxBorder(null, null, null, null));
                        borderDepth = -1;
                    }
                }
                else if (borderDepth >= 0 && reader.Depth == borderDepth + 1
                    && (string.Equals(reader.LocalName, "left", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "right", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "top", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "bottom", StringComparison.Ordinal)))
                {
                    borderSideDepth = reader.Depth;
                    currentBorderSideName = reader.LocalName;
                    currentBorderSideStyle = reader.GetAttribute("style");
                    currentBorderSideColor = null;
                    if (reader.IsEmptyElement)
                    {
                        SetBorderSide(ref currentBorderLeft, ref currentBorderRight,
                            ref currentBorderTop, ref currentBorderBottom,
                            currentBorderSideName,
                            new XlsxBorderSide(currentBorderSideStyle, null));
                        borderSideDepth = -1;
                    }
                }
                else if (borderSideDepth >= 0 && reader.Depth == borderSideDepth + 1
                    && string.Equals(reader.LocalName, "color", StringComparison.Ordinal))
                {
                    currentBorderSideColor = TryParseShort(reader.GetAttribute("indexed"));
                }
                else if (string.Equals(reader.LocalName, "fills", StringComparison.Ordinal))
                {
                    fillsDepth = reader.Depth;
                }
                else if (fillsDepth >= 0 && reader.Depth == fillsDepth + 1
                    && string.Equals(reader.LocalName, "fill", StringComparison.Ordinal))
                {
                    fillDepth = reader.Depth;
                    currentFillPattern = null;
                    currentFillForeground = null;
                    currentFillBackground = null;
                    if (reader.IsEmptyElement)
                    {
                        fills.Add(null);
                        fillDepth = -1;
                    }
                }
                else if (fillDepth >= 0 && reader.Depth == fillDepth + 1
                    && string.Equals(reader.LocalName, "patternFill", StringComparison.Ordinal))
                {
                    patternFillDepth = reader.Depth;
                    currentFillPattern = reader.GetAttribute("patternType");
                }
                else if (patternFillDepth >= 0 && reader.Depth == patternFillDepth + 1
                    && string.Equals(reader.LocalName, "fgColor", StringComparison.Ordinal))
                {
                    currentFillForeground = TryParseShort(reader.GetAttribute("indexed"));
                }
                else if (patternFillDepth >= 0 && reader.Depth == patternFillDepth + 1
                    && string.Equals(reader.LocalName, "bgColor", StringComparison.Ordinal))
                {
                    currentFillBackground = TryParseShort(reader.GetAttribute("indexed"));
                }
                else if (fontsDepth >= 0 && reader.Depth == fontsDepth + 1
                    && string.Equals(reader.LocalName, "font", StringComparison.Ordinal))
                {
                    fontDepth = reader.Depth;
                    currentFontBold = false;
                    currentFontItalic = false;
                    currentFontStrikeout = false;
                    currentFontUnderline = false;
                    currentFontName = null;
                    currentFontHeight = 0;
                    currentFontColorIndex = null;
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "b", StringComparison.Ordinal))
                {
                    currentFontBold = !string.Equals(reader.GetAttribute("val"), "0",
                        StringComparison.Ordinal);
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "i", StringComparison.Ordinal))
                {
                    currentFontItalic = !string.Equals(reader.GetAttribute("val"), "0",
                        StringComparison.Ordinal);
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "strike", StringComparison.Ordinal))
                {
                    currentFontStrikeout = !string.Equals(reader.GetAttribute("val"), "0",
                        StringComparison.Ordinal);
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "u", StringComparison.Ordinal))
                {
                    var underline = reader.GetAttribute("val");
                    currentFontUnderline = string.IsNullOrWhiteSpace(underline)
                        || string.Equals(underline, "single", StringComparison.OrdinalIgnoreCase);
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "name", StringComparison.Ordinal))
                {
                    currentFontName = reader.GetAttribute("val");
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "sz", StringComparison.Ordinal))
                {
                    if (double.TryParse(reader.GetAttribute("val"), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out var height)
                        && !double.IsNaN(height) && !double.IsInfinity(height) && height > 0)
                        currentFontHeight = height;
                }
                else if (fontDepth >= 0 && reader.Depth == fontDepth + 1
                    && string.Equals(reader.LocalName, "color", StringComparison.Ordinal))
                {
                    if (short.TryParse(reader.GetAttribute("indexed"), NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out var colorIndex) && colorIndex >= 0)
                        currentFontColorIndex = colorIndex;
                }
                else if (numberFormatsDepth >= 0 && reader.Depth == numberFormatsDepth + 1
                    && string.Equals(reader.LocalName, "numFmt", StringComparison.Ordinal))
                {
                    if (TryParseInt(reader.GetAttribute("numFmtId"), out var formatId))
                        customFormats[formatId] = reader.GetAttribute("formatCode") ?? string.Empty;
                }
                else if (string.Equals(reader.LocalName, "cellXfs", StringComparison.Ordinal))
                {
                    cellFormatsDepth = reader.Depth;
                }
                else if (cellFormatsDepth >= 0 && reader.Depth == cellFormatsDepth + 1
                    && string.Equals(reader.LocalName, "xf", StringComparison.Ordinal))
                {
                    cellFormatIds.Add(TryParseInt(reader.GetAttribute("numFmtId"), out var formatId)
                        ? formatId : 0);
                    cellFontIds.Add(TryParseInt(reader.GetAttribute("fontId"), out var fontId)
                        ? fontId : 0);
                    cellFillIds.Add(TryParseInt(reader.GetAttribute("fillId"), out var fillId)
                        ? fillId : 0);
                    cellBorderIds.Add(TryParseInt(reader.GetAttribute("borderId"), out var borderId)
                        ? borderId : 0);
                    cellFormatDepth = reader.Depth;
                    currentAlignment = null;
                    currentProtection = null;
                    if (reader.IsEmptyElement)
                    {
                        cellAlignments.Add(null);
                        cellProtections.Add(null);
                        cellFormatDepth = -1;
                    }
                }
                else if (cellFormatDepth >= 0 && reader.Depth == cellFormatDepth + 1
                    && string.Equals(reader.LocalName, "alignment", StringComparison.Ordinal))
                {
                    currentAlignment = new XlsxAlignment(
                        reader.GetAttribute("horizontal"), reader.GetAttribute("vertical"),
                        IsXmlTrue(reader.GetAttribute("wrapText")),
                        IsXmlTrue(reader.GetAttribute("shrinkToFit")),
                        TryParseShort(reader.GetAttribute("textRotation")),
                        TryParseShort(reader.GetAttribute("indent")) ?? 0);
                }
                else if (cellFormatDepth >= 0 && reader.Depth == cellFormatDepth + 1
                    && string.Equals(reader.LocalName, "protection", StringComparison.Ordinal))
                {
                    currentProtection = new XlsxProtection(
                        TryParseOptionalBool(reader.GetAttribute("locked")),
                        TryParseOptionalBool(reader.GetAttribute("hidden")));
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == fontDepth
                    && string.Equals(reader.LocalName, "font", StringComparison.Ordinal))
                {
                    fontBold.Add(currentFontBold);
                    fontItalic.Add(currentFontItalic);
                    fontStrikeout.Add(currentFontStrikeout);
                    fontUnderline.Add(currentFontUnderline);
                    fontNames.Add(currentFontName);
                    fontHeights.Add(currentFontHeight);
                    fontColorIndexes.Add(currentFontColorIndex);
                    fontDepth = -1;
                }
                if (reader.Depth == fillDepth
                    && string.Equals(reader.LocalName, "fill", StringComparison.Ordinal))
                {
                    fills.Add(new XlsxFill(currentFillPattern, currentFillForeground,
                        currentFillBackground));
                    fillDepth = -1;
                    currentFillPattern = null;
                    currentFillForeground = null;
                    currentFillBackground = null;
                }
                if (reader.Depth == borderSideDepth
                    && (string.Equals(reader.LocalName, "left", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "right", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "top", StringComparison.Ordinal)
                        || string.Equals(reader.LocalName, "bottom", StringComparison.Ordinal)))
                {
                    SetBorderSide(ref currentBorderLeft, ref currentBorderRight,
                        ref currentBorderTop, ref currentBorderBottom,
                        currentBorderSideName,
                        new XlsxBorderSide(currentBorderSideStyle, currentBorderSideColor));
                    borderSideDepth = -1;
                    currentBorderSideName = null;
                    currentBorderSideStyle = null;
                    currentBorderSideColor = null;
                }
                if (reader.Depth == borderDepth
                    && string.Equals(reader.LocalName, "border", StringComparison.Ordinal))
                {
                    borders.Add(new XlsxBorder(currentBorderLeft, currentBorderRight,
                        currentBorderTop, currentBorderBottom));
                    borderDepth = -1;
                }
                if (reader.Depth == patternFillDepth
                    && string.Equals(reader.LocalName, "patternFill", StringComparison.Ordinal))
                    patternFillDepth = -1;
                if (reader.Depth == cellFormatDepth
                    && string.Equals(reader.LocalName, "xf", StringComparison.Ordinal))
                {
                    cellAlignments.Add(currentAlignment);
                    cellProtections.Add(currentProtection);
                    currentAlignment = null;
                    currentProtection = null;
                    cellFormatDepth = -1;
                }
                if (reader.Depth == numberFormatsDepth
                    && string.Equals(reader.LocalName, "numFmts", StringComparison.Ordinal))
                    numberFormatsDepth = -1;
                if (reader.Depth == fontsDepth
                    && string.Equals(reader.LocalName, "fonts", StringComparison.Ordinal))
                    fontsDepth = -1;
                if (reader.Depth == fillsDepth
                    && string.Equals(reader.LocalName, "fills", StringComparison.Ordinal))
                    fillsDepth = -1;
                if (reader.Depth == bordersDepth
                    && string.Equals(reader.LocalName, "borders", StringComparison.Ordinal))
                    bordersDepth = -1;
                if (reader.Depth == cellFormatsDepth
                    && string.Equals(reader.LocalName, "cellXfs", StringComparison.Ordinal))
                    cellFormatsDepth = -1;
            }
        }
        return new XlsxStyleTable(cellFormatIds, cellFontIds, cellFillIds, fills,
            cellBorderIds, borders, fontBold, fontItalic,
            fontStrikeout, fontUnderline, fontNames, fontHeights, fontColorIndexes,
            cellAlignments, cellProtections, customFormats);
    }

    /// <summary>
    /// 前向读取 worksheet 行。
    /// </summary>
    private static IEnumerable<ForwardRow> ReadRows(Stream stream, IReadOnlyList<string> sharedStrings,
        XlsxStyleTable styles, bool date1904, CancellationToken cancellationToken)
    {
        using var reader = CreateReader(stream);
        var nextRow = 0;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "row", StringComparison.Ordinal))
                continue;
            var rowIndex = ParseRowIndex(reader.GetAttribute("r"), nextRow);
            nextRow = rowIndex + 1;
            var cells = ReadRowCells(reader, sharedStrings, styles, date1904,
                cancellationToken, out var styleIndexes);
            yield return new ForwardRow(rowIndex, cells, styleIndexes);
        }
    }

    /// <summary>
    /// 读取当前 XML row 的单元格集合。
    /// </summary>
    private static IReadOnlyDictionary<int, ExcelCellValue> ReadRowCells(XmlReader reader,
        IReadOnlyList<string> sharedStrings, XlsxStyleTable styles, bool date1904,
        CancellationToken cancellationToken, out IReadOnlyDictionary<int, int> styleIndexes)
    {
        var cells = new Dictionary<int, ExcelCellValue>();
        var capturedStyles = new Dictionary<int, int>();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (subtree.NodeType != XmlNodeType.Element
                || !string.Equals(subtree.LocalName, "c", StringComparison.Ordinal))
                continue;
            var cell = ReadCell(subtree, sharedStrings, styles, date1904);
            if (cell.ColumnIndex >= 0)
            {
                cells[cell.ColumnIndex] = cell.Value;
                if (cell.StyleIndex > 0)
                    capturedStyles[cell.ColumnIndex] = cell.StyleIndex;
            }
        }
        styleIndexes = capturedStyles;
        return cells;
    }

    /// <summary>
    /// 读取单个 worksheet 单元格。
    /// </summary>
    private static ForwardCell ReadCell(XmlReader reader, IReadOnlyList<string> sharedStrings,
        XlsxStyleTable styles, bool date1904)
    {
        var reference = reader.GetAttribute("r");
        var type = reader.GetAttribute("t");
        var styleIndex = TryParseInt(reader.GetAttribute("s"), out var parsedStyleIndex)
            ? parsedStyleIndex : 0;
        var value = (string)null;
        var formula = (string)null;
        var inline = new StringBuilder();
        using (var subtree = reader.ReadSubtree())
        {
            while (subtree.Read())
            {
                if (subtree.NodeType != XmlNodeType.Element)
                    continue;
                if (string.Equals(subtree.LocalName, "f", StringComparison.Ordinal))
                    formula = subtree.ReadElementContentAsString();
                else if (string.Equals(subtree.LocalName, "v", StringComparison.Ordinal))
                    value = subtree.ReadElementContentAsString();
                else if (string.Equals(subtree.LocalName, "t", StringComparison.Ordinal))
                    inline.Append(subtree.ReadElementContentAsString());
            }
        }
        if (!string.IsNullOrWhiteSpace(formula))
            Unsupported("公式单元格");
        var columnIndex = ParseColumnIndex(reference);
        return new ForwardCell(columnIndex, CreateCellValue(type, value, inline.ToString(),
            sharedStrings, styles, styleIndex, date1904), styleIndex);
    }

    /// <summary>
    /// 将 OOXML 单元格文本转换为公共单元格值描述。
    /// </summary>
    private static ExcelCellValue CreateCellValue(string type, string value, string inline,
        IReadOnlyList<string> sharedStrings, XlsxStyleTable styles, int styleIndex, bool date1904)
    {
        if (string.Equals(type, "s", StringComparison.Ordinal))
        {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                || index < 0 || index >= sharedStrings.Count)
                throw new BingOfficesImportException("XLSX shared string 索引无效。", provider: "NPOI",
                    stage: BingOfficesStage.Read);
            var text = sharedStrings[index] ?? string.Empty;
            return new ExcelCellValue(text, text, ExcelCellKind.Text, isDate1904: date1904);
        }
        if (string.Equals(type, "inlineStr", StringComparison.Ordinal)
            || string.Equals(type, "str", StringComparison.Ordinal))
            return new ExcelCellValue(inline ?? value ?? string.Empty, inline ?? value ?? string.Empty,
                ExcelCellKind.Text, isDate1904: date1904);
        if (string.Equals(type, "b", StringComparison.Ordinal))
        {
            var boolean = string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
            return new ExcelCellValue(boolean, boolean ? "TRUE" : "FALSE", ExcelCellKind.Boolean,
                isDate1904: date1904);
        }
        if (string.Equals(type, "e", StringComparison.Ordinal))
            return new ExcelCellValue(null, value ?? string.Empty, ExcelCellKind.Error,
                isDate1904: date1904);
        if (string.IsNullOrEmpty(value))
            return new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty, isDate1904: date1904);
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return new ExcelCellValue(value, value, ExcelCellKind.Text, isDate1904: date1904);
        var formatIndex = styles.GetFormatIndex(styleIndex);
        if (!styles.TryFormat(number, styleIndex, date1904, out var formatted, out var isDate))
            return new ExcelCellValue(number, value, ExcelCellKind.Number, formatIndex: formatIndex,
                isDate1904: date1904);
        if (isDate && TryParseDate(number, date1904, out var date))
            return new ExcelCellValue(date, formatted, ExcelCellKind.DateTime, formatIndex: formatIndex,
                isDate1904: date1904);
        // 数值文本仍保留 XML 原文，避免千分位格式影响 decimal、整数等默认转换；格式索引由上下文保留。
        return new ExcelCellValue(number, value, ExcelCellKind.Number, formatIndex: formatIndex,
            isDate1904: date1904);
    }

    /// <summary>
    /// 按工作簿日期系统解析数值日期。
    /// </summary>
    private static bool TryParseDate(double serial, bool date1904, out DateTime date)
    {
        var text = serial.ToString(CultureInfo.InvariantCulture);
        var cell = new ExcelCellValue(serial, text, ExcelCellKind.DateTime,
            isDate1904: date1904);
        if (new DateTimeExcelValidationRule().TryParseValue(cell, text, typeof(DateTime),
            CultureInfo.InvariantCulture, null, out var parsed) && parsed is DateTime value)
        {
            date = value;
            return true;
        }
        date = default;
        return false;
    }

    /// <summary>
    /// 创建启用安全设置的 XML 读取器。
    /// </summary>
    private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0,
        MaxCharactersInDocument = 0
    });

    /// <summary>
    /// 将 Sheet 关系目标规范化为 ZIP 部件路径。
    /// </summary>
    private static string NormalizePartPath(string target)
    {
        var value = (target ?? string.Empty).Replace('\\', '/');
        while (value.StartsWith("../", StringComparison.Ordinal))
            value = value.Substring(3);
        value = value.TrimStart('/');
        return value.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? value : "xl/" + value;
    }

    /// <summary>
    /// 解析 worksheet 单元格的一基列号为零基索引。
    /// </summary>
    private static int ParseColumnIndex(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return -1;
        var column = 0;
        foreach (var character in reference)
        {
            if (character < 'A' && character < 'a' || character > 'Z' && character > 'z')
                break;
            var upper = char.ToUpperInvariant(character);
            column = checked(column * 26 + upper - 'A' + 1);
        }
        return column - 1;
    }

    /// <summary>
    /// 解析 worksheet 行号。
    /// </summary>
    private static int ParseRowIndex(string reference, int fallback)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return fallback;
        var index = 0;
        while (index < reference.Length && !char.IsDigit(reference[index]))
            index++;
        return index < reference.Length && int.TryParse(reference.Substring(index),
            NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)
            ? Math.Max(0, row - 1) : fallback;
    }

    /// <summary>
    /// 解析非负整数属性。
    /// </summary>
    private static bool TryParseInt(string value, out int result) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result)
        && result >= 0;

    /// <summary>
    /// 解析非负短整数属性。
    /// </summary>
    private static short? TryParseShort(string value)
        => short.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture,
            out var result) && result >= 0 ? result : null;

    /// <summary>
    /// 解析 XLSX 布尔属性。
    /// </summary>
    private static bool IsXmlTrue(string value)
        => string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 解析可选的 XLSX 布尔属性。
    /// </summary>
    private static bool? TryParseOptionalBool(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;
        if (string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.Equals(value, "0", StringComparison.Ordinal)
            || string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
            return false;
        return null;
    }

    /// <summary>
    /// 保存当前解析到的标准边框一侧。
    /// </summary>
    private static void SetBorderSide(ref XlsxBorderSide left, ref XlsxBorderSide right,
        ref XlsxBorderSide top, ref XlsxBorderSide bottom, string name, XlsxBorderSide side)
    {
        switch (name)
        {
            case "left":
                left = side;
                break;
            case "right":
                right = side;
                break;
            case "top":
                top = side;
                break;
            case "bottom":
                bottom = side;
                break;
        }
    }

    /// <summary>
    /// 保存 XLSX 单元格的最小对齐状态。
    /// </summary>
    private sealed class XlsxAlignment
    {
        /// <summary>
        /// 初始化一个 XLSX 对齐状态描述。
        /// </summary>
        internal XlsxAlignment(string horizontal, string vertical, bool wrapText,
            bool shrinkToFit, short? rotation, short indent)
        {
            Horizontal = horizontal;
            Vertical = vertical;
            WrapText = wrapText;
            ShrinkToFit = shrinkToFit;
            Rotation = rotation;
            Indent = indent;
        }

        /// <summary>
        /// 获取水平对齐名称。
        /// </summary>
        internal string Horizontal { get; }

        /// <summary>
        /// 获取垂直对齐名称。
        /// </summary>
        internal string Vertical { get; }

        /// <summary>
        /// 获取是否自动换行。
        /// </summary>
        internal bool WrapText { get; }

        /// <summary>
        /// 获取是否缩小字体适应单元格。
        /// </summary>
        internal bool ShrinkToFit { get; }

        /// <summary>
        /// 获取文本旋转角度。
        /// </summary>
        internal short? Rotation { get; }

        /// <summary>
        /// 获取文本缩进量。
        /// </summary>
        internal short Indent { get; }
    }

    /// <summary>
    /// 保存 XLSX 单元格的最小填充状态。
    /// </summary>
    private sealed class XlsxFill
    {
        /// <summary>
        /// 初始化一个 XLSX 填充状态描述。
        /// </summary>
        internal XlsxFill(string pattern, short? foreground, short? background)
        {
            Pattern = string.IsNullOrWhiteSpace(pattern) ? null : pattern.Trim();
            Foreground = foreground >= 0 ? foreground : null;
            Background = background >= 0 ? background : null;
        }

        /// <summary>
        /// 获取填充模式。
        /// </summary>
        internal string Pattern { get; }

        /// <summary>
        /// 获取前景索引色。
        /// </summary>
        internal short? Foreground { get; }

        /// <summary>
        /// 获取背景索引色。
        /// </summary>
        internal short? Background { get; }
    }

    /// <summary>
    /// 保存 XLSX 一侧边框的线型和索引色。
    /// </summary>
    private sealed class XlsxBorderSide
    {
        /// <summary>
        /// 初始化一侧边框状态。
        /// </summary>
        internal XlsxBorderSide(string style, short? color)
        {
            Style = string.IsNullOrWhiteSpace(style) ? null : style.Trim();
            Color = color >= 0 ? color : null;
        }

        /// <summary>
        /// 获取线型名称。
        /// </summary>
        internal string Style { get; }

        /// <summary>
        /// 获取索引色。
        /// </summary>
        internal short? Color { get; }
    }

    /// <summary>
    /// 保存 XLSX 标准四边框状态。
    /// </summary>
    private sealed class XlsxBorder
    {
        /// <summary>
        /// 初始化一个边框状态描述。
        /// </summary>
        internal XlsxBorder(XlsxBorderSide left, XlsxBorderSide right,
            XlsxBorderSide top, XlsxBorderSide bottom)
        {
            Left = left;
            Right = right;
            Top = top;
            Bottom = bottom;
        }

        /// <summary>
        /// 获取左边框状态。
        /// </summary>
        internal XlsxBorderSide Left { get; }

        /// <summary>
        /// 获取右边框状态。
        /// </summary>
        internal XlsxBorderSide Right { get; }

        /// <summary>
        /// 获取上边框状态。
        /// </summary>
        internal XlsxBorderSide Top { get; }

        /// <summary>
        /// 获取下边框状态。
        /// </summary>
        internal XlsxBorderSide Bottom { get; }
    }

    /// <summary>
    /// 保存 XLSX 单元格的最小保护状态。
    /// </summary>
    private sealed class XlsxProtection
    {
        /// <summary>
        /// 初始化一个保护状态描述。
        /// </summary>
        internal XlsxProtection(bool? isLocked, bool? isHidden)
        {
            IsLocked = isLocked;
            IsHidden = isHidden;
        }

        /// <summary>
        /// 获取是否锁定单元格。
        /// </summary>
        internal bool? IsLocked { get; }

        /// <summary>
        /// 获取是否隐藏单元格公式。
        /// </summary>
        internal bool? IsHidden { get; }
    }

    /// <summary>
    /// 保存 XLSX 单元格样式和数字格式信息。
    /// </summary>
    private sealed class XlsxStyleTable
    {
        private readonly IReadOnlyList<int> _formatIds;
        private readonly IReadOnlyList<int> _fontIds;
        private readonly IReadOnlyList<int> _fillIds;
        private readonly IReadOnlyList<XlsxFill> _fills;
        private readonly IReadOnlyList<int> _borderIds;
        private readonly IReadOnlyList<XlsxBorder> _borders;
        private readonly IReadOnlyList<XlsxProtection> _protections;
        private readonly IReadOnlyList<bool> _fontBold;
        /// <summary>
        /// 按字体索引保存斜体状态。
        /// </summary>
        private readonly IReadOnlyList<bool> _fontItalic;
        /// <summary>
        /// 按字体索引保存删除线状态。
        /// </summary>
        private readonly IReadOnlyList<bool> _fontStrikeout;
        /// <summary>
        /// 按字体索引保存单下划线状态。
        /// </summary>
        private readonly IReadOnlyList<bool> _fontUnderline;
        private readonly IReadOnlyList<string> _fontNames;
        private readonly IReadOnlyList<double> _fontHeights;
        /// <summary>
        /// 按字体索引保存字体颜色索引。
        /// </summary>
        private readonly IReadOnlyList<short?> _fontColorIndexes;
        /// <summary>
        /// 按单元格样式保存对齐状态。
        /// </summary>
        private readonly IReadOnlyList<XlsxAlignment> _alignments;
        private readonly IReadOnlyDictionary<int, string> _customFormats;
        private readonly DataFormatter _formatter = new(CultureInfo.InvariantCulture);

        internal XlsxStyleTable(IReadOnlyList<int> formatIds, IReadOnlyList<int> fontIds,
            IReadOnlyList<int> fillIds, IReadOnlyList<XlsxFill> fills,
            IReadOnlyList<int> borderIds, IReadOnlyList<XlsxBorder> borders,
            IReadOnlyList<bool> fontBold,
            IReadOnlyList<bool> fontItalic, IReadOnlyList<bool> fontStrikeout,
            IReadOnlyList<bool> fontUnderline,
            IReadOnlyList<string> fontNames, IReadOnlyList<double> fontHeights,
            IReadOnlyList<short?> fontColorIndexes,
            IReadOnlyList<XlsxAlignment> alignments,
            IReadOnlyList<XlsxProtection> protections,
            IReadOnlyDictionary<int, string> customFormats)
        {
            _formatIds = formatIds ?? Array.Empty<int>();
            _fontIds = fontIds ?? Array.Empty<int>();
            _fillIds = fillIds ?? Array.Empty<int>();
            _fills = fills ?? Array.Empty<XlsxFill>();
            _borderIds = borderIds ?? Array.Empty<int>();
            _borders = borders ?? Array.Empty<XlsxBorder>();
            _fontBold = fontBold ?? Array.Empty<bool>();
            _fontItalic = fontItalic ?? Array.Empty<bool>();
            _fontStrikeout = fontStrikeout ?? Array.Empty<bool>();
            _fontUnderline = fontUnderline ?? Array.Empty<bool>();
            _fontNames = fontNames ?? Array.Empty<string>();
            _fontHeights = fontHeights ?? Array.Empty<double>();
            _fontColorIndexes = fontColorIndexes ?? Array.Empty<short?>();
            _alignments = alignments ?? Array.Empty<XlsxAlignment>();
            _protections = protections ?? Array.Empty<XlsxProtection>();
            _customFormats = customFormats ?? new Dictionary<int, string>();
        }

        internal static XlsxStyleTable Empty { get; } = new(Array.Empty<int>(),
            Array.Empty<int>(), Array.Empty<int>(), Array.Empty<XlsxFill>(),
            Array.Empty<int>(), Array.Empty<XlsxBorder>(),
            Array.Empty<bool>(), Array.Empty<bool>(),
            Array.Empty<bool>(), Array.Empty<bool>(), Array.Empty<string>(),
            Array.Empty<double>(), Array.Empty<short?>(), Array.Empty<XlsxAlignment>(),
            Array.Empty<XlsxProtection>(),
            new Dictionary<int, string>());

        internal int GetFormatIndex(int styleIndex) =>
            styleIndex >= 0 && styleIndex < _formatIds.Count ? _formatIds[styleIndex] : 0;

        internal string GetFormatCode(int? styleIndex)
        {
            if (!styleIndex.HasValue)
                return null;
            var formatIndex = GetFormatIndex(styleIndex.Value);
            return _customFormats.TryGetValue(formatIndex, out var custom)
                ? custom : BuiltinFormats.GetBuiltinFormat(formatIndex);
        }

        internal bool IsBold(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= _fontIds.Count)
                return false;
            var fontIndex = _fontIds[styleIndex];
            return fontIndex >= 0 && fontIndex < _fontBold.Count && _fontBold[fontIndex];
        }

        internal bool TryGetFontStyle(int styleIndex,
            out ExcelDataReaderFailureFontStyle style)
        {
            style = null;
            if (styleIndex < 0 || styleIndex >= _fontIds.Count)
                return false;
            var fontIndex = _fontIds[styleIndex];
            if (fontIndex < 0 || fontIndex >= _fontBold.Count)
                return false;
            var fill = GetFill(styleIndex);
            var border = GetBorder(styleIndex);
            var protection = styleIndex < _protections.Count ? _protections[styleIndex] : null;
            style = new ExcelDataReaderFailureFontStyle(
                _fontBold[fontIndex],
                fontIndex < _fontItalic.Count && _fontItalic[fontIndex],
                fontIndex < _fontStrikeout.Count && _fontStrikeout[fontIndex],
                fontIndex < _fontUnderline.Count && _fontUnderline[fontIndex],
                fontIndex < _fontNames.Count ? _fontNames[fontIndex] : null,
                fontIndex < _fontHeights.Count ? _fontHeights[fontIndex] : 0,
                fontIndex < _fontColorIndexes.Count ? _fontColorIndexes[fontIndex] : null,
                styleIndex < _alignments.Count ? _alignments[styleIndex]?.Horizontal : null,
                styleIndex < _alignments.Count ? _alignments[styleIndex]?.Vertical : null,
                styleIndex < _alignments.Count && _alignments[styleIndex]?.WrapText == true,
                 styleIndex < _alignments.Count && _alignments[styleIndex]?.ShrinkToFit == true,
                 styleIndex < _alignments.Count ? _alignments[styleIndex]?.Rotation : null,
                 styleIndex < _alignments.Count ? _alignments[styleIndex]?.Indent ?? 0 : (short)0,
                 fill?.Pattern, fill?.Foreground, fill?.Background,
                 ToFailureBorderSide(border?.Left), ToFailureBorderSide(border?.Right),
                 ToFailureBorderSide(border?.Top), ToFailureBorderSide(border?.Bottom),
                 protection?.IsLocked, protection?.IsHidden);
            return true;
        }

        /// <summary>
        /// 获取单元格样式引用的填充状态。
        /// </summary>
        private XlsxFill GetFill(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= _fillIds.Count)
                return null;
            var fillIndex = _fillIds[styleIndex];
            return fillIndex >= 0 && fillIndex < _fills.Count ? _fills[fillIndex] : null;
        }

        /// <summary>
        /// 获取单元格样式引用的边框状态。
        /// </summary>
        private XlsxBorder GetBorder(int styleIndex)
        {
            if (styleIndex < 0 || styleIndex >= _borderIds.Count)
                return null;
            var borderIndex = _borderIds[styleIndex];
            return borderIndex >= 0 && borderIndex < _borders.Count
                ? _borders[borderIndex] : null;
        }

        /// <summary>
        /// 将 XLSX 边框一侧转换为失败输出状态。
        /// </summary>
        private static ExcelDataReaderFailureBorderSide ToFailureBorderSide(XlsxBorderSide side)
            => side == null || string.IsNullOrWhiteSpace(side.Style)
                ? null : new ExcelDataReaderFailureBorderSide(side.Style, side.Color);

        internal bool TryFormat(double value, int styleIndex, bool date1904,
            out string formatted, out bool isDate)
        {
            var formatIndex = GetFormatIndex(styleIndex);
            var formatCode = _customFormats.TryGetValue(formatIndex, out var custom)
                ? custom : BuiltinFormats.GetBuiltinFormat(formatIndex);
            if (string.IsNullOrWhiteSpace(formatCode))
            {
                formatted = null;
                isDate = false;
                return false;
            }
            formatted = _formatter.FormatRawCellContents(value, formatIndex, formatCode, date1904);
            isDate = IsDateFormat(formatIndex, formatCode);
            return true;
        }

        private static bool IsDateFormat(int formatIndex, string formatCode)
        {
            if ((formatIndex >= 14 && formatIndex <= 22)
                || (formatIndex >= 27 && formatIndex <= 36)
                || (formatIndex >= 45 && formatIndex <= 47))
                return true;
            if (string.IsNullOrWhiteSpace(formatCode))
                return false;

            var quoted = false;
            var escaped = false;
            for (var index = 0; index < formatCode.Length; index++)
            {
                var character = formatCode[index];
                if (escaped)
                {
                    escaped = false;
                    continue;
                }
                if (character == '\\')
                {
                    escaped = true;
                    continue;
                }
                if (character == '"')
                {
                    quoted = !quoted;
                    continue;
                }
                if (quoted)
                    continue;
                if (character == '[')
                {
                    var end = formatCode.IndexOf(']', index + 1);
                    if (end < 0)
                        break;
                    var bracket = formatCode.Substring(index + 1, end - index - 1);
                    if (bracket.IndexOfAny(new[] { 'y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 's', 'S' }) >= 0)
                        return true;
                    index = end;
                    continue;
                }
                if (character == 'y' || character == 'Y' || character == 'm' || character == 'M'
                    || character == 'd' || character == 'D' || character == 'h' || character == 'H'
                    || character == 's' || character == 'S')
                    return true;
            }
            return false;
        }
    }

    /// <summary>
    /// 保存工作簿关系信息。
    /// </summary>
    private sealed class WorkbookInfo
    {
        internal WorkbookInfo(IReadOnlyList<WorkbookSheetInfo> sheets, bool date1904,
            IReadOnlyList<WorkbookDefinedName> definedNames)
        {
            Sheets = Array.AsReadOnly((sheets ?? Array.Empty<WorkbookSheetInfo>()).ToArray());
            if (Sheets.Count == 0)
                throw new ArgumentException("工作簿至少需要一个工作表。", nameof(sheets));
            Date1904 = date1904;
            DefinedNames = Array.AsReadOnly((definedNames ?? Array.Empty<WorkbookDefinedName>()).ToArray());
        }

        internal IReadOnlyList<WorkbookSheetInfo> Sheets { get; }
        internal string SheetName => Sheets[0].Name;
        internal string SheetPath => Sheets[0].Path;
        internal bool Date1904 { get; }
        internal IReadOnlyList<WorkbookDefinedName> DefinedNames { get; }

        internal WorkbookSheetInfo FindSheet(string name) => Sheets.FirstOrDefault(sheet =>
            string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));

        internal int GetSheetIndex(string name)
        {
            for (var index = 0; index < Sheets.Count; index++)
            {
                if (string.Equals(Sheets[index].Name, name, StringComparison.OrdinalIgnoreCase))
                    return index;
            }
            return -1;
        }
    }

    /// <summary>
    /// 保存 Forward XLSX 工作簿中的命名范围定义。
    /// </summary>
    private sealed class WorkbookDefinedName
    {
        internal WorkbookDefinedName(string name, int localSheetId, string formula)
        {
            Name = name;
            LocalSheetId = localSheetId;
            Formula = formula;
        }

        internal string Name { get; }
        internal int LocalSheetId { get; }
        internal string Formula { get; }
    }

    /// <summary>
    /// 保存工作簿工作表的名称和部件路径。
    /// </summary>
    private sealed class WorkbookSheetInfo
    {
        internal WorkbookSheetInfo(string name, string path)
        {
            Name = name;
            Path = path;
        }

        internal string Name { get; }
        internal string Path { get; }
    }

    /// <summary>
    /// 保存前向解析的一行单元格。
    /// </summary>
    private sealed class ForwardRow
    {
        internal ForwardRow(int rowIndex, IReadOnlyDictionary<int, ExcelCellValue> cells,
            IReadOnlyDictionary<int, int> styleIndexes)
        {
            RowIndex = rowIndex;
            Cells = cells;
            StyleIndexes = styleIndexes;
        }

        internal int RowIndex { get; }
        internal IReadOnlyDictionary<int, ExcelCellValue> Cells { get; }
        internal IReadOnlyDictionary<int, int> StyleIndexes { get; }
    }

    /// <summary>
    /// 保存解析出的物理单元格。
    /// </summary>
    private readonly struct ForwardCell
    {
        internal ForwardCell(int columnIndex, ExcelCellValue value, int styleIndex)
        {
            ColumnIndex = columnIndex;
            Value = value;
            StyleIndex = styleIndex;
        }

        internal int ColumnIndex { get; }
        internal ExcelCellValue Value { get; }
        internal int StyleIndex { get; }
    }

    /// <summary>
    /// 保存固定单元格导入结果及错误截断状态。
    /// </summary>
    private readonly struct FixedCellResult
    {
        internal FixedCellResult(IReadOnlyList<ExcelImportError> errors, bool isTruncated)
        {
            Errors = errors ?? Array.Empty<ExcelImportError>();
            IsTruncated = isTruncated;
        }

        internal IReadOnlyList<ExcelImportError> Errors { get; }
        internal bool IsTruncated { get; }

        internal static FixedCellResult Empty => new(Array.Empty<ExcelImportError>(), false);
    }

    /// <summary>
    /// 保存前向物化所需的列计划。
    /// </summary>
    private sealed class ForwardContext<TItem> where TItem : class, new()
    {
        internal ForwardContext(ExcelImportExecutionOptions<TItem> options,
            NpoiImportRowMaterializer materializer, IReadOnlyList<string> sharedStrings,
            ExcelXlsxWorkbookValidation workbookValidation,
            ExcelEntityForwardRelationTracker relationTracker)
        {
            Options = options;
            Materializer = materializer;
            SharedStrings = sharedStrings;
            WorkbookValidation = workbookValidation;
            RelationTracker = relationTracker;
        }

        internal ExcelImportExecutionOptions<TItem> Options { get; }
        internal NpoiImportRowMaterializer Materializer { get; }
        internal IReadOnlyList<string> SharedStrings { get; }
        internal ExcelXlsxWorkbookValidation WorkbookValidation { get; }
        internal ExcelEntityForwardRelationTracker RelationTracker { get; }
        internal IReadOnlyDictionary<int, ExcelColumnPlan> Columns { get; set; }
    }

    /// <summary>
    /// 保存单行处理结果。
    /// </summary>
    private readonly struct RowResult
    {
        internal RowResult(int rowIndex, object item, IReadOnlyList<ExcelImportError> errors,
            bool hasBatch)
        {
            RowIndex = rowIndex;
            Item = item;
            Errors = errors ?? Array.Empty<ExcelImportError>();
            HasBatch = hasBatch;
        }

        internal int RowIndex { get; }
        internal object Item { get; }
        internal IReadOnlyList<ExcelImportError> Errors { get; }
        internal bool HasBatch { get; }
    }

    /// <summary>
    /// 保存批次交付所需的根对象和区域描述。
    /// </summary>
    private sealed class SyncBatchDispatcher<TEntity> where TEntity : class, new()
    {
        private readonly TEntity _entity;
        private readonly ExcelEntityListRegion<TEntity> _region;
        private readonly int _batchSize;
        private readonly Action<ExcelEntityImportBatch<TEntity>> _onBatch;
        private List<object> _items = new();
        private List<ExcelImportError> _errors = new();
        private int _firstRow;
        private int _batchNumber = 1;

        internal SyncBatchDispatcher(TEntity entity, ExcelEntityListRegion<TEntity> region, int batchSize,
            Action<ExcelEntityImportBatch<TEntity>> onBatch)
        {
            _entity = entity;
            _region = region;
            _batchSize = batchSize;
            _onBatch = onBatch;
        }

        internal void Add(int rowIndex, object item, IReadOnlyList<ExcelImportError> errors)
        {
            if (_firstRow == 0)
                _firstRow = rowIndex + 1;
            if (item != null)
                _items.Add(item);
            if (errors != null)
                _errors.AddRange(errors);
            if (_items.Count >= _batchSize || _errors.Count >= _batchSize)
                Flush();
        }

        internal void Flush()
        {
            if (_items.Count == 0 && _errors.Count == 0)
                return;
            _onBatch(CreateBatch());
        }

        private ExcelEntityImportBatch<TEntity> CreateBatch()
        {
            IReadOnlyList<object> items = _items.Count == 0 ? Array.Empty<object>() : _items;
            IReadOnlyList<ExcelImportError> errors = _errors.Count == 0
                ? Array.Empty<ExcelImportError>()
                : _errors;
            if (_items.Count > 0)
                _items = new List<object>();
            if (_errors.Count > 0)
                _errors = new List<ExcelImportError>();
            var firstRow = _firstRow;
            _firstRow = 0;
            return new ExcelEntityImportBatch<TEntity>(_entity, _region.SheetName,
                _region.PropertyName, _region.ItemType, 0, _batchNumber++, firstRow,
                items, errors, false);
        }
    }

    /// <summary>
    /// 保存异步批次交付所需的根对象和区域描述。
    /// </summary>
    private sealed class AsyncBatchDispatcher<TEntity> where TEntity : class, new()
    {
        private readonly TEntity _entity;
        private readonly ExcelEntityListRegion<TEntity> _region;
        private readonly int _batchSize;
        private readonly Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> _onBatch;
        private List<object> _items = new();
        private List<ExcelImportError> _errors = new();
        private int _firstRow;
        private int _batchNumber = 1;

        internal AsyncBatchDispatcher(TEntity entity, ExcelEntityListRegion<TEntity> region, int batchSize,
            Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch)
        {
            _entity = entity;
            _region = region;
            _batchSize = batchSize;
            _onBatch = onBatch;
        }

        internal async Task AddAsync(int rowIndex, object item, IReadOnlyList<ExcelImportError> errors,
            CancellationToken cancellationToken)
        {
            if (_firstRow == 0)
                _firstRow = rowIndex + 1;
            if (item != null)
                _items.Add(item);
            if (errors != null)
                _errors.AddRange(errors);
            if (_items.Count >= _batchSize || _errors.Count >= _batchSize)
                await FlushAsync(cancellationToken).ConfigureAwait(false);
        }

        internal async Task FlushAsync(CancellationToken cancellationToken)
        {
            if (_items.Count == 0 && _errors.Count == 0)
                return;
            await _onBatch(CreateBatch(), cancellationToken).ConfigureAwait(false);
        }

        private ExcelEntityImportBatch<TEntity> CreateBatch()
        {
            IReadOnlyList<object> items = _items.Count == 0 ? Array.Empty<object>() : _items;
            IReadOnlyList<ExcelImportError> errors = _errors.Count == 0
                ? Array.Empty<ExcelImportError>()
                : _errors;
            if (_items.Count > 0)
                _items = new List<object>();
            if (_errors.Count > 0)
                _errors = new List<ExcelImportError>();
            var firstRow = _firstRow;
            _firstRow = 0;
            return new ExcelEntityImportBatch<TEntity>(_entity, _region.SheetName,
                _region.PropertyName, _region.ItemType, 0, _batchNumber++, firstRow,
                items, errors, false);
        }
    }

    /// <summary>
    /// 前向解析的根对象占位类型。
    /// </summary>
    private sealed class ForwardRoot<TItem> where TItem : class, new()
    {
        public ICollection<TItem> Items { get; } = new List<TItem>();
    }
}
