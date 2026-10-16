using System.IO.Compression;
using Bing.Offices.Entities;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Exceptions;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.ClosedXml.Imports;
using ClosedXML.Excel;

namespace Bing.Offices.ClosedXml.Entities;

/// <summary>
/// 执行 ClosedXML Provider 的受限 XLSX 实体批次前向导入。
/// </summary>
/// <remarks>
/// 该路径只读取带表头、固定列列表、显式动态列组和直接 A1 固定单元格；复杂布局明确返回配置错误，避免退回 DOM。
/// 可定位且位于起始位置的输入流会直接读取；其他输入按资源上限暂存到自动清理的临时文件。
/// 完整实体结果只在 Complete 模式保留；SummaryOnly 的最终实体为空聚合对象，批次回调仍接收读取期间实体。
/// SummaryOnly 与 ErrorRowsOnly 组合直接捕获表头和错误行；其他失败模式仍按原有 DOM 路径生成失败工作簿。
/// </remarks>
internal sealed class ClosedXmlForwardXlsxEntityBatchImporter
{
    /// <summary>
    /// 负责输出失败工作簿的 ClosedXML 导入器。
    /// </summary>
    private readonly ClosedXmlExcelImporter _importer;

    /// <summary>
    /// ClosedXML 实体布局执行器。
    /// </summary>
    private readonly ClosedXmlEntityLayoutExecutor _executor;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlForwardXlsxEntityBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="importer">负责实体执行和失败工作簿输出的 ClosedXML 导入器。</param>
    internal ClosedXmlForwardXlsxEntityBatchImporter(ClosedXmlExcelImporter importer)
    {
        _importer = importer ?? throw new ArgumentNullException(nameof(importer));
        _executor = importer.EntityExecutor;
    }

    /// <summary>
    /// 同步执行 ClosedXML XLSX 前向实体批次导入。
    /// </summary>
    internal ExcelEntityBatchImportSummary<TEntity> Import<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        ValidateRequest(source, layout, options, onBatch);
        var temporaryStorage = new ExcelTemporaryStorageBudget(options.MaxTemporaryStorageBytes, "ClosedXML");
        using var input = OpenSource(source, options.ImportOptions.ResourceLimits, cancellationToken,
            temporaryStorage);
        var failureCapture = CreateFailureCapture(options);
        var result = Read(input.Stream, layout, options, onBatch, cancellationToken, failureCapture,
            temporaryStorage);
        WriteFailureWorkbook(input.Stream, layout, options, result, cancellationToken, failureCapture);
        return result;
    }

    /// <summary>
    /// 异步执行 ClosedXML XLSX 前向实体批次导入。
    /// </summary>
    internal async Task<ExcelEntityBatchImportSummary<TEntity>> ImportAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        ValidateRequest(source, layout, options, onBatch);
        var temporaryStorage = new ExcelTemporaryStorageBudget(options.MaxTemporaryStorageBytes, "ClosedXML");
        await using var input = await OpenSourceAsync(source,
            options.ImportOptions.ResourceLimits, cancellationToken, temporaryStorage).ConfigureAwait(false);
        var failureCapture = CreateFailureCapture(options);
        var result = await ReadAsync(input.Stream, layout, options, onBatch, cancellationToken,
            failureCapture, temporaryStorage).ConfigureAwait(false);
        await WriteFailureWorkbookAsync(input.Stream, layout, options, result, cancellationToken,
            failureCapture)
            .ConfigureAwait(false);
        return result;
    }

    /// <summary>
    /// 为前向导入结果生成错误行失败工作簿。
    /// </summary>
    private void WriteFailureWorkbook<TEntity>(Stream buffered, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null) where TEntity : class, new()
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
            using var workbook = new XLWorkbook(buffered);
            _importer.WriteEntityFailureWorkbookForBatch(workbook, layout, failureOptions,
                result.Result.Errors, cancellationToken);
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
    private async Task WriteFailureWorkbookAsync<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        ExcelEntityBatchImportSummary<TEntity> result,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null) where TEntity : class, new()
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
            using var workbook = new XLWorkbook(buffered);
            await _importer.WriteEntityFailureWorkbookForBatchAsync(workbook, layout,
                failureOptions, result.Result.Errors, cancellationToken).ConfigureAwait(false);
        }
        catch (BingOfficesException exception)
        {
            _importer.ObserveException(exception);
            throw;
        }
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
        ValidateLayout(layout, options.ImportOptions, allowRelations: true);
        if (options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly)
            ExcelEntityForwardRelationTracker.ValidateSummaryLayout(layout);
        using var input = OpenSource(template.Template,
            options.ImportOptions.ResourceLimits, cancellationToken);
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
        ValidateLayout(layout, options.ImportOptions, allowRelations: true);
        if (options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly)
            ExcelEntityForwardRelationTracker.ValidateSummaryLayout(layout);
        await using var input = await OpenSourceAsync(template.Template,
            options.ImportOptions.ResourceLimits, cancellationToken).ConfigureAwait(false);
        ValidateTemplateStream(input.Stream, layout, options.ImportOptions.ResourceLimits,
            cancellationToken);
    }

    /// <summary>
    /// 校验前向模板的 XLSX 结构和实体布局工作表。
    /// </summary>
    private void ValidateTemplateStream<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelResourceLimits limits,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "ClosedXML", requireZip: true,
            cancellationToken: cancellationToken);
        buffered.Position = 0;
        ClosedXmlTemplatePreflight.Validate(buffered, BingOfficesOperation.Import,
            cancellationToken, limits);
        buffered.Position = 0;
        using var workbook = new XLWorkbook(buffered);
        _executor.ValidateTemplate(workbook, layout, cancellationToken);
    }

    /// <summary>
    /// 从已缓冲的 XLSX 输入读取同步批次。
    /// </summary>
    private ExcelEntityBatchImportSummary<TEntity> Read<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        var limits = options.ImportOptions.ResourceLimits;
        limits?.Validate();
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "ClosedXML", requireZip: true,
            cancellationToken: cancellationToken);
        var workbookValidation = UsesWorkbookValidation(options.ImportOptions.ValidationMode)
            ? ExcelXlsxWorkbookValidation.Load(buffered, cancellationToken)
            : null;
        buffered.Position = 0;
        using var archive = new ZipArchive(buffered, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ExcelForwardXlsxReader.ReadWorkbook(archive, cancellationToken);
        failureCapture?.CaptureWorkbookMetadata(
            ExcelXlsxWorkbookValidation.LoadFailureMetadata(archive, cancellationToken));
        ValidateForwardMerges(archive, workbook, layout, cancellationToken);
        using var sharedStrings = ExcelForwardXlsxReader.ReadSharedStrings(archive,
            cancellationToken, options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
                || options.MaxTemporaryStorageBytes.HasValue, temporaryStorage);
        var styles = ExcelForwardXlsxReader.ReadStyles(archive, cancellationToken);
        var entity = new TEntity();
        var errorBudget = new ClosedXmlForwardErrorBudget(
            options.ImportOptions.ResourceLimits.MaxErrors);
        var fixedCells = _executor.ReadForwardFixedCells(workbook,
            sheetName => ReadSheetRows(archive, workbook, sharedStrings, styles, sheetName,
                cancellationToken), layout.Cells, options, entity, cancellationToken,
            workbookValidation, errorBudget, failureCapture);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var sheets = new List<ExcelSheetImportResult>(layout.ListRegions.Count);
        var errors = fixedCells.Errors.ToList();
        var errorsTruncated = fixedCells.IsTruncated;
        var rowBudget = new ClosedXmlForwardRowBudget(options.ImportOptions.ResourceLimits.MaxRows);
        var uniqueBudget = new ClosedXmlForwardUniqueBudget(
            options.ImportOptions.ResourceLimits.MaxTrackedUniqueValues);
        var relationKeyBudget = new ClosedXmlRelationKeyBudget(
            options.ImportOptions.ResourceLimits.MaxTrackedRelationKeys);
        var relationTracker = options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && layout.Relations.Count > 0
            ? ExcelEntityForwardRelationTracker.Create(layout, relationKeyBudget.TryReserve,
                "ClosedXML")
            : null;
        var sourceRows = options.ResultMode == ExcelEntityBatchResultMode.Complete
            && layout.Relations.Count > 0
            ? new Dictionary<object, ExcelSheetImportResult>(
                ClosedXmlReferenceObjectComparer.Instance)
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
                : region.WithResolvedStart(ExcelForwardXlsxReader.ResolveSingleCellAnchor(
                    workbook, region.SheetName, region.AnchorName));
            ValidateSheet(workbook, resolvedRegion.SheetName);
            var sheet = workbook.FindSheet(resolvedRegion.SheetName);
            var sheetEntry = sheet == null ? null : archive.GetEntry(sheet.Path);
            if (sheetEntry == null)
                throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "ClosedXML",
                    stage: BingOfficesStage.Read, sheetName: resolvedRegion.SheetName);
            if (failureCapture != null)
                failureCapture.CaptureSheetColumnMetadata(resolvedRegion.SheetName,
                    ExcelForwardXlsxReader.ReadColumnMetadata(sheetEntry, cancellationToken));
            using var sheetStream = sheetEntry.Open();
            var regionResult = _executor.ReadForwardRegion(workbook,
                ExcelForwardXlsxReader.ReadRows(sheetStream, sharedStrings, styles, workbook.Date1904,
                    cancellationToken), layout, resolvedRegion, options, batch =>
                {
                    onBatch(new ExcelEntityImportBatch<TEntity>(entity, batch.SheetName,
                        batch.PropertyName, batch.ItemType, index, globalBatchNumber++,
                        batch.FirstRowIndex, batch.Items, batch.Errors, batch.IsDiagnostic));
                    }, cancellationToken, entity, workbookValidation, failureCapture, rowBudget,
                    errorBudget, uniqueBudget, relationTracker);
            if (sourceRows != null)
                CaptureSourceRows(entity, resolvedRegion, regionResult, sourceRows);
            AppendRegionResult(regionResult, index, summaries, sheets, errors, ref errorsTruncated);
            if (relationTracker?.IsTruncated == true)
                break;
        }
        if (relationTracker != null)
            errorsTruncated |= relationTracker.IsTruncated;
        else
            errorsTruncated |= ClosedXmlRelationCoordinator.Bind(entity, layout.Relations, errors,
                options.ImportOptions.ResourceLimits.MaxErrors, relationKeyBudget, cancellationToken,
                sourceSheets: sourceRows);
        return CreateCombinedSummary(entity, summaries, sheets, errors, errorsTruncated, options);
    }

    /// <summary>
    /// 从已缓冲的 XLSX 输入读取异步批次。
    /// </summary>
    private async Task<ExcelEntityBatchImportSummary<TEntity>> ReadAsync<TEntity>(Stream buffered,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken,
        ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        var limits = options.ImportOptions.ResourceLimits;
        limits?.Validate();
        buffered.Position = 0;
        ExcelXlsxZipPreflight.Validate(buffered, limits, "ClosedXML", requireZip: true,
            cancellationToken: cancellationToken);
        var workbookValidation = UsesWorkbookValidation(options.ImportOptions.ValidationMode)
            ? ExcelXlsxWorkbookValidation.Load(buffered, cancellationToken)
            : null;
        buffered.Position = 0;
        using var archive = new ZipArchive(buffered, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ExcelForwardXlsxReader.ReadWorkbook(archive, cancellationToken);
        failureCapture?.CaptureWorkbookMetadata(
            ExcelXlsxWorkbookValidation.LoadFailureMetadata(archive, cancellationToken));
        ValidateForwardMerges(archive, workbook, layout, cancellationToken);
        using var sharedStrings = ExcelForwardXlsxReader.ReadSharedStrings(archive,
            cancellationToken, options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
                || options.MaxTemporaryStorageBytes.HasValue, temporaryStorage);
        var styles = ExcelForwardXlsxReader.ReadStyles(archive, cancellationToken);
        var entity = new TEntity();
        var errorBudget = new ClosedXmlForwardErrorBudget(
            options.ImportOptions.ResourceLimits.MaxErrors);
        var fixedCells = _executor.ReadForwardFixedCells(workbook,
            sheetName => ReadSheetRows(archive, workbook, sharedStrings, styles, sheetName,
                cancellationToken), layout.Cells, options, entity, cancellationToken,
            workbookValidation, errorBudget, failureCapture);
        var summaries = new List<ExcelEntityBatchRegionSummary>(layout.ListRegions.Count);
        var sheets = new List<ExcelSheetImportResult>(layout.ListRegions.Count);
        var errors = fixedCells.Errors.ToList();
        var errorsTruncated = fixedCells.IsTruncated;
        var rowBudget = new ClosedXmlForwardRowBudget(options.ImportOptions.ResourceLimits.MaxRows);
        var uniqueBudget = new ClosedXmlForwardUniqueBudget(
            options.ImportOptions.ResourceLimits.MaxTrackedUniqueValues);
        var relationKeyBudget = new ClosedXmlRelationKeyBudget(
            options.ImportOptions.ResourceLimits.MaxTrackedRelationKeys);
        var relationTracker = options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && layout.Relations.Count > 0
            ? ExcelEntityForwardRelationTracker.Create(layout, relationKeyBudget.TryReserve,
                "ClosedXML")
            : null;
        var sourceRows = options.ResultMode == ExcelEntityBatchResultMode.Complete
            && layout.Relations.Count > 0
            ? new Dictionary<object, ExcelSheetImportResult>(
                ClosedXmlReferenceObjectComparer.Instance)
            : null;
        var globalBatchNumber = 1;
        if (fixedCells.Errors.Count > 0)
        {
            await onBatch(new ExcelEntityImportBatch<TEntity>(entity, null, null,
                typeof(TEntity), -1, globalBatchNumber++, 0, Array.Empty<object>(),
                fixedCells.Errors, true), cancellationToken).ConfigureAwait(false);
        }
        for (var index = 0; index < layout.ListRegions.Count; index++)
        {
            var region = layout.ListRegions[index];
            var resolvedRegion = region.AnchorName == null
                ? region
                : region.WithResolvedStart(ExcelForwardXlsxReader.ResolveSingleCellAnchor(
                    workbook, region.SheetName, region.AnchorName));
            ValidateSheet(workbook, resolvedRegion.SheetName);
            var sheet = workbook.FindSheet(resolvedRegion.SheetName);
            var sheetEntry = sheet == null ? null : archive.GetEntry(sheet.Path);
            if (sheetEntry == null)
                throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "ClosedXML",
                    stage: BingOfficesStage.Read, sheetName: resolvedRegion.SheetName);
            if (failureCapture != null)
                failureCapture.CaptureSheetColumnMetadata(resolvedRegion.SheetName,
                    ExcelForwardXlsxReader.ReadColumnMetadata(sheetEntry, cancellationToken));
            await using var sheetStream = sheetEntry.Open();
            var regionResult = await _executor.ReadForwardRegionAsync(workbook,
                ExcelForwardXlsxReader.ReadRows(sheetStream, sharedStrings, styles, workbook.Date1904,
                    cancellationToken), layout, resolvedRegion, options, async (batch, _) =>
                {
                    await onBatch(new ExcelEntityImportBatch<TEntity>(entity, batch.SheetName,
                        batch.PropertyName, batch.ItemType, index, globalBatchNumber++,
                        batch.FirstRowIndex, batch.Items, batch.Errors, batch.IsDiagnostic),
                        cancellationToken).ConfigureAwait(false);
                }, cancellationToken, entity, workbookValidation, failureCapture, rowBudget,
                    errorBudget, uniqueBudget, relationTracker).ConfigureAwait(false);
            if (sourceRows != null)
                CaptureSourceRows(entity, resolvedRegion, regionResult, sourceRows);
            AppendRegionResult(regionResult, index, summaries, sheets, errors, ref errorsTruncated);
            if (relationTracker?.IsTruncated == true)
                break;
        }
        if (relationTracker != null)
            errorsTruncated |= relationTracker.IsTruncated;
        else
            errorsTruncated |= ClosedXmlRelationCoordinator.Bind(entity, layout.Relations, errors,
                options.ImportOptions.ResourceLimits.MaxErrors, relationKeyBudget, cancellationToken,
                sourceSheets: sourceRows);
        return CreateCombinedSummary(entity, summaries, sheets, errors, errorsTruncated, options);
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

    private static void CaptureSourceRows<TEntity>(TEntity entity,
        ExcelEntityListRegion<TEntity> region, ExcelEntityBatchImportSummary<TEntity> result,
        IDictionary<object, ExcelSheetImportResult> sourceRows) where TEntity : class, new()
    {
        var sheet = result.Result.Sheets.SingleOrDefault(item =>
            string.Equals(item.Name, region.SheetName, StringComparison.OrdinalIgnoreCase)
            && item.ItemType == region.ItemType);
        var collection = region.Getter(entity);
        if (sheet != null && collection != null)
            sourceRows[collection] = sheet;
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
    /// 校验 Forward 路径的布局边界。
    /// </summary>
    private static void ValidateRequest<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options, Delegate onBatch) where TEntity : class, new()
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
        ValidateLayout(layout, options.ImportOptions, allowRelations: true);
        if (options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly)
            ExcelEntityForwardRelationTracker.ValidateSummaryLayout(layout);
    }

    /// <summary>
    /// 校验前向路径支持的实体布局边界。
    /// </summary>
    private static void ValidateLayout<TEntity>(ExcelEntityLayout<TEntity> layout,
        ExcelEntityImportOptions importOptions, bool allowRelations = true)
        where TEntity : class, new()
    {
        if (!allowRelations && layout.Relations.Count > 0)
            Unsupported("摘要结果中的关系绑定");
        if (layout.LookupLists.Count > 0
            || layout.CascadeLists.Count > 0
            || layout.DataValidations.Count > 0 || layout.Protections.Count > 0)
            Unsupported("关系、合并以外的模板规则或工作表元数据");
        if (layout.ListRegions.Count == 0)
            Unsupported("空列表布局");
        foreach (var group in layout.ListRegions.GroupBy(region => region.SheetName,
                     StringComparer.OrdinalIgnoreCase))
        {
            if (group.Count() > 1 && group.Any(region => !region.End.HasValue))
                Unsupported("同一工作表中的多个列表区域必须声明结束地址");
        }
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
    /// 校验工作簿 Sheet 与布局目标一致。
    /// </summary>
    private static void ValidateSheet(ExcelForwardXlsxWorkbook workbook, string sheetName)
    {
        if (workbook.FindSheet(sheetName) == null)
            throw new BingOfficesConfigurationException(
                $"Forward 实体布局要求 Sheet {sheetName}，工作簿中不存在该工作表。",
                stage: BingOfficesStage.Plan);
    }

    /// <summary>
    /// 校验 Forward 布局声明的合并区域存在于源 worksheet。
    /// </summary>
    private static void ValidateForwardMerges<TEntity>(ZipArchive archive,
        ExcelForwardXlsxWorkbook workbook, ExcelEntityLayout<TEntity> layout,
        CancellationToken cancellationToken) where TEntity : class, new()
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
                    provider: "ClosedXML", stage: BingOfficesStage.Read, sheetName: sheet.Name);
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
    /// 创建低保留失败行捕获器。
    /// </summary>
    private static ExcelDataReaderFailureWorkbookCapture CreateFailureCapture(
        ExcelEntityBatchImportOptions options)
    {
        return options?.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && options.ImportOptions?.FailureOptions?.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly
            ? new ExcelDataReaderFailureWorkbookCapture(options.ImportOptions.FailureOptions, "ClosedXML",
                captureErrors: false)
            : null;
    }

    /// <summary>
    /// 判断是否使用只保留错误行的失败工作簿路径。
    /// </summary>
    private static bool IsLowRetentionFailureOutput(ExcelEntityBatchImportOptions options,
        ExcelDataReaderFailureWorkbookCapture failureCapture)
    {
        return options?.ResultMode == ExcelEntityBatchResultMode.SummaryOnly
            && options.ImportOptions?.FailureOptions?.Mode == ExcelImportFailureWorkbookMode.ErrorRowsOnly
            && failureCapture != null;
    }

    /// <summary>
    /// 抛出明确的前向路径能力边界错误。
    /// </summary>
    private static void Unsupported(string feature) =>
        throw new BingOfficesConfigurationException($"ClosedXML Forward 实体批次暂不支持{feature}。",
            stage: BingOfficesStage.Plan);

    /// <summary>
    /// 判断是否执行工作簿原生校验。
    /// </summary>
    private static bool UsesWorkbookValidation(ExcelImportValidationMode mode) =>
        mode == ExcelImportValidationMode.WorkbookRules
        || mode == ExcelImportValidationMode.ConfiguredAndWorkbook;

    /// <summary>
    /// 创建指定工作表的前向行事件序列。
    /// </summary>
    private static IEnumerable<ExcelForwardXlsxRow> ReadSheetRows(ZipArchive archive,
        ExcelForwardXlsxWorkbook workbook, IReadOnlyList<string> sharedStrings,
        ExcelForwardXlsxStyles styles, string sheetName, CancellationToken cancellationToken)
    {
        var sheet = workbook.FindSheet(sheetName);
        var entry = sheet == null ? null : archive.GetEntry(sheet.Path);
        if (entry == null)
            throw new BingOfficesImportException("XLSX 工作表部件不存在。", provider: "ClosedXML",
                stage: BingOfficesStage.Read, sheetName: sheetName);
        using var stream = entry.Open();
        foreach (var row in ExcelForwardXlsxReader.ReadRows(stream, sharedStrings, styles,
                     workbook.Date1904, cancellationToken))
            yield return row;
    }

    /// <summary>
    /// 同步复制输入并应用最大字节数限制。
    /// </summary>
    private static Stream BufferSource(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        limits?.Validate();
        var maximum = limits?.MaxInputBytes;
        if (source.CanSeek && maximum.HasValue && source.Length - source.Position > maximum.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                provider: "ClosedXML", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Preflight);
        var buffer = CreateTempBuffer(temporaryStorage);
        var bytes = new byte[81920];
        long total = 0;
        try
        {
            int read;
            while ((read = source.Read(bytes, 0, bytes.Length)) != 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                total += read;
                if (maximum.HasValue && total > maximum.Value)
                    throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                        provider: "ClosedXML", operation: BingOfficesOperation.Import,
                        stage: BingOfficesStage.Preflight);
                buffer.Write(bytes, 0, read);
            }
            buffer.Flush();
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
    /// 异步复制输入并应用最大字节数限制。
    /// </summary>
    private static async Task<Stream> BufferSourceAsync(Stream source,
        ExcelResourceLimits limits, CancellationToken cancellationToken,
        ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        limits?.Validate();
        var maximum = limits?.MaxInputBytes;
        if (source.CanSeek && maximum.HasValue && source.Length - source.Position > maximum.Value)
            throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                provider: "ClosedXML", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Preflight);
        var buffer = CreateTempBuffer(temporaryStorage);
        var bytes = new byte[81920];
        long total = 0;
        try
        {
            while (true)
            {
                var requested = bytes.Length;
                if (maximum.HasValue)
                    requested = (int)Math.Min(requested, maximum.Value - total + 1);
                var read = await source.ReadAsync(bytes.AsMemory(0, requested), cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0)
                    break;
                total += read;
                if (maximum.HasValue && total > maximum.Value)
                    throw new BingOfficesResourceLimitException("输入工作簿超过最大字节数。",
                        provider: "ClosedXML", operation: BingOfficesOperation.Import,
                        stage: BingOfficesStage.Preflight);
                await buffer.WriteAsync(bytes.AsMemory(0, read), cancellationToken)
                    .ConfigureAwait(false);
            }
            await buffer.FlushAsync(cancellationToken).ConfigureAwait(false);
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
    /// 打开可直接读取或需要暂存的输入流。
    /// </summary>
    private static ForwardInput OpenSource(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        limits?.Validate();
        if (CanReadInPlace(source, limits?.MaxInputBytes))
            return new ForwardInput(source, ownsStream: false);
        return new ForwardInput(BufferSource(source, limits, cancellationToken, temporaryStorage), ownsStream: true);
    }

    /// <summary>
    /// 异步打开可直接读取或需要暂存的输入流。
    /// </summary>
    private static async Task<ForwardInput> OpenSourceAsync(Stream source,
        ExcelResourceLimits limits, CancellationToken cancellationToken,
        ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        limits?.Validate();
        if (CanReadInPlace(source, limits?.MaxInputBytes))
            return new ForwardInput(source, ownsStream: false);
        return new ForwardInput(await BufferSourceAsync(source, limits, cancellationToken, temporaryStorage)
            .ConfigureAwait(false), ownsStream: true);
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
                provider: "ClosedXML", operation: BingOfficesOperation.Import,
                stage: BingOfficesStage.Preflight);
        return true;
    }

    /// <summary>
    /// 创建由流所有权自动清理的 Forward 输入暂存文件。
    /// </summary>
    private static Stream CreateTempBuffer(ExcelTemporaryStorageBudget temporaryStorage)
    {
        var path = Path.Combine(Path.GetTempPath(),
            $"bing-offices-forward-{Guid.NewGuid():N}.xlsx");
        return (temporaryStorage ?? new ExcelTemporaryStorageBudget(null, "ClosedXML")).CreateFile(path,
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
}
