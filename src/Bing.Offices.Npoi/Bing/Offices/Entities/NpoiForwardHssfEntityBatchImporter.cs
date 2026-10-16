using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 执行 NPOI HSSF/XLS 的受限实体批次前向导入。
/// </summary>
/// <remarks>
/// HSSF 支持无模板、同一或不同工作表的多个非重叠列表、直接 A1 固定字段、声明合并及单格命名锚点的前向读取；事件解析结果通过
/// 磁盘暂存的 XLSX 投影交给既有 Forward 读取器，以保持映射、转换、验证、批次和摘要行为一致。投影只保留列表、固定字段或声明合并所在工作表。
/// </remarks>
internal sealed class NpoiForwardHssfEntityBatchImporter
{
    private readonly NpoiForwardXlsxEntityBatchImporter _forwardImporter;

    /// <summary>
    /// 初始化一个 <see cref="NpoiForwardHssfEntityBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="forwardImporter">负责读取转换后 XLSX 投影的前向导入器。</param>
    internal NpoiForwardHssfEntityBatchImporter(NpoiForwardXlsxEntityBatchImporter forwardImporter)
    {
        _forwardImporter = forwardImporter
            ?? throw new ArgumentNullException(nameof(forwardImporter));
    }

    /// <summary>
    /// 预检 HSSF/XLS 或 XLSX 前向实体模板。
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

        using var input = NpoiForwardInputDetector.Open(template.Template,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken);
        if (!input.IsHssf)
        {
            _forwardImporter.ValidateTemplate(
                new ExcelEntityTemplateOptions(input.Stream, leaveOpen: true), layout, options,
                cancellationToken);
            return;
        }

        ValidateRequestOptions(layout, options);
        using var converted = Convert(input.Stream, layout, options, cancellationToken);
        _forwardImporter.ValidateTemplate(
            new ExcelEntityTemplateOptions(converted, leaveOpen: true), layout, options,
            cancellationToken);
    }

    /// <summary>
    /// 异步预检 HSSF/XLS 或 XLSX 前向实体模板。
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

        await using var input = await NpoiForwardInputDetector.OpenAsync(template.Template,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken)
            .ConfigureAwait(false);
        if (!input.IsHssf)
        {
            await _forwardImporter.ValidateTemplateAsync(
                new ExcelEntityTemplateOptions(input.Stream, leaveOpen: true), layout, options,
                cancellationToken).ConfigureAwait(false);
            return;
        }

        ValidateRequestOptions(layout, options);
        using var converted = Convert(input.Stream, layout, options, cancellationToken);
        await _forwardImporter.ValidateTemplateAsync(
            new ExcelEntityTemplateOptions(converted, leaveOpen: true), layout, options,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 同步执行 HSSF 实体批次前向导入。
    /// </summary>
    internal ExcelEntityBatchImportSummary<TEntity> Import<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken,
        ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        ValidateRequest(source, layout, options, onBatch);
        var failureCapture = NpoiForwardXlsxEntityBatchImporter.CreateFailureCapture(options);
        using var converted = Convert(source, layout, options, cancellationToken, failureCapture,
            temporaryStorage);
        var result = _forwardImporter.Import(converted, layout, options, onBatch,
            cancellationToken, writeFailureWorkbook: false,
            failureCaptureOverride: failureCapture,
            captureFixedCellFailures: true, temporaryStorage: temporaryStorage);
        _forwardImporter.WriteHssfFailureWorkbook(source, layout, options, result,
            cancellationToken, failureCapture);
        return result;
    }

    /// <summary>
    /// 异步执行 HSSF 实体批次前向导入。
    /// </summary>
    internal async Task<ExcelEntityBatchImportSummary<TEntity>> ImportAsync<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken, ExcelTemporaryStorageBudget temporaryStorage = null)
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
        ValidateRequestOptions(layout, options);
        var failureCapture = NpoiForwardXlsxEntityBatchImporter.CreateFailureCapture(options);
        await using var converted = Convert(source, layout, options, cancellationToken, failureCapture,
            temporaryStorage);
        var result = await _forwardImporter.ImportAsync(converted, layout, options, onBatch,
            cancellationToken, writeFailureWorkbook: false,
            failureCaptureOverride: failureCapture,
            captureFixedCellFailures: true, temporaryStorage: temporaryStorage)
            .ConfigureAwait(false);
        await _forwardImporter.WriteHssfFailureWorkbookAsync(source, layout, options, result,
            cancellationToken, failureCapture).ConfigureAwait(false);
        return result;
    }

    /// <summary>
    /// 校验 HSSF 前向路径的请求范围。
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
        ValidateRequestOptions(layout, options);
    }

    /// <summary>
    /// 校验 HSSF 前向路径的布局边界。
    /// </summary>
    private static void ValidateRequestOptions<TEntity>(ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options) where TEntity : class, new()
    {
        NpoiForwardXlsxEntityBatchImporter.ValidateBatchLayout(layout, options);
        NpoiForwardXlsxEntityBatchImporter.ValidateResultMode(options);
    }

    /// <summary>
    /// 将 HSSF 输入转换为前向读取器使用的临时 XLSX。
    /// </summary>
    private static Stream Convert<TEntity>(Stream source, ExcelEntityLayout<TEntity> layout,
        ExcelEntityBatchImportOptions options, CancellationToken cancellationToken,
        Bing.Offices.ExcelDataReader.ExcelDataReaderFailureWorkbookCapture failureCapture = null,
        ExcelTemporaryStorageBudget temporaryStorage = null)
        where TEntity : class, new()
    {
        var sheetNames = layout.ListRegions.Select(region => region.SheetName)
            .Concat(layout.Cells.Select(cell => cell.SheetName))
            .Concat(layout.Merges.Select(merge => merge.SheetName))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        return NpoiHssfForwardXlsxConverter.Convert(source, sheetNames,
            options.ImportOptions.ResourceLimits, cancellationToken, failureCapture,
            options.ImportOptions.ValidationMode == ExcelImportValidationMode.WorkbookRules
            || options.ImportOptions.ValidationMode == ExcelImportValidationMode.ConfiguredAndWorkbook,
            temporaryStorage);
    }
}
