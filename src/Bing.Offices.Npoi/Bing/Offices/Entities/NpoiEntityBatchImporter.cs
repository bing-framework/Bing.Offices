using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 将 NPOI 实体布局导入结果投影为列表区域批次。
/// </summary>
/// <remarks>
/// Complete 结果复用完整实体导入；SummaryOnly 结果复用受限 Forward 解析，避免保留完整实体集合。
/// Forward 预检仍负责拒绝不支持的复杂布局。
/// </remarks>
internal sealed class NpoiEntityBatchImporter : IExcelEntityBatchImporter
{
    /// <summary>
    /// 执行完整实体布局导入的 NPOI 导入器。
    /// </summary>
    private readonly Bing.Offices.Imports.NpoiExcelImporter _importer;

    /// <summary>
    /// 执行能力受限的 NPOI XLSX 前向实体批次导入。
    /// </summary>
    private readonly NpoiForwardXlsxEntityBatchImporter _forwardImporter;

    /// <summary>
    /// 执行能力受限的 NPOI HSSF 前向实体批次导入。
    /// </summary>
    private readonly NpoiForwardHssfEntityBatchImporter _forwardHssfImporter;

    /// <summary>
    /// 初始化一个 <see cref="NpoiEntityBatchImporter" /> 类型的实例。
    /// </summary>
    /// <param name="importer">负责执行完整实体布局导入的 NPOI 导入器。</param>
    internal NpoiEntityBatchImporter(Bing.Offices.Imports.NpoiExcelImporter importer)
    {
        _importer = importer ?? throw new ArgumentNullException(nameof(importer));
        _forwardImporter = new NpoiForwardXlsxEntityBatchImporter(_importer);
        _forwardHssfImporter = new NpoiForwardHssfEntityBatchImporter(_forwardImporter);
    }

    /// <inheritdoc />
    public ExcelEntityBatchImportSummary<TEntity> ImportEntityBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => ImportEntityBatches(source, layout, new ExcelEntityBatchImportOptions(), onBatch,
            cancellationToken);

    /// <inheritdoc />
    public ExcelEntityBatchImportSummary<TEntity> ImportEntityBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateResultMode(options);
        if (ShouldUseForward(options))
            return ImportForward(source, layout, ToForwardOptions(options), onBatch,
                cancellationToken);
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        var result = _importer.ImportEntity(source, layout, options.ImportOptions, cancellationToken);
        return ExcelEntityBatchProjection.Deliver(result, layout, options.BatchSize, onBatch,
            cancellationToken);
    }

    /// <inheritdoc />
    public async Task<ExcelEntityBatchImportSummary<TEntity>> ImportEntityBatchesAsync<TEntity>(
        Stream source, ExcelEntityLayout<TEntity> layout,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => await ImportEntityBatchesAsync(source, layout, new ExcelEntityBatchImportOptions(), onBatch,
            cancellationToken).ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<ExcelEntityBatchImportSummary<TEntity>> ImportEntityBatchesAsync<TEntity>(
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateResultMode(options);
        if (ShouldUseForward(options))
            return await ImportForwardAsync(source, layout, ToForwardOptions(options), onBatch,
                cancellationToken).ConfigureAwait(false);
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        var result = await _importer.ImportEntityAsync(source, layout, options.ImportOptions,
            cancellationToken).ConfigureAwait(false);
        return await ExcelEntityBatchProjection.DeliverAsync(result, layout, options.BatchSize,
            onBatch, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public ExcelEntityBatchImportSummary<TEntity> ImportForTemplateBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => ImportForTemplateBatches(source, layout, template, new ExcelEntityBatchImportOptions(),
            onBatch, cancellationToken);

    /// <inheritdoc />
    public ExcelEntityBatchImportSummary<TEntity> ImportForTemplateBatches<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityBatchImportOptions options, Action<ExcelEntityImportBatch<TEntity>> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateResultMode(options);
        if (ShouldUseForward(options))
        {
            var forwardOptions = ToForwardOptions(options);
            try
            {
                _forwardHssfImporter.ValidateTemplate(template, layout, forwardOptions,
                    cancellationToken);
                using var input = NpoiForwardInputDetector.Open(source,
                    forwardOptions.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken);
                if (input.IsHssf)
                    return _forwardHssfImporter.Import(input.Stream, layout, forwardOptions, onBatch,
                        cancellationToken);
                return _forwardImporter.Import(input.Stream, layout, forwardOptions, onBatch,
                    cancellationToken);
            }
            finally
            {
                DisposeForwardTemplate(template);
            }
        }
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        var result = _importer.ImportForTemplate(source, layout, template, options.ImportOptions,
            cancellationToken);
        return ExcelEntityBatchProjection.Deliver(result, layout, options.BatchSize, onBatch,
            cancellationToken);
    }

    /// <inheritdoc />
    public async Task<ExcelEntityBatchImportSummary<TEntity>> ImportForTemplateBatchesAsync<TEntity>(
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
        => await ImportForTemplateBatchesAsync(source, layout, template,
            new ExcelEntityBatchImportOptions(), onBatch, cancellationToken).ConfigureAwait(false);

    /// <inheritdoc />
    public async Task<ExcelEntityBatchImportSummary<TEntity>> ImportForTemplateBatchesAsync<TEntity>(
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
        ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken = default) where TEntity : class, new()
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));
        ValidateResultMode(options);
        if (ShouldUseForward(options))
        {
            var forwardOptions = ToForwardOptions(options);
            try
            {
                await _forwardHssfImporter.ValidateTemplateAsync(template, layout, forwardOptions,
                    cancellationToken).ConfigureAwait(false);
                await using var input = await NpoiForwardInputDetector.OpenAsync(source,
                    forwardOptions.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken)
                    .ConfigureAwait(false);
                if (input.IsHssf)
                    return await _forwardHssfImporter.ImportAsync(input.Stream, layout, forwardOptions,
                        onBatch, cancellationToken).ConfigureAwait(false);
                return await _forwardImporter.ImportAsync(input.Stream, layout, forwardOptions, onBatch,
                    cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                DisposeForwardTemplate(template);
            }
        }
        if (onBatch == null)
            throw new ArgumentNullException(nameof(onBatch));
        var result = await _importer.ImportForTemplateAsync(source, layout, template,
            options.ImportOptions, cancellationToken).ConfigureAwait(false);
        return await ExcelEntityBatchProjection.DeliverAsync(result, layout, options.BatchSize,
            onBatch, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 按模板选项释放前向预检使用的模板流。
    /// </summary>
    private static void DisposeForwardTemplate(ExcelEntityTemplateOptions template)
    {
        if (template != null && !template.LeaveOpen)
            template.Template.Dispose();
    }

    /// <summary>
    /// 按输入格式分派同步前向导入。
    /// </summary>
    private ExcelEntityBatchImportSummary<TEntity> ImportForward<TEntity>(Stream source,
        ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Action<ExcelEntityImportBatch<TEntity>> onBatch, CancellationToken cancellationToken)
        where TEntity : class, new()
    {
        using var input = NpoiForwardInputDetector.Open(source,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken);
        if (input.IsHssf)
            return _forwardHssfImporter.Import(input.Stream, layout, options, onBatch,
                cancellationToken);
        return _forwardImporter.Import(input.Stream, layout, options, onBatch, cancellationToken);
    }

    /// <summary>
    /// 按输入格式分派异步前向导入。
    /// </summary>
    private async Task<ExcelEntityBatchImportSummary<TEntity>> ImportForwardAsync<TEntity>(
        Stream source, ExcelEntityLayout<TEntity> layout, ExcelEntityBatchImportOptions options,
        Func<ExcelEntityImportBatch<TEntity>, CancellationToken, Task> onBatch,
        CancellationToken cancellationToken) where TEntity : class, new()
    {
        await using var input = await NpoiForwardInputDetector.OpenAsync(source,
            options.ImportOptions.ResourceLimits.MaxInputBytes, cancellationToken)
            .ConfigureAwait(false);
        if (input.IsHssf)
            return await _forwardHssfImporter.ImportAsync(input.Stream, layout, options, onBatch,
                cancellationToken).ConfigureAwait(false);
        return await _forwardImporter.ImportAsync(input.Stream, layout, options, onBatch,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// 判断批次是否应使用有界的前向解析路径。
    /// </summary>
    /// <remarks>
    /// 摘要结果不需要保留完整实体集合，因此 Buffered 请求也可复用 Forward 解析器；
    /// Forward 解析器仍会在布局预检阶段拒绝其不支持的复杂布局。
    /// </remarks>
    /// <param name="options">实体批次选项。</param>
    /// <returns>需要使用 Forward 解析时返回 <see langword="true" />。</returns>
    private static bool ShouldUseForward(ExcelEntityBatchImportOptions options) =>
        options.ParsingMode == ExcelEntityBatchParsingMode.Forward
        || options.ResultMode == ExcelEntityBatchResultMode.SummaryOnly;

    /// <summary>
    /// 创建用于有界前向解析的选项快照。
    /// </summary>
    /// <param name="options">原始实体批次选项。</param>
    /// <returns>解析模式为 Forward 的选项。</returns>
    private static ExcelEntityBatchImportOptions ToForwardOptions(
        ExcelEntityBatchImportOptions options) =>
        options.ParsingMode == ExcelEntityBatchParsingMode.Forward
            ? options
            : new ExcelEntityBatchImportOptions(options.BatchSize, options.ImportOptions,
                ExcelEntityBatchParsingMode.Forward, options.ResultMode);

    /// <summary>
    /// 校验批次结果保留模式及前向失败输出限制。
    /// </summary>
    private static void ValidateResultMode(ExcelEntityBatchImportOptions options)
    {
        if (ShouldUseForward(options)
            && options.ImportOptions.FailureOptions != null
            && options.ImportOptions.FailureOptions.Mode != ExcelImportFailureWorkbookMode.None
            && options.ImportOptions.FailureOptions.Mode != ExcelImportFailureWorkbookMode.ErrorRowsOnly
            && options.ImportOptions.FailureOptions.Mode != ExcelImportFailureWorkbookMode.AnnotatedOriginal)
            throw new BingOfficesUnsupportedFeatureException(
                "NPOI 前向实体批次仅支持 ErrorRowsOnly 或 AnnotatedOriginal 失败工作簿输出。", provider: "NPOI",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Plan);
    }
}
