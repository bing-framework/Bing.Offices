using Bing.Offices.Exceptions;
using Bing.Offices.Providers;

namespace Bing.Offices.Exports;

/// <summary>
/// 按报表选择工作簿导出能力。
/// </summary>
/// <remarks>
/// 完整工作簿模式由所选 Provider 决定内部实现，可能使用流式优化；前向流式模式只调用
/// <see cref="IExcelStreamingExporter"/>，不接受模板，也不会回退到完整工作簿导出器。
/// </remarks>
public sealed class ExcelWorkbookExportStrategy
{
    /// <summary>
    /// 完整工作簿导出器。
    /// </summary>
    private readonly IExcelExporter _completeExporter;

    /// <summary>
    /// 前向流式导出器。
    /// </summary>
    private readonly IExcelStreamingExporter _streamingExporter;

    /// <summary>
    /// 初始化一个 <see cref="ExcelWorkbookExportStrategy"/> 类型的实例。
    /// </summary>
    /// <param name="completeExporter">完整工作簿导出器；可以为空。</param>
    /// <param name="streamingExporter">前向流式导出器；可以为空。</param>
    public ExcelWorkbookExportStrategy(IExcelExporter completeExporter = null,
        IExcelStreamingExporter streamingExporter = null)
    {
        if (completeExporter == null && streamingExporter == null)
            throw new ArgumentException("至少需要配置一种工作簿导出能力。");
        _completeExporter = completeExporter;
        _streamingExporter = streamingExporter;
    }

    /// <summary>
    /// 按指定模式写入工作簿流。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="destination">调用方拥有的可写目标流。</param>
    /// <param name="mode">本次报表的执行模式。</param>
    /// <param name="streamingOptions">前向流式选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public void Export(ExcelWorkbookExportRequest request, Stream destination, ExcelWorkbookExportMode mode,
        ExcelStreamingExportOptions streamingOptions = null, CancellationToken cancellationToken = default)
    {
        Validate(request, mode, streamingOptions, cancellationToken);
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));

        if (mode == ExcelWorkbookExportMode.CompleteWorkbook)
            _completeExporter.Export(request, destination, cancellationToken);
        else
            _streamingExporter.ExportBatches(request, destination, streamingOptions, cancellationToken);
    }

    /// <summary>
    /// 按指定模式异步写入工作簿流。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="destination">调用方拥有的可写目标流。</param>
    /// <param name="mode">本次报表的执行模式。</param>
    /// <param name="streamingOptions">前向流式选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination, ExcelWorkbookExportMode mode,
        ExcelStreamingExportOptions streamingOptions = null, CancellationToken cancellationToken = default)
    {
        Validate(request, mode, streamingOptions, cancellationToken);
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));

        return mode == ExcelWorkbookExportMode.CompleteWorkbook
            ? _completeExporter.ExportAsync(request, destination, cancellationToken)
            : _streamingExporter.ExportBatchesAsync(request, destination, streamingOptions, cancellationToken);
    }

    /// <summary>
    /// 按指定模式写入工作簿文件。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="mode">本次报表的执行模式。</param>
    /// <param name="streamingOptions">前向流式选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public void ExportToFile(ExcelWorkbookExportRequest request, string path, ExcelWorkbookExportMode mode,
        ExcelStreamingExportOptions streamingOptions = null, CancellationToken cancellationToken = default)
    {
        Validate(request, mode, streamingOptions, cancellationToken);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));

        if (mode == ExcelWorkbookExportMode.CompleteWorkbook)
            _completeExporter.ExportToFile(request, path, cancellationToken);
        else
            _streamingExporter.ExportBatchesToFile(request, path, streamingOptions, cancellationToken);
    }

    /// <summary>
    /// 按指定模式异步写入工作簿文件。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="mode">本次报表的执行模式。</param>
    /// <param name="streamingOptions">前向流式选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    public Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path, ExcelWorkbookExportMode mode,
        ExcelStreamingExportOptions streamingOptions = null, CancellationToken cancellationToken = default)
    {
        Validate(request, mode, streamingOptions, cancellationToken);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("目标文件路径不能为空。", nameof(path));

        return mode == ExcelWorkbookExportMode.CompleteWorkbook
            ? _completeExporter.ExportToFileAsync(request, path, cancellationToken)
            : _streamingExporter.ExportBatchesToFileAsync(request, path, streamingOptions, cancellationToken);
    }

    /// <summary>
    /// 校验模式与导出器能力。
    /// </summary>
    /// <param name="request">工作簿导出请求。</param>
    /// <param name="mode">本次报表的执行模式。</param>
    /// <param name="streamingOptions">前向流式选项。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private void Validate(ExcelWorkbookExportRequest request, ExcelWorkbookExportMode mode,
        ExcelStreamingExportOptions streamingOptions, CancellationToken cancellationToken)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (mode != ExcelWorkbookExportMode.CompleteWorkbook && mode != ExcelWorkbookExportMode.ForwardStreaming)
            throw new ArgumentOutOfRangeException(nameof(mode));
        if (mode == ExcelWorkbookExportMode.CompleteWorkbook && streamingOptions != null)
            throw new ArgumentException("完整工作簿模式不接受前向流式选项。", nameof(streamingOptions));
        if (mode == ExcelWorkbookExportMode.ForwardStreaming)
            streamingOptions?.Validate();
        cancellationToken.ThrowIfCancellationRequested();

        var provider = mode == ExcelWorkbookExportMode.CompleteWorkbook
            ? (object)_completeExporter : _streamingExporter;
        if (provider == null)
            throw Unsupported("未配置所选模式的导出器。", "Core");

        var descriptor = provider as IExcelProviderCapabilityDescriptor;
        var name = (provider as IExcelProviderCapabilities)?.ProviderName ?? provider.GetType().Name;
        if (descriptor != null)
        {
            if (mode == ExcelWorkbookExportMode.CompleteWorkbook && !descriptor.SupportsCompleteWorkbookExport)
                throw Unsupported("Provider 未声明完整工作簿导出能力。", name);
            if (!descriptor.WriteFormats.Contains(request.Format))
                throw Unsupported($"Provider 不支持写入 {request.Format} 格式。", name);
        }
        else if (provider is IExcelProviderCapabilities capabilities
            && !capabilities.Supports(ExcelProviderCapabilities.Workbook))
            throw Unsupported("Provider 未声明工作簿导出能力。", name);

        if (mode == ExcelWorkbookExportMode.ForwardStreaming)
        {
            if (request.Template != null)
                throw Unsupported("前向流式模式不支持模板工作簿。", name);
            if (provider is IExcelProviderFeatureDescriptor features
                && !features.Supports(ExcelProviderFeatures.StreamingWorkbookCreation))
                throw Unsupported("Provider 未声明前向流式创建能力。", name);
        }
        else if (request.Template != null)
        {
            if (provider is IExcelProviderFeatureDescriptor features
                && !features.Supports(ExcelProviderFeatures.TemplateEditing))
                throw Unsupported("Provider 未声明模板编辑能力。", name);
            if (provider is IExcelProviderCapabilities capabilities
                && !capabilities.Supports(ExcelProviderCapabilities.Template))
                throw Unsupported("Provider 未声明模板编辑能力。", name);
        }
    }

    /// <summary>
    /// 创建带有预检上下文的不支持功能异常。
    /// </summary>
    /// <param name="message">具体限制。</param>
    /// <param name="provider">Provider 名称。</param>
    /// <returns>不支持功能异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message, string provider) =>
        new BingOfficesUnsupportedFeatureException(message, provider: provider,
            operation: BingOfficesOperation.Export, stage: BingOfficesStage.Preflight);
}
