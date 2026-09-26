namespace Bing.Offices.Exports;

/// <summary>
/// 只创建新工作簿的前向流式导出能力。
/// </summary>
/// <remarks>
/// 该接口与完整工作簿导出能力分离。Provider 不应通过此接口打开或修改已有模板，
/// 也不应为了满足请求而回退到 DOM 工作簿。批次大小用于约束 Provider 的枚举和缓存边界，
/// 不代表调用方可以在已提交的输出中回滚之前的数据。
/// </remarks>
public interface IExcelStreamingExporter
{
    /// <summary>
    /// 按前向批次将新工作簿写入调用方拥有的目标流。
    /// </summary>
    /// <param name="request">不包含模板的 Workbook 导出请求。</param>
    /// <param name="destination">调用方拥有的可写目标流。</param>
    /// <param name="options">流式导出选项；为空使用默认批次大小。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void ExportBatches(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 按前向批次异步写入新工作簿。
    /// </summary>
    /// <remarks>外围文件和流写入必须使用真实异步 IO。</remarks>
    /// <param name="request">不包含模板的 Workbook 导出请求。</param>
    /// <param name="destination">调用方拥有的可写目标流。</param>
    /// <param name="options">流式导出选项；为空使用默认批次大小。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    Task ExportBatchesAsync(ExcelWorkbookExportRequest request, Stream destination,
        ExcelStreamingExportOptions options = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 按前向批次原子写入新工作簿文件。
    /// </summary>
    /// <param name="request">不包含模板的 Workbook 导出请求。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="options">流式导出选项；为空使用默认批次大小。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    void ExportBatchesToFile(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null,
        CancellationToken cancellationToken = default);

    /// <summary>
    /// 按前向批次异步原子写入新工作簿文件。
    /// </summary>
    /// <param name="request">不包含模板的 Workbook 导出请求。</param>
    /// <param name="path">目标文件路径。</param>
    /// <param name="options">流式导出选项；为空使用默认批次大小。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    Task ExportBatchesToFileAsync(ExcelWorkbookExportRequest request, string path,
        ExcelStreamingExportOptions options = null,
        CancellationToken cancellationToken = default);
}

/// <summary>
/// 前向流式导出选项。
/// </summary>
public sealed class ExcelStreamingExportOptions
{
    /// <summary>
    /// 获取或初始化每批最大行数。
    /// </summary>
    /// <remarks>默认值为 1,000，用于 Provider 的边界检查和批次提交。</remarks>
    public int BatchSize { get; init; } = 1000;

    /// <summary>
    /// 获取或初始化是否拒绝模板、回溯和需要完整 DOM 的请求。
    /// </summary>
    public bool RequireForwardOnly { get; init; } = true;

    /// <summary>
    /// 验证选项。
    /// </summary>
    public void Validate()
    {
        if (BatchSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(BatchSize));
    }
}
