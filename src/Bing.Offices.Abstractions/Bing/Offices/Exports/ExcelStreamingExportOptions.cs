namespace Bing.Offices.Exports;

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
