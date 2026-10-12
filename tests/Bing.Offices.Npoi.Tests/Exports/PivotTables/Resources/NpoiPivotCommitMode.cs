namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;

/// <summary>
/// 指定 NPOI 透视文件提交测试注入的故障阶段。
/// </summary>
internal enum NpoiPivotCommitMode
{
    /// <summary>
    /// 完成真实默认原子文件提交。
    /// </summary>
    Success,

    /// <summary>
    /// 在真实暂存流首次写入后抛出 IO 错误。
    /// </summary>
    FailDuringWrite,

    /// <summary>
    /// 在真实暂存流首次写入后取消。
    /// </summary>
    CancelDuringWrite,

    /// <summary>
    /// 完整序列化后在目标替换前抛出提交异常。
    /// </summary>
    FailAfterSerialization,

    /// <summary>
    /// 完整序列化后在目标替换前取消。
    /// </summary>
    CancelAfterSerialization
}
