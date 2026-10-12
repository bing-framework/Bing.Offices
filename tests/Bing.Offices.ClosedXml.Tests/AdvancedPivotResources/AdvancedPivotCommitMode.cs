namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 指定高级透视文件提交测试行为。
/// </summary>
internal enum AdvancedPivotCommitMode
{
    /// <summary>
    /// 正常完成真实默认提交。
    /// </summary>
    Success,

    /// <summary>
    /// 暂存流首次写入后取消。
    /// </summary>
    CancelDuringWrite,

    /// <summary>
    /// 序列化结束后取消默认提交。
    /// </summary>
    CancelAfterSerialization,

    /// <summary>
    /// 序列化结束后抛出提交异常。
    /// </summary>
    FailAfterSerialization
}
