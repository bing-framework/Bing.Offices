namespace Bing.Offices.Exceptions;

/// <summary>
/// 原子文件提交异常。
/// </summary>
public sealed class BingOfficesFileCommitException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesFileCommitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述文件提交失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="stage">发生提交错误的处理阶段。</param>
    public BingOfficesFileCommitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesStage stage = BingOfficesStage.Commit)
        : base(BingOfficesErrorCode.FileCommitFailed, BingOfficesOperation.FileCommit,
            provider, stage, message, innerException)
    {
    }
}
