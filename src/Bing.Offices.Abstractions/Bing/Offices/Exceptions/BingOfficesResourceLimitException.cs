namespace Bing.Offices.Exceptions;

/// <summary>
/// 输入或输出资源预算超出异常。
/// </summary>
public sealed class BingOfficesResourceLimitException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesResourceLimitException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述资源限制失败的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="operation">触发资源限制的操作类型。</param>
    /// <param name="stage">触发资源限制的处理阶段。</param>
    public BingOfficesResourceLimitException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Preflight)
        : base(BingOfficesErrorCode.ResourceLimitExceeded, operation, provider, stage,
            message, innerException)
    {
    }
}
