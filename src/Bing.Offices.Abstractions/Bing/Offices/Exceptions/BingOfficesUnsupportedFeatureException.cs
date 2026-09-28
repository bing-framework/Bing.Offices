namespace Bing.Offices.Exceptions;

/// <summary>
/// 当前提供程序不支持请求功能异常。
/// </summary>
public sealed class BingOfficesUnsupportedFeatureException : BingOfficesException
{
    /// <summary>
    /// 初始化一个 <see cref="BingOfficesUnsupportedFeatureException" /> 类型的实例。
    /// </summary>
    /// <param name="message">描述不支持功能的消息。</param>
    /// <param name="innerException">导致当前异常的内部异常。</param>
    /// <param name="provider">报告异常的提供程序名称。</param>
    /// <param name="operation">触发不支持功能错误的操作类型。</param>
    /// <param name="stage">发生不支持功能错误的处理阶段。</param>
    public BingOfficesUnsupportedFeatureException(string message, Exception innerException = null,
        string provider = "Core", BingOfficesOperation operation = BingOfficesOperation.Import,
        BingOfficesStage stage = BingOfficesStage.Read)
        : base(BingOfficesErrorCode.UnsupportedFeature, operation, provider, stage,
            message, innerException)
    {
    }
}
