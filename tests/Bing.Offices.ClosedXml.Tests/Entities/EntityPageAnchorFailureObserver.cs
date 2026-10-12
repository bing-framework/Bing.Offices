using Bing.Offices.Exceptions;

namespace Bing.Offices.ClosedXml.Tests.Entities;

/// <summary>
/// 记录逐页锚点资源测试中的公共异常观察结果。
/// </summary>
internal sealed class EntityPageAnchorFailureObserver : IBingOfficesExceptionObserver
{
    /// <summary>
    /// 获取观察次数。
    /// </summary>
    public int Count { get; private set; }

    /// <summary>
    /// 获取最近一次观察到的异常。
    /// </summary>
    public BingOfficesException LastException { get; private set; }

    /// <inheritdoc />
    public void Observe(BingOfficesException exception)
    {
        Count++;
        LastException = exception;
    }
}
