using System.Collections.Generic;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Tests;

/// <summary>
/// 记录独立逐页锚点资源测试的异常观察。
/// </summary>
internal sealed class IndependentPageAnchorRecordingObserver : IBingOfficesExceptionObserver
{
    /// <summary>
    /// 初始化一个 <see cref="IndependentPageAnchorRecordingObserver" /> 类型的实例。
    /// </summary>
    public IndependentPageAnchorRecordingObserver()
    {
    }

    /// <summary>
    /// 获取观察到的异常。
    /// </summary>
    public List<BingOfficesException> Exceptions { get; } = new();

    /// <inheritdoc />
    public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
}
