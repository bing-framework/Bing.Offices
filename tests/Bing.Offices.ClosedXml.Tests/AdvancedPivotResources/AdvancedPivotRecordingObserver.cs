using System.Collections.Generic;
using Bing.Offices.Exceptions;

namespace Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;

/// <summary>
/// 记录高级透视资源测试观察到的公共异常。
/// </summary>
internal sealed class AdvancedPivotRecordingObserver : IBingOfficesExceptionObserver
{
    /// <summary>
    /// 获取已观察的 Bing.Offices 异常。
    /// </summary>
    public List<BingOfficesException> Exceptions { get; } = new();

    /// <inheritdoc />
    public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
}
