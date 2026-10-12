using System.Collections.Generic;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Npoi.Tests.Exports.PivotTables.Resources;

/// <summary>
/// 记录 NPOI 透视资源测试中被观察的公共异常。
/// </summary>
internal sealed class NpoiPivotRecordingObserver : IBingOfficesExceptionObserver
{
    /// <summary>
    /// 获取按观察顺序记录的公共异常。
    /// </summary>
    public List<BingOfficesException> Exceptions { get; } = new();

    /// <inheritdoc />
    public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
}
