using System;

namespace Bing.Offices.ClosedXml;

/// <summary>
/// ClosedXML Provider 的 DOM 资源准入配置。
/// </summary>
public sealed class ClosedXmlProviderOptions
{
    /// <summary>
    /// 获取或设置同一服务注册范围内允许同时创建的 Workbook 数量。
    /// </summary>
    public int MaxConcurrentWorkbooks { get; set; } = 1;

    /// <summary>
    /// 获取或设置等待 Workbook 准入的最大排队数量。
    /// </summary>
    public int MaxQueuedOperations { get; set; } = 64;

    /// <summary>
    /// 验证配置。
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">
    /// 并发 Workbook 数量小于等于零，或排队操作数量小于零。
    /// </exception>
    internal void Validate()
    {
        if (MaxConcurrentWorkbooks <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxConcurrentWorkbooks));
        if (MaxQueuedOperations < 0)
            throw new ArgumentOutOfRangeException(nameof(MaxQueuedOperations));
    }
}
