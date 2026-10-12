namespace Bing.Offices.ClosedXml.Entities;

/// <summary>
/// 跟踪一次 ClosedXML Forward 导入跨列表区域共享的错误预算。
/// </summary>
internal sealed class ClosedXmlForwardErrorBudget
{
    /// <summary>
    /// 允许保留的最大错误数。
    /// </summary>
    private readonly int? _maximum;

    /// <summary>
    /// 已预留的错误数量。
    /// </summary>
    private int _count;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlForwardErrorBudget" /> 类型的实例。
    /// </summary>
    /// <param name="maximum">允许保留的最大错误数；为空时不限制。</param>
    internal ClosedXmlForwardErrorBudget(int? maximum)
    {
        _maximum = maximum;
    }

    /// <summary>
    /// 获取允许保留的最大错误数。
    /// </summary>
    internal int? Maximum => _maximum;

    /// <summary>
    /// 获取是否已达到错误数量上限。
    /// </summary>
    internal bool IsTruncated { get; private set; }

    /// <summary>
    /// 尝试为一条错误预留预算。
    /// </summary>
    /// <returns>成功预留时为 true；达到上限时为 false。</returns>
    internal bool TryReserve()
    {
        if (_maximum.HasValue && _count >= _maximum.Value)
        {
            IsTruncated = true;
            return false;
        }

        _count++;
        if (_maximum.HasValue && _count >= _maximum.Value)
            IsTruncated = true;
        return true;
    }

    /// <summary>
    /// 将当前导入标记为已截断。
    /// </summary>
    internal void MarkTruncated() => IsTruncated = true;
}
