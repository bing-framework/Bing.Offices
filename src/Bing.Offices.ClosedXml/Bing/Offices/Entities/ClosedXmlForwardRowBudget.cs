namespace Bing.Offices.ClosedXml.Entities;

/// <summary>
/// 跟踪一次 ClosedXML Forward 导入跨列表区域共享的数据行预算。
/// </summary>
internal sealed class ClosedXmlForwardRowBudget
{
    /// <summary>
    /// 允许读取的最大数据行数。
    /// </summary>
    private readonly int? _maximum;

    /// <summary>
    /// 已消耗的数据行数。
    /// </summary>
    private int _count;

    /// <summary>
    /// 指示是否已经报告行数上限错误。
    /// </summary>
    private bool _limitReported;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlForwardRowBudget" /> 类型的实例。
    /// </summary>
    /// <param name="maximum">允许读取的最大数据行数；为空时不限制。</param>
    internal ClosedXmlForwardRowBudget(int? maximum)
    {
        _maximum = maximum;
    }

    /// <summary>
    /// 尝试消耗一行数据预算。
    /// </summary>
    /// <returns>成功消耗预算时为 true；达到上限时为 false。</returns>
    internal bool TryConsume()
    {
        if (_maximum.HasValue && _count >= _maximum.Value)
            return false;
        _count++;
        return true;
    }

    /// <summary>
    /// 确保行数上限错误只报告一次。
    /// </summary>
    /// <returns>本次调用首次标记错误时为 true。</returns>
    internal bool TryMarkLimitReported()
    {
        if (_limitReported)
            return false;
        _limitReported = true;
        return true;
    }
}
