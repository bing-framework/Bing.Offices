namespace Bing.Offices.ClosedXml.Internals;

/// <summary>
/// 维护一次 ClosedXML 导入中关系父键索引的剩余预算。
/// </summary>
internal sealed class ClosedXmlRelationKeyBudget
{
    /// <summary>
    /// 剩余可建立的父键索引数量；为空表示不限制。
    /// </summary>
    private int? _remaining;

    /// <summary>
    /// 初始化一个 <see cref="ClosedXmlRelationKeyBudget" /> 类型的实例。
    /// </summary>
    /// <param name="maximum">允许建立的最大父键索引数量。</param>
    internal ClosedXmlRelationKeyBudget(int? maximum)
    {
        _remaining = maximum;
    }

    /// <summary>
    /// 尝试为一个新的父键索引预留预算。
    /// </summary>
    /// <returns>预留成功或未配置上限时返回 <see langword="true" />。</returns>
    internal bool TryReserve()
    {
        if (!_remaining.HasValue)
            return true;
        if (_remaining.Value <= 0)
            return false;
        _remaining--;
        return true;
    }
}
