namespace Bing.Offices.Npoi.Entities;

/// <summary>
/// 跟踪 NPOI Forward 多个列表区域共享的唯一值资源预算。
/// </summary>
internal sealed class NpoiForwardUniqueBudget
{
    /// <summary>
    /// 允许跟踪的唯一值总数；为空表示不限制。
    /// </summary>
    private readonly int? _maximum;

    /// <summary>
    /// 前序列表区域已提交的唯一值数量。
    /// </summary>
    private int _consumed;

    /// <summary>
    /// 初始化一个 <see cref="NpoiForwardUniqueBudget" /> 类型的实例。
    /// </summary>
    /// <param name="maximum">一次 Forward 导入允许跟踪的唯一值总数。</param>
    internal NpoiForwardUniqueBudget(int? maximum)
    {
        _maximum = maximum;
    }

    /// <summary>
    /// 获取当前列表区域可使用的唯一值数量。
    /// </summary>
    internal int? Remaining => _maximum.HasValue
        ? Math.Max(0, _maximum.Value - _consumed)
        : null;

    /// <summary>
    /// 累加当前列表区域实际提交的唯一值数量。
    /// </summary>
    /// <param name="count">当前区域已提交的唯一值数量。</param>
    internal void Consume(int count)
    {
        if (!_maximum.HasValue || count <= 0)
            return;
        _consumed = Math.Min(_maximum.Value, _consumed + count);
    }
}
