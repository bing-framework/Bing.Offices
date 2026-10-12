namespace Bing.Offices.Entities;

/// <summary>
/// 获取实体明细页眉的当前页数据。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
public sealed class ExcelEntityPageHeaderContext<TItem>
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityPageHeaderContext{TItem}" /> 类型的实例。
    /// </summary>
    internal ExcelEntityPageHeaderContext(int pageNumber, int pageCount, IReadOnlyList<TItem> items)
    {
        PageNumber = pageNumber;
        PageCount = pageCount;
        Items = items;
    }

    /// <summary>获取从一开始的页码。</summary>
    public int PageNumber { get; }

    /// <summary>获取逻辑页总数。</summary>
    public int PageCount { get; }

    /// <summary>获取当前页的明细快照。</summary>
    public IReadOnlyList<TItem> Items { get; }
}
