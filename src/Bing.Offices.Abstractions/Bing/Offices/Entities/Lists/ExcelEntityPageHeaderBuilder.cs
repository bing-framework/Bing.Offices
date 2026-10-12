using Bing.Offices.Styles;
using Bing.Offices.Entities.Layout;

namespace Bing.Offices.Entities;

/// <summary>
/// 配置实体列表的重复页眉。
/// </summary>
/// <typeparam name="TItem">列表项类型。</typeparam>
public sealed class ExcelEntityPageHeaderBuilder<TItem> where TItem : class, new()
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelEntityPageHeaderBuilder{TItem}" /> 类型的实例。
    /// </summary>
    public ExcelEntityPageHeaderBuilder()
    {
    }

    /// <summary>保存页眉单元格配置。</summary>
    private readonly List<ExcelEntityPageHeaderCell> _cells = new();

    /// <summary>保存相对合并区域。</summary>
    private readonly List<ExcelEntityCellRange> _merges = new();

    /// <summary>
    /// 添加固定值页眉单元格。
    /// </summary>
    /// <param name="address">相对于源区域左上角的 A1 地址。</param>
    /// <param name="value">单元格值。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityPageHeaderBuilder<TItem> Cell(string address, object value,
        ExcelCellStyle style = null, string numberFormat = null) =>
        Cell<object>(address, _ => value, style, numberFormat);

    /// <summary>
    /// 添加按页计算的页眉单元格。
    /// </summary>
    /// <typeparam name="TValue">单元格值类型。</typeparam>
    /// <param name="address">相对于源区域左上角的 A1 地址。</param>
    /// <param name="valueFactory">根据当前页数据获取单元格值的委托。</param>
    /// <param name="style">可选的单元格样式。</param>
    /// <param name="numberFormat">可选的数字格式。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityPageHeaderBuilder<TItem> Cell<TValue>(string address,
        Func<ExcelEntityPageHeaderContext<TItem>, TValue> valueFactory,
        ExcelCellStyle style = null, string numberFormat = null)
    {
        if (valueFactory == null)
            throw new ArgumentNullException(nameof(valueFactory));
        var reference = ExcelEntityCellReference.Parse(address);
        if (_cells.Any(cell => cell.Reference.Equals(reference)))
            throw new ArgumentException("重复页眉包含重复单元格。", nameof(address));
        _cells.Add(new ExcelEntityPageHeaderCell(reference,
            (page, count, items) => valueFactory(new ExcelEntityPageHeaderContext<TItem>(page,
                count, Array.AsReadOnly(items.Cast<TItem>().ToArray()))), style, numberFormat));
        return this;
    }

    /// <summary>
    /// 添加相对页眉合并区域。
    /// </summary>
    /// <param name="range">相对于源区域左上角的 A1 区域。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityPageHeaderBuilder<TItem> Merge(string range)
    {
        var parsed = ExcelEntityCellRange.Parse(range);
        if (_merges.Any(existing => ExcelEntityLayoutValidation.Overlaps(existing, parsed)))
            throw new ArgumentException("重复页眉合并区域相互重叠。", nameof(range));
        _merges.Add(parsed);
        return this;
    }

    /// <summary>构建重复页眉描述。</summary>
    internal ExcelEntityPageHeaderDefinition Build(string sourceRange, string markerAddress,
        string markerText)
    {
        var source = ExcelEntityCellRange.Parse(sourceRange);
        var marker = ExcelEntityCellReference.Parse(markerAddress);
        if (!source.Contains(marker) || string.IsNullOrWhiteSpace(markerText))
            throw new ArgumentException("重复页眉 marker 必须位于源区域内且文本非空。");
        var width = source.Last.Column - source.First.Column;
        var height = source.Last.Row - source.First.Row;
        if (_cells.Any(cell => cell.Reference.Row > height || cell.Reference.Column > width
            || source.First.Row + cell.Reference.Row == marker.Row
            && source.First.Column + cell.Reference.Column == marker.Column)
            || _merges.Any(merge => merge.Last.Row > height || merge.Last.Column > width))
            throw new ArgumentException("重复页眉单元格或合并区域超出源区域，或占用 marker。");
        var relativeMarker = new ExcelEntityCellReference(marker.Row - source.First.Row,
            marker.Column - source.First.Column);
        if (_merges.Any(merge => merge.Contains(relativeMarker)
            && !merge.First.Equals(relativeMarker)))
            throw new ArgumentException("重复页眉 marker 必须位于合并区域左上角。");
        return new ExcelEntityPageHeaderDefinition(source, marker, markerText.Trim(), _cells, _merges);
    }
}
