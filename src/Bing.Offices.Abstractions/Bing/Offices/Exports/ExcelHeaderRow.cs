namespace Bing.Offices.Exports;

/// <summary>
/// Excel 自定义表头行布局。
/// </summary>
public sealed class ExcelHeaderRow
{
    /// <summary>
    /// 初始化一个<see cref="ExcelHeaderRow"/>类型的实例。
    /// </summary>
    /// <param name="rowIndex">从零开始的行索引。</param>
    /// <param name="cells">表头单元格集合。</param>
    public ExcelHeaderRow(int rowIndex, IReadOnlyList<ExcelHeaderCell> cells)
    {
        if (rowIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(rowIndex));
        RowIndex = rowIndex;
        Cells = cells ?? throw new ArgumentNullException(nameof(cells));
    }

    /// <summary>
    /// 获取从零开始的行索引。
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// 获取表头单元格集合。
    /// </summary>
    public IReadOnlyList<ExcelHeaderCell> Cells { get; }
}
