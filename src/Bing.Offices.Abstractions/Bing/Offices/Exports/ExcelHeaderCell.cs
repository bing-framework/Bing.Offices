namespace Bing.Offices.Exports;

/// <summary>
/// Excel 自定义表头单元格布局。
/// </summary>
public sealed class ExcelHeaderCell
{
    /// <summary>
    /// 初始化一个<see cref="ExcelHeaderCell"/>类型的实例。
    /// </summary>
    /// <param name="columnIndex">从零开始的列索引。</param>
    /// <param name="value">单元格值。</param>
    /// <param name="columnSpan">列跨度。</param>
    /// <param name="rowSpan">行跨度。</param>
    /// <param name="comment">单元格批注。</param>
    public ExcelHeaderCell(int columnIndex, object value, int columnSpan = 1, int rowSpan = 1,
        ExcelComment comment = null)
    {
        if (columnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));
        if (columnSpan < 1)
            throw new ArgumentOutOfRangeException(nameof(columnSpan));
        if (rowSpan < 1)
            throw new ArgumentOutOfRangeException(nameof(rowSpan));
        ColumnIndex = columnIndex;
        Value = value;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
        Comment = comment;
    }

    /// <summary>
    /// 获取从零开始的列索引。
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// 获取单元格值。
    /// </summary>
    public object Value { get; }

    /// <summary>
    /// 获取列跨度。
    /// </summary>
    public int ColumnSpan { get; }

    /// <summary>
    /// 获取行跨度。
    /// </summary>
    public int RowSpan { get; }

    /// <summary>
    /// 获取表头批注。
    /// </summary>
    public ExcelComment Comment { get; }
}
