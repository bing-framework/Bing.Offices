namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 单元范围
/// </summary>
public class Range : IRange
{
    #region 字段

    /// <summary>
    /// 单元行列表
    /// </summary>
    private readonly IList<IRow> _rows;

    /// <summary>
    /// 起始索引
    /// </summary>
    private readonly int _startIndex;

    #endregion

    #region 属性

    /// <summary>
    /// 单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    public IRow this[int rowIndex] => _rows[rowIndex];

    /// <summary>
    /// 最大列数
    /// </summary>
    public int MaxColumnCount => _rows.Count > 0 ? _rows[0].ColumnCount : 0;

    /// <summary>
    /// 行数
    /// </summary>
    public int RowCount => _rows.Count;

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化一个<see cref="Range"/>类型的实例
    /// </summary>
    /// <param name="startIndex">起始索引</param>
    public Range(int startIndex = 0)
    {
        _rows = new List<IRow>();
        _startIndex = startIndex;
    }

    #endregion

    #region GetRow(获取单元行)

    /// <summary>
    /// 获取单元行
    /// </summary>
    /// <param name="rowIndex">行索引。对应Excel表格行号</param>
    public IRow GetRow(int rowIndex)
    {
        var realIndex = rowIndex - _startIndex;
        if (realIndex < 0)
            return null;
        if (realIndex > _rows.Count - 1)
            return null;
        return _rows[realIndex];
    }

    #endregion

    #region GetRows(获取单元行列表)

    /// <summary>
    /// 获取单元行列表
    /// </summary>
    public IList<IRow> GetRows() => _rows;

    #endregion

    #region AddRow(添加单元行)

    /// <summary>
    /// 添加单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="cells">单元格列表</param>
    public void AddRow(int rowIndex, IEnumerable<ICell> cells)
    {
        if (cells == null)
            return;
        var row = CreateRow(rowIndex);
        foreach (var cell in cells)
            AddCell(row, cell, rowIndex);
    }

    /// <summary>
    /// 添加单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="physicalRowIndex">物理行索引</param>
    /// <param name="cells">单元格列表</param>
    public void AddRow(int rowIndex, int physicalRowIndex, IEnumerable<ICell> cells)
    {
        if (cells == null)
            return;
        var row = CreateRow(rowIndex);
        row.PhysicalRowIndex = physicalRowIndex;
        foreach (var cell in cells)
            AddCell(row, cell, rowIndex);
    }

    /// <summary>
    /// 创建单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    private IRow CreateRow(int rowIndex)
    {
        var row = GetRow(rowIndex);
        if (row != null)
            return row;
        row = new Row(rowIndex);
        _rows.Add(row);
        return row;
    }

    /// <summary>
    /// 添加单元格
    /// </summary>
    /// <param name="row">单元行</param>
    /// <param name="cell">单元格</param>
    /// <param name="rowIndex">行索引</param>
    private void AddCell(IRow row, ICell cell, int rowIndex)
    {
        row.Add(cell);
        if (cell.RowSpan <= 1)
            return;
        for (var i = 1; i < cell.RowSpan; i++)
            AddPlaceholderCell(cell, rowIndex + 1);
    }

    /// <summary>
    /// 为下方受RowSpan影响的单元行添加占位单元格
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="rowIndex">行索引</param>
    private void AddPlaceholderCell(ICell cell, int rowIndex)
    {
        var row = CreateRow(rowIndex);
        row.Add(new NullCell() {ColumnIndex = cell.ColumnIndex, ColumnSpan = cell.ColumnSpan});
    }

    #endregion

    /// <summary>
    /// 清空单元行
    /// </summary>
    public void Clear()
    {
        throw new NotImplementedException();
    }
}