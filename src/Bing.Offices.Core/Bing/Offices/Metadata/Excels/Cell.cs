using Bing.Extensions;

namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 单元格
/// </summary>
public class Cell : ICell
{
    #region 字段

    /// <summary>
    /// 列跨度
    /// </summary>
    private int _columnSpan;

    /// <summary>
    /// 行跨度
    /// </summary>
    private int _rowSpan;

    #endregion

    #region 属性

    /// <summary>
    /// 值
    /// </summary>
    public object Value { get; set; }

    /// <summary>
    /// 行
    /// </summary>
    public IRow Row { get; set; }

    /// <summary>
    /// 列跨度
    /// </summary>
    public int ColumnSpan
    {
        get => _columnSpan;
        set
        {
            if (value < 1)
                value = 1;
            _columnSpan = value;
        }
    }

    /// <summary>
    /// 行跨度
    /// </summary>
    public int RowSpan
    {
        get => _rowSpan;
        set
        {
            if (value < 1)
                value = 1;
            _rowSpan = value;
        }
    }

    /// <summary>
    /// 列索引
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// 行索引
    /// </summary>
    public int RowIndex
    {
        get
        {
            Row.CheckNull(nameof(Row));
            return Row.RowIndex;
        }
    }

    /// <summary>
    /// 物理行索引
    /// </summary>
    public int PhysicalRowIndex {
        get
        {
            Row.CheckNull(nameof(Row));
            return Row.PhysicalRowIndex;
        }
    }

    /// <summary>
    /// 结束列索引
    /// </summary>
    public int EndColumnIndex => ColumnIndex + ColumnSpan - 1;

    /// <summary>
    /// 结束行索引
    /// </summary>
    public int EndRowIndex => RowIndex + RowSpan - 1;

    /// <summary>
    /// 是否需要合并单元格。true:是,false:否
    /// </summary>
    public bool NeedMerge => ColumnSpan > 1 || RowSpan > 1;

    /// <summary>
    /// 名称
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 属性名称
    /// </summary>
    public string PropertyName { get; set; }

    /// <summary>
    /// 是否动态单元格
    /// </summary>
    public bool IsDynamic { get; set; }

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化一个<see cref="Cell"/>类型的实例
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="columnSpan">列跨度</param>
    /// <param name="rowSpan">行跨度</param>
    public Cell(object value, int columnSpan = 1, int rowSpan = 1)
    {
        Value = value;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
    }

    /// <summary>
    /// 初始化一个<see cref="Cell"/>类型的实例
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="columnIndex">列索引</param>
    /// <param name="columnSpan">列跨度</param>
    /// <param name="rowSpan">行跨度</param>
    public Cell(object value, int columnIndex, int columnSpan = 1, int rowSpan = 1)
    {
        Value = value;
        ColumnIndex = columnIndex;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
    }
    #endregion

    #region IsNull(是否为空单元格)

    /// <summary>
    /// 是否为空单元格
    /// </summary>
    public virtual bool IsNull() => false;

    #endregion

}