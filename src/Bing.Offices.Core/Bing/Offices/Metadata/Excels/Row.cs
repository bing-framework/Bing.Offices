﻿using Bing.Offices.Metadata.Excels.Internal;

namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 单元行
/// </summary>
public class Row : IRow
{
    #region 字段

    /// <summary>
    /// 索引管理器
    /// </summary>
    private readonly IndexManager _indexManager;

    #endregion

    #region 属性

    /// <summary>
    /// 行索引
    /// </summary>
    public int RowIndex { get; set; }

    /// <summary>
    /// 物理行索引
    /// </summary>
    public int PhysicalRowIndex { get; set; }

    /// <summary>
    /// 是否有效
    /// </summary>
    public bool Valid => string.IsNullOrWhiteSpace(ErrorMsg);

    /// <summary>
    /// 错误消息
    /// </summary>
    public string ErrorMsg { get; set; }

    /// <summary>
    /// 单元格列表
    /// </summary>
    public IList<ICell> Cells { get; set; }

    /// <summary>
    /// 单元格
    /// </summary>
    /// <param name="columnIndex">列索引</param>
    public ICell this[int columnIndex] => Cells[columnIndex];

    /// <summary>
    /// 列数
    /// </summary>
    public int ColumnCount => Cells.Sum(t => t.ColumnSpan);

    #endregion

    #region 构造函数

    /// <summary>
    /// 初始化一个<see cref="Row"/>类型的实例
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    public Row(int rowIndex)
    {
        Cells = new List<ICell>();
        RowIndex = rowIndex;
        _indexManager = new IndexManager();
    }

    #endregion

    #region Add(添加单元格)

    /// <summary>
    /// 添加单元格
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="columnSpan">列跨度</param>
    /// <param name="rowSpan">行跨度</param>
    public void Add(object value, int columnSpan = 1, int rowSpan = 1) => Add(new Cell(value, columnSpan, rowSpan));

    /// <summary>
    /// 添加单元格
    /// </summary>
    /// <param name="cell">单元格</param>
    public void Add(ICell cell)
    {
        if (cell == null)
            return;
        cell.Row = this;
        SetColumnIndex(cell);
        Cells.Add(cell);
    }

    /// <summary>
    /// 设置列索引
    /// </summary>
    /// <param name="cell">单元格</param>
    private void SetColumnIndex(ICell cell)
    {
        if (cell.ColumnIndex > 0)
        {
            _indexManager.AddIndex(cell.ColumnIndex, cell.ColumnSpan);
            return;
        }
        cell.ColumnIndex = _indexManager.GetIndex(cell.ColumnSpan);
    }

    #endregion

}