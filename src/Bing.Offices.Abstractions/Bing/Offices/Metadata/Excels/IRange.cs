﻿namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 单元范围
/// </summary>
public interface IRange
{
    /// <summary>
    /// 单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    IRow this[int rowIndex] { get; }

    /// <summary>
    /// 最大列数
    /// </summary>
    int MaxColumnCount { get; }

    /// <summary>
    /// 行数
    /// </summary>
    int RowCount { get; }

    /// <summary>
    /// 获取单元行
    /// </summary>
    /// <param name="rowIndex">行索引。对应Excel表格行号</param>
    IRow GetRow(int rowIndex);

    /// <summary>
    /// 获取单元行列表
    /// </summary>
    IList<IRow> GetRows();

    /// <summary>
    /// 添加单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="cells">单元格列表</param>
    void AddRow(int rowIndex, IEnumerable<ICell> cells);

    /// <summary>
    /// 添加单元行
    /// </summary>
    /// <param name="rowIndex">行索引</param>
    /// <param name="physicalRowIndex">物理行索引</param>
    /// <param name="cells">单元格列表</param>
    void AddRow(int rowIndex, int physicalRowIndex, IEnumerable<ICell> cells);

    /// <summary>
    /// 清空单元行
    /// </summary>
    void Clear();
}
