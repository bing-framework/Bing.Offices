using System.Collections.Generic;

namespace Bing.Offices.Metadata.Excels;

/// <summary>
/// 单元行
/// </summary>
public interface IRow
{
    /// <summary>
    /// 行索引
    /// </summary>
    int RowIndex { get; set; }

    /// <summary>
    /// 物理行索引
    /// </summary>
    int PhysicalRowIndex { get; set; }

    /// <summary>
    /// 是否有效
    /// </summary>
    bool Valid { get; }

    /// <summary>
    /// 错误消息
    /// </summary>
    string ErrorMsg { get; set; }

    /// <summary>
    /// 单元格列表
    /// </summary>
    IList<ICell> Cells { get; set; }

    /// <summary>
    /// 单元格
    /// </summary>
    /// <param name="columnIndex">列索引</param>
    ICell this[int columnIndex] { get; }

    /// <summary>
    /// 列数
    /// </summary>
    int ColumnCount { get;}

    /// <summary>
    /// 添加单元格
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="columnSpan">列跨度</param>
    /// <param name="rowSpan">行跨度</param>
    void Add(object value, int columnSpan = 1, int rowSpan = 1);

    /// <summary>
    /// 添加单元格
    /// </summary>
    /// <param name="cell">单元格</param>
    void Add(ICell cell);
}