using System.Collections.Generic;

namespace Bing.Offices.Metadata.Excels
{
    /// <summary>
    /// 单元行
    /// </summary>
    public interface IRow
    {
        /// <summary>
        /// 单元格数量。
        /// <para>
        /// 获取已定义的单元格数（实际行中并没有单元格数）。
        /// 也就是说，如果仅列0、4、5具有值，则将为3
        /// </para>
        /// </summary>
        int CellsCount { get; }

        /// <summary>
        /// 首列列号。从1开始
        /// </summary>
        int FirstCellNum { get; }

        /// <summary>
        /// 尾列列号。从1开始
        /// </summary>
        int LastCellNum { get; }

        /// <summary>
        /// 底层对象值
        /// </summary>
        object UnderlyingValue { get; }

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引</param>
        ICell GetCell(int cellIndex);

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="cellIndex">
        /// 单元格索引
        /// <para>
        /// 最大值：(255 for *.xls, 1048576 for *.xlsx)
        /// </para>
        /// </param>
        ICell CreateCell(int cellIndex);

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
}
