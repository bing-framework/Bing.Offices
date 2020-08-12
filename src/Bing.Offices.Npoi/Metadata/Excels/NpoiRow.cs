using System;
using System.Collections.Generic;
using Bing.Offices.Metadata.Excels;
using NModel = NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Metadata.Excels
{
    /// <summary>
    /// Npoi单元行
    /// </summary>
    internal class NpoiRow : IRow
    {
        /// <summary>
        /// 单元行
        /// </summary>
        private readonly NModel.IRow _row;

        /// <summary>
        /// 单元格数量。
        /// <para>
        /// 获取已定义的单元格数（实际行中并没有单元格数）。
        /// 也就是说，如果仅列0、4、5具有值，则将为3
        /// </para>
        /// </summary>
        public int CellsCount => _row.PhysicalNumberOfCells;

        /// <summary>
        /// 首列列号。从1开始
        /// </summary>
        public int FirstCellNum => _row.FirstCellNum + 1;

        /// <summary>
        /// 尾列列号。从1开始
        /// </summary>
        public int LastCellNum => _row.LastCellNum;

        /// <summary>
        /// 底层对象值
        /// </summary>
        public object UnderlyingValue => _row;

        /// <summary>
        /// 初始化一个<see cref="NpoiRow"/>类型的实例
        /// </summary>
        /// <param name="row">单元行</param>
        public NpoiRow(NModel.IRow row) => _row = row ?? throw new ArgumentNullException(nameof(row));

        /// <summary>
        /// 获取单元格
        /// </summary>
        /// <param name="cellIndex">单元格索引</param>
        public ICell GetCell(int cellIndex)
        {
            var nCell = _row.GetCell(cellIndex);
            return null == nCell ? null : new NpoiCell(nCell);
        }

        /// <summary>
        /// 创建单元格
        /// </summary>
        /// <param name="cellIndex">
        /// 单元格索引
        /// <para>
        /// 最大值：(255 for *.xls, 1048576 for *.xlsx)
        /// </para>
        /// </param>
        public ICell CreateCell(int cellIndex) => new NpoiCell(_row.CreateCell(cellIndex));

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
        public bool Valid { get; }

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
        public ICell this[int columnIndex] => throw new NotImplementedException();

        /// <summary>
        /// 列数
        /// </summary>
        public int ColumnCount { get; }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="columnSpan">列跨度</param>
        /// <param name="rowSpan">行跨度</param>
        public void Add(object value, int columnSpan = 1, int rowSpan = 1)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加单元格
        /// </summary>
        /// <param name="cell">单元格</param>
        public void Add(ICell cell)
        {
            throw new NotImplementedException();
        }
    }
}
