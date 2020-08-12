using System;
using System.Collections.Generic;
using Bing.Offices.Metadata.Excels;
using NPOI.SS.Util;
using NModel = NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Metadata.Excels
{
    /// <summary>
    /// Npoi工作表
    /// </summary>
    internal class NpoiSheet : ISheet
    {
        /// <summary>
        /// 工作表
        /// </summary>
        private readonly NModel.ISheet _sheet;

        /// <summary>
        /// 首行行号。从1开始，如果没有行，则为0
        /// </summary>
        public int FirstRowNum => _sheet.FirstRowNum + 1;

        /// <summary>
        /// 尾行行号。从1开始，如果没有行，则为0
        /// </summary>
        public int LastRowNum => _sheet.LastRowNum + 1;

        /// <summary>
        /// 初始化一个<see cref="NpoiSheet"/>类型的实例
        /// </summary>
        /// <param name="sheet">工作表</param>
        public NpoiSheet(NModel.ISheet sheet) => _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        public IRow GetRow(int rowIndex)
        {
            var nRow = _sheet.GetRow(rowIndex);
            return null == nRow ? null : new NpoiRow(nRow);
        }

        /// <summary>
        /// 创建单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        public IRow CreateRow(int rowIndex) => new NpoiRow(_sheet.CreateRow(rowIndex));

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <param name="width">宽度</param>
        public void SetColumnWidth(int columnIndex, int width) => _sheet.SetColumnWidth(columnIndex, width);

        /// <summary>
        /// 设置自动列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        public void AutoSizeColumn(int columnIndex) => _sheet.AutoSizeColumn(columnIndex);

        /// <summary>
        /// 创建冻结窗格
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        /// <param name="leftMostCol">可见列数</param>
        /// <param name="topRow">可见行数</param>
        public void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow) => _sheet.CreateFreezePane(colSplit, rowSplit, leftMostCol, topRow);

        /// <summary>
        /// 设置自动筛选
        /// </summary>
        /// <param name="firstRowIndex">首行索引</param>
        /// <param name="lastRowIndex">尾行索引</param>
        /// <param name="firstColumnIndex">首列索引</param>
        /// <param name="lastColumnIndex">尾列索引</param>
        public void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex) => _sheet.SetAutoFilter(new CellRangeAddress(firstRowIndex, lastRowIndex, firstColumnIndex, lastColumnIndex));

        /// <summary>
        /// 移动行
        /// </summary>
        /// <param name="startRow">起始行索引</param>
        /// <param name="endRow">结束行索引</param>
        /// <param name="n">行数。正数: 下移, 负数: 上移</param>
        public void ShiftRows(int startRow, int endRow, int n) => _sheet.ShiftRows(startRow, endRow, n);

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="sourceIndex">源索引</param>
        /// <param name="targetIndex">目标索引</param>
        public IRow CopyRow(int sourceIndex, int targetIndex) => new NpoiRow(_sheet.CopyRow(sourceIndex, targetIndex));

        /// <summary>
        /// 移除行
        /// </summary>
        /// <param name="row">单元行</param>
        public void RemoveRow(IRow row) => _sheet.RemoveRow(row.UnderlyingValue as NModel.IRow);

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 最大列数
        /// </summary>
        public int MaxColumnCount { get; }

        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; }

        /// <summary>
        /// 表头行数
        /// </summary>
        public int HeadRowCount { get; }

        /// <summary>
        /// 正文行数
        /// </summary>
        public int BodyRowCount { get; }

        /// <summary>
        /// 页脚行数
        /// </summary>
        public int FootRowCount { get; }

        /// <summary>
        /// 获取表头
        /// </summary>
        public IList<IRow> GetHeader()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 获取正文
        /// </summary>
        public IList<IRow> GetBody()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 获取页脚
        /// </summary>
        public IList<IRow> GetFooter()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="titles">标题</param>
        public void AddHeadRow(params string[] titles)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="cells">单元格</param>
        public void AddHeadRow(params ICell[] cells)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cellValues">值</param>
        public void AddBodyRow(params object[] cellValues)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        public void AddBodyRow(IEnumerable<ICell> cells)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        /// <param name="physicalRowIndex">物理行索引</param>
        public void AddBodyRow(IEnumerable<ICell> cells, int physicalRowIndex)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cellValues">值</param>
        public void AddFootRow(params string[] cellValues)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cells">单元格集合</param>
        public void AddFootRow(params ICell[] cells)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 清空表头
        /// </summary>
        public void ClearHeader()
        {
            throw new NotImplementedException();
        }
    }
}
