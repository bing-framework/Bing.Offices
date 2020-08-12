using System.Collections.Generic;

namespace Bing.Offices.Metadata.Excels
{
    /// <summary>
    /// 工作表
    /// </summary>
    public interface ISheet
    {
        /// <summary>
        /// 首行行号。从1开始，如果没有行，则为0
        /// </summary>
        int FirstRowNum { get; }

        /// <summary>
        /// 尾行行号。从1开始，如果没有行，则为0
        /// </summary>
        int LastRowNum { get; }

        /// <summary>
        /// 获取单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        IRow GetRow(int rowIndex);

        /// <summary>
        /// 创建单元行
        /// </summary>
        /// <param name="rowIndex">行索引</param>
        IRow CreateRow(int rowIndex);

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        /// <param name="width">宽度</param>
        void SetColumnWidth(int columnIndex, int width);

        /// <summary>
        /// 设置自动列宽
        /// </summary>
        /// <param name="columnIndex">列索引</param>
        void AutoSizeColumn(int columnIndex);

        /// <summary>
        /// 创建冻结窗格
        /// </summary>
        /// <param name="colSplit">垂直分割位置</param>
        /// <param name="rowSplit">水平分割位置</param>
        /// <param name="leftMostCol">可见列数</param>
        /// <param name="topRow">可见行数</param>
        void CreateFreezePane(int colSplit, int rowSplit, int leftMostCol, int topRow);

        /// <summary>
        /// 设置自动筛选
        /// </summary>
        /// <param name="firstRowIndex">首行索引</param>
        /// <param name="lastRowIndex">尾行索引</param>
        /// <param name="firstColumnIndex">首列索引</param>
        /// <param name="lastColumnIndex">尾列索引</param>
        void SetAutoFilter(int firstRowIndex, int lastRowIndex, int firstColumnIndex, int lastColumnIndex);

        /// <summary>
        /// 移动行
        /// </summary>
        /// <param name="startRow">起始行索引</param>
        /// <param name="endRow">结束行索引</param>
        /// <param name="n">行数。正数: 下移, 负数: 上移</param>
        void ShiftRows(int startRow, int endRow, int n);

        /// <summary>
        /// 复制行
        /// </summary>
        /// <param name="sourceIndex">源索引</param>
        /// <param name="targetIndex">目标索引</param>
        IRow CopyRow(int sourceIndex, int targetIndex);

        /// <summary>
        /// 移除行
        /// </summary>
        /// <param name="row">单元行</param>
        void RemoveRow(IRow row);

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; }

        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 最大列数
        /// </summary>
        int MaxColumnCount { get; }

        /// <summary>
        /// 行数
        /// </summary>
        int RowCount { get; }

        /// <summary>
        /// 表头行数
        /// </summary>
        int HeadRowCount { get; }

        /// <summary>
        /// 正文行数
        /// </summary>
        int BodyRowCount { get; }

        /// <summary>
        /// 页脚行数
        /// </summary>
        int FootRowCount { get; }

        /// <summary>
        /// 获取表头
        /// </summary>
        IList<IRow> GetHeader();

        /// <summary>
        /// 获取正文
        /// </summary>
        IList<IRow> GetBody();

        /// <summary>
        /// 获取页脚
        /// </summary>
        IList<IRow> GetFooter();

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="titles">标题</param>
        void AddHeadRow(params string[] titles);

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="cells">单元格</param>
        void AddHeadRow(params ICell[] cells);

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cellValues">值</param>
        void AddBodyRow(params object[] cellValues);

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        void AddBodyRow(IEnumerable<ICell> cells);

        /// <summary>
        /// 添加正文
        /// </summary>
        /// <param name="cells">单元格集合</param>
        /// <param name="physicalRowIndex">物理行索引</param>
        void AddBodyRow(IEnumerable<ICell> cells, int physicalRowIndex);

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cellValues">值</param>
        void AddFootRow(params string[] cellValues);

        /// <summary>
        /// 添加页脚
        /// </summary>
        /// <param name="cells">单元格集合</param>
        void AddFootRow(params ICell[] cells);

        /// <summary>
        /// 清空表头
        /// </summary>
        void ClearHeader();
    }
}
