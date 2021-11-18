using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 工作表(<see cref="ISheet"/> ) 扩展
    /// </summary>
    public static partial class SheetExtensions
    {
        /// <summary>
        /// 获取所有合并单元格区域。格式：(x1,y1,x2,y2)
        /// </summary>
        /// <param name="sheet">工作表</param>
        public static int[,] GetAllMergedRegions(this ISheet sheet)
        {
            // 工作表合并单元格数量
            var mergedRegions = sheet.NumMergedRegions;
            var output = new int[mergedRegions, 4];
            for (var i = 0; i < mergedRegions; i++)
            {
                var cellRangeAddress = sheet.GetMergedRegion(i);
                output[i, 0] = cellRangeAddress.FirstRow;
                output[i, 1] = cellRangeAddress.FirstColumn;
                output[i, 2] = cellRangeAddress.LastRow;
                output[i, 3] = cellRangeAddress.LastColumn;
            }
            return output;
        }

        /// <summary>
        /// 删除行。
        /// 解决因为shiftRows上移删除行而造成的格式错乱问题
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="deleteRowStartIndex">删除行起始索引。从0开始</param>
        /// <param name="count">删除行数</param>
        public static void DeleteRows(this ISheet sheet, int deleteRowStartIndex, int count)
        {
            var mergedRegions = sheet.NumMergedRegions;
            for (var i = mergedRegions - 1; i >= 0; i--)
            {
                var cellRangeAddress = sheet.GetMergedRegion(i);
                if (cellRangeAddress.FirstRow >= deleteRowStartIndex + count
                    || cellRangeAddress.LastRow <= deleteRowStartIndex)
                {
                    // 只有一行的合并单元格 FirstRow==LastRow
                    if (cellRangeAddress.FirstRow == cellRangeAddress.LastRow)
                    {
                        // 刚好在删除区域的 StartRow 或 EndRow
                        if (cellRangeAddress.FirstRow == deleteRowStartIndex
                            || cellRangeAddress.FirstRow == deleteRowStartIndex + count - 1)
                        {
                            sheet.RemoveMergedRegion(i);
                        }
                    }
                }
                else
                {
                    // 该合并单元格的行已经在被删除的区域内，所以解除该合并单元格
                    // 如果不删除该合并单元格将会导致全部格式错乱
                    sheet.RemoveMergedRegion(i);
                }
            }
            sheet.ShiftRows(deleteRowStartIndex + count, sheet.LastRowNum, -count, true, false);
        }

        /// <summary>
        /// 获取有数据（非空行）的最后一行的行数。如果sheet中一行数据都没有则返回-1，只有第一行有数据则返回0，最后有数据的行是第n行则返回n-1。
        /// </summary>
        /// <param name="sheet">工作表</param>
        public static int GetHasDataRowNum(this ISheet sheet)
        {
            for (var i = sheet.LastRowNum; i >= 0; i--)
            {
                if (!sheet.GetRow(i).IsEmptyRow())
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// 添加行
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="sheet">工作表</param>
        /// <param name="data">数据</param>
        /// <param name="action">操作</param>
        public static void AddRow<T>(this ISheet sheet, IEnumerable<T> data, Action<IRow, T> action)
        {
            var index = 1;
            foreach (var item in data)
            {
                var row = sheet.CreateRow(index);
                row.Height = 20 * 20;
                action(row, item);
                index++;
            }
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        /// <returns>NPOI单元行</returns>
        public static NPOI.SS.UserModel.IRow InsertRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.InsertRows(rowIndex, 1).FirstOrDefault();

        /// <summary>
        /// 批量插入行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="rowsCount">插入行数量</param>
        /// <returns>NPOI单元行数组</returns>
        public static NPOI.SS.UserModel.IRow[] InsertRows(this NPOI.SS.UserModel.ISheet sheet, int rowIndex,
            int rowsCount)
        {
            if (rowIndex <= sheet.LastRowNum)
                sheet.ShiftRows(rowIndex, sheet.LastRowNum, rowsCount, true, false);
            var rows = new List<NPOI.SS.UserModel.IRow>();
            for (var i = 0; i < rowsCount; i++)
            {
                var row = sheet.CreateRow(rowIndex + i);
                rows.Add(row);
            }
            return rows.ToArray();
        }

        /// <summary>
        /// 移除行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="rowIndex">行索引</param>
        public static int RemoveRow(this NPOI.SS.UserModel.ISheet sheet, int rowIndex) => sheet.RemoveRows(rowIndex, rowIndex);

        /// <summary>
        /// 移除行
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="endRowIndex">结束行索引</param>
        public static int RemoveRows(this NPOI.SS.UserModel.ISheet sheet, int startRowIndex, int endRowIndex)
        {
            var span = endRowIndex - startRowIndex + 1;
            sheet.RemoveMergedRegions(startRowIndex, endRowIndex, null, null);
            sheet.RemovePictures(startRowIndex, endRowIndex, null, null);
            for (var i = endRowIndex; i >= startRowIndex; i--)
            {
                var row = sheet.GetRow(i);
                sheet.RemoveRow(row);
            }
            if (endRowIndex + 1 <= sheet.LastRowNum)
            {
                sheet.ShiftRows(endRowIndex + 1, sheet.LastRowNum, -span, true, false);
                sheet.MovePictures(endRowIndex + 1, null, null, null, moveRowCount: -span);
            }
            return span;
        }

        /// <summary>
        /// 判断指定区域是否在内部或交叉
        /// </summary>
        /// <param name="rangeMinRow">区域最小行索引</param>
        /// <param name="rangeMaxRow">区域最大行索引</param>
        /// <param name="rangeMinCol">区域最小列索引</param>
        /// <param name="rangeMaxCol">区域最大列索引</param>
        /// <param name="targetMinRow">目标最小行索引</param>
        /// <param name="targetMaxRow">目标最大行索引</param>
        /// <param name="targetMinCol">目标最小列索引</param>
        /// <param name="targetMaxCol">目标最大列索引</param>
        /// <param name="onlyInternal">仅在内部</param>
        private static bool IsInternalOrIntersect(int? rangeMinRow, int? rangeMaxRow, int? rangeMinCol,
            int? rangeMaxCol, int targetMinRow, int targetMaxRow, int targetMinCol, int targetMaxCol, bool onlyInternal)
        {
            var tempMinRow = rangeMinRow ?? targetMinRow;
            var tempMaxRow = rangeMaxRow ?? targetMaxRow;
            var tempMinCol = rangeMinCol ?? targetMinCol;
            var tempMaxCol = rangeMaxCol ?? targetMaxCol;
            if (onlyInternal)
            {
                return tempMinRow <= targetMinRow &&
                       tempMaxRow >= targetMaxRow &&
                       tempMinCol <= targetMinCol &&
                       tempMaxCol >= targetMaxCol;
            }

            return Math.Abs(tempMaxRow - tempMinRow) + Math.Abs(targetMaxRow - targetMinRow) >=
                   Math.Abs(tempMaxRow + tempMinRow - targetMaxRow - targetMinRow) &&
                   Math.Abs(tempMaxCol - tempMinCol) + Math.Abs(targetMaxCol - targetMinCol) >=
                   Math.Abs(tempMaxCol + tempMinCol - targetMaxCol - targetMinCol);
        }
    }
}
