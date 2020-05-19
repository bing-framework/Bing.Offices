using System;
using System.Collections.Generic;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 工作表(<see cref="ISheet"/> ) 扩展
    /// </summary>
    public static class SheetExtensions
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
    }
}
