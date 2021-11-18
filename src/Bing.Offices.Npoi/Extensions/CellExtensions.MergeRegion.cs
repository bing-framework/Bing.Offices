using System;
using Bing.Offices.Exceptions;
using Bing.Offices.Metadata;
using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
    /// </summary>
    public static partial class CellExtensions
    {
        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="fromCell">起始单元格</param>
        /// <param name="toCell">终止单元格</param>
        /// <param name="isExpand">扩充模式</param>
        public static void Merge(this NPOI.SS.UserModel.ICell fromCell, NPOI.SS.UserModel.ICell toCell,
            bool isExpand = false)
        {
            if (!fromCell.Sheet.Equals(toCell.Sheet))
                throw new OfficeException("单元格不在同一个工作表上");
            var sheet = fromCell.Sheet;
            var fromRange = fromCell.GetRangeInfo();
            var toRange = toCell.GetRangeInfo();
            var firstRowIndex = Math.Min(fromRange.FirstRow, toRange.FirstRow);
            var firstColIndex = Math.Min(fromRange.FirstCol, toRange.FirstCol);
            var lastRowIndex = Math.Max(fromRange.LastRow, toRange.LastRow);
            var lastColIndex = Math.Max(fromRange.LastCol, toRange.LastCol);
            var regionInfoList = sheet.GetMergedRegionInfos(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex, false);
            foreach (var regionInfo in regionInfoList)
            {
                if (isExpand)
                {
                    firstRowIndex = Math.Min(firstRowIndex, regionInfo.FirstRow);
                    firstColIndex = Math.Min(firstColIndex, regionInfo.FirstCol);
                    lastRowIndex = Math.Max(lastRowIndex, regionInfo.LastRow);
                    lastColIndex = Math.Max(lastColIndex, regionInfo.LastCol);
                }
                sheet.RemoveMergedRegion(regionInfo.Index);
            }
            var region = new CellRangeAddress(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex);
            fromCell.Sheet.AddMergedRegion(region);
        }

        /// <summary>
        /// 获取合并区域信息
        /// </summary>
        /// <param name="cell">NPOI单元格</param>
        private static MergedRegionInfo GetRangeInfo(this NPOI.SS.UserModel.ICell cell)
        {
            var sheet = cell.Sheet;
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (range.IsInRange(cell.RowIndex, cell.ColumnIndex))
                    return new MergedRegionInfo(i, range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn);
            }
            return new MergedRegionInfo(-1, cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex);
        }
    }
}
