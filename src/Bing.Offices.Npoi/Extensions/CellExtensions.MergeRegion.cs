using Bing.Offices.Metadata;
using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// 单元格合并和合并区域解析扩展。
/// </summary>
internal static partial class CellExtensions
{
    /// <summary>
    /// 合并两个单元格所在区域，并按选项删除或吸收相交的既有合并区域。
    /// </summary>
    /// <param name="fromCell">起始单元格</param>
    /// <param name="toCell">终止单元格</param>
    /// <param name="isExpand">是否将相交既有区域的边界吸收到新合并区域中。</param>
    public static void Merge(this NPOI.SS.UserModel.ICell fromCell, NPOI.SS.UserModel.ICell toCell,
        bool isExpand = false)
    {
        if (!fromCell.Sheet.Equals(toCell.Sheet))
            throw new InvalidOperationException("单元格不在同一个工作表上");
        var sheet = fromCell.Sheet;
        var fromRange = fromCell.GetRangeInfo();
        var toRange = toCell.GetRangeInfo();
        var firstRowIndex = Math.Min(fromRange.FirstRow, toRange.FirstRow);
        var firstColIndex = Math.Min(fromRange.FirstCol, toRange.FirstCol);
        var lastRowIndex = Math.Max(fromRange.LastRow, toRange.LastRow);
        var lastColIndex = Math.Max(fromRange.LastCol, toRange.LastCol);
        var regionInfoList = sheet.GetMergedRegionInfos(firstRowIndex, lastRowIndex, firstColIndex, lastColIndex, false);
        foreach (var regionInfo in regionInfoList.OrderByDescending(x => x.Index))
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
    /// 获取单元格所在的合并区域；未合并时返回仅包含该单元格的区域信息。
    /// </summary>
    /// <param name="cell">NPOI单元格</param>
    /// <returns>合并区域信息；未合并单元格的区域索引为 -1。</returns>
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
