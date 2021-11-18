using System.Collections.Generic;
using Bing.Offices.Metadata;
using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// NPOI工作表(<see cref="NPOI.SS.UserModel.ISheet"/>) 扩展
    /// </summary>
    public static partial class SheetExtensions
    {
        /// <summary>
        /// 添加合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="regionInfo">合并区域信息</param>
        public static void AddMergedRegion(this NPOI.SS.UserModel.ISheet sheet, MergedRegionInfo regionInfo)
        {
            var region = new CellRangeAddress(regionInfo.FirstRow, regionInfo.LastRow, regionInfo.FirstCol,
                regionInfo.LastCol);
            sheet.AddMergedRegion(region);
        }

        /// <summary>
        /// 获取工作表中包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this NPOI.SS.UserModel.ISheet sheet) => sheet.GetMergedRegionInfos(null, null, null, null);

        /// <summary>
        /// 获取工作表中指定区域包含合并区域的信息列表
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="minRow">最小行索引</param>
        /// <param name="maxRow">最大行索引</param>
        /// <param name="minCol">最小列索引</param>
        /// <param name="maxCol">最大列索引</param>
        /// <param name="onlyInternal">是否仅在内部</param>
        public static List<MergedRegionInfo> GetMergedRegionInfos(this NPOI.SS.UserModel.ISheet sheet, int? minRow,
            int? maxRow, int? minCol, int? maxCol, bool onlyInternal = true)
        {
            var regionInfoList = new List<MergedRegionInfo>();
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn, onlyInternal))
                    regionInfoList.Add(new MergedRegionInfo(i, range.FirstRow, range.LastRow, range.FirstColumn, range.LastColumn));
            }
            return regionInfoList;
        }

        /// <summary>
        /// 移除工作表中所有的合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        public static void RemoveMergedRegions(this NPOI.SS.UserModel.ISheet sheet) => sheet.RemoveMergedRegions(null, null, null, null);

        /// <summary>
        /// 移除工作表中指定区域内的合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="minRow">最小行索引</param>
        /// <param name="maxRow">最大行索引</param>
        /// <param name="minCol">最小列索引</param>
        /// <param name="maxCol">最大列索引</param>
        public static void RemoveMergedRegions(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow,
            int? minCol, int? maxCol)
        {
            List<MergedRegionInfo> regionInfos;
            do
            {
                regionInfos = sheet.GetMergedRegionInfos(minRow, maxRow, minCol, maxCol);
                foreach (var regionInfo in regionInfos)
                    sheet.RemoveMergedRegion(regionInfo.Index);
            } while (regionInfos.Count > 0);
        }

        /// <summary>
        /// 移动工作表中所有合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="moveRowCount">移动行数</param>
        /// <param name="moveColCount">移动列数</param>
        public static void MoveMergedRegions(this NPOI.SS.UserModel.ISheet sheet, int moveRowCount,
            int moveColCount = 0) => sheet.MoveMergedRegions(null, null, null, null, moveRowCount, moveColCount);

        /// <summary>
        /// 移动工作表指定区域内的合并区域
        /// </summary>
        /// <param name="sheet">NPOI工作表</param>
        /// <param name="minRow">最小行索引</param>
        /// <param name="maxRow">最大行索引</param>
        /// <param name="minCol">最小列索引</param>
        /// <param name="maxCol">最大列索引</param>
        /// <param name="moveRowCount">移动行数</param>
        /// <param name="moveColCount">移动列数</param>
        public static void MoveMergedRegions(this NPOI.SS.UserModel.ISheet sheet, int? minRow, int? maxRow, int? minCol,
            int? maxCol, int moveRowCount, int moveColCount = 0)
        {
            for (var i = 0; i < sheet.NumMergedRegions; i++)
            {
                var range = sheet.GetMergedRegion(i);
                if (IsInternalOrIntersect(minRow, maxRow, minCol, maxCol, range.FirstRow, range.LastRow,
                    range.FirstColumn, range.LastColumn, true))
                {
                    range.FirstRow += moveRowCount;
                    range.LastRow += moveRowCount;
                    range.FirstColumn += moveColCount;
                    range.LastColumn += moveColCount;
                }
            }
        }
    }
}
