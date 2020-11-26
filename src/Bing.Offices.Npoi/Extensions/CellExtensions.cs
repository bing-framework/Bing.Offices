using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Bing.Helpers;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 单元格(<see cref="ICell"/>) 扩展
    /// </summary>
    public static class CellExtensions
    {
        #region GetStringValue(获取单元格的字符串值)

        /// <summary>
        /// 获取单元格的字符串值
        /// </summary>
        /// <param name="cell">单元格</param>
        public static string GetStringValue(this ICell cell)
        {
            var result = string.Empty;
            if (cell == null)
                return result;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        result = cell.StringCellValue;
                        break;
                    case CellType.Boolean:
                        result = cell.BooleanCellValue.ToString();
                        break;
                    case CellType.Error:
                        result = cell.ErrorCellValue.ToString();
                        break;
                    case CellType.Formula:
                        result = cell.CellFormula;
                        break;
                    case CellType.Numeric:
                        //result = DateUtil.IsCellDateFormatted(cell)
                        //    ? cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss.sss", CultureInfo.InvariantCulture)
                        //    : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        //日期是按1900/1/0作为0起点，相差的天数就是整数部分，小数部分是这样来的：24h*3600s/h=86400s，那么一天有86400秒，用1/86400*现在经过的秒数，就是小数部分
                        result = DateUtil.IsCellDateFormatted(cell)
                            ? DateTime.Parse("1900/1/1").AddDays(-2)
                                .AddDays((int)cell.NumericCellValue)
                                .AddMilliseconds(Math.Ceiling((cell.NumericCellValue - (int)cell.NumericCellValue) * 1000 * 86400))
                                .ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                            : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        break;
                    default:
                        result = cell.ToString();
                        break;
                }
                return result.Trim();
            }
            catch
            {
                return result;
            }
        }

        #endregion

        #region SetValue(设置单元格值)

        /// <summary>
        /// 设置单元值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="value">值</param>
        /// <param name="scale">保留小数位数</param>
        public static void SetValue(this ICell cell, object value,byte? scale = null)
        {
            if (cell == null)
                return;
            if (value == null)
            {
                cell.SetCellValue(string.Empty);
                return;
            }

            var type = value.GetType();
            if (!string.IsNullOrWhiteSpace(type.FullName) && type.FullName.Equals("System.Byte[]"))
            {
                var pictureIndex = cell.Sheet.Workbook.AddPicture(value as byte[], PictureType.PNG);
                var anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
                anchor.Col1 = cell.ColumnIndex;
                anchor.Col2 = cell.ColumnIndex + cell.GetSpan().Item2;
                anchor.Row1 = cell.RowIndex;
                anchor.Row2 = cell.RowIndex + cell.GetSpan().Item1;
                var drawing = cell.Sheet.CreateDrawingPatriarch();
                var picture = drawing.CreatePicture(anchor, pictureIndex);
                return;
            }

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.String:
                    cell.SetCellValue(Convert.ToString(value));
                    break;
                case TypeCode.DateTime:
                    cell.SetCellValue(Convert.ToDateTime(value));
                    break;
                case TypeCode.Boolean:
                    cell.SetCellValue(Convert.ToBoolean(value));
                    break;
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Byte:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                    cell.SetCellValue(Conv.ToDouble(value, scale));
                    break;
                default:
                    cell.SetCellValue(string.Empty);
                    break;
            }
        }

        /// <summary>
        /// 获取单元格跨度信息
        /// </summary>
        /// <param name="cell">单元格</param>
        private static Tuple<int, int> GetSpan(this ICell cell)
        {
            if (cell.IsMergedCell)
            {
                var regionsNum = cell.Sheet.NumMergedRegions;
                for (var i = 0; i < regionsNum; i++)
                {
                    var range = cell.Sheet.GetMergedRegion(i);
                    if (range.FirstRow != cell.RowIndex || range.FirstColumn != cell.ColumnIndex)
                        continue;
                    var rowSpan = range.LastRow - range.FirstRow + 1;
                    var columnSpan = range.LastColumn - range.FirstColumn + 1;
                    return new Tuple<int, int>(rowSpan, columnSpan);
                }
            }
            return new Tuple<int, int>(1, 1);
        }

        #endregion

        #region AddConditionalFormattingRules(添加条件格式规则)

        /// <summary>
        /// 添加条件格式规则
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="cfrs">条件格式规则</param>
        public static void AddConditionalFormattingRules(this ICell cell, IConditionalFormattingRule[] cfrs)
        {
            var regions = new[]
            {
                new CellRangeAddress(cell.RowIndex, cell.RowIndex, cell.ColumnIndex, cell.ColumnIndex),
            };
            cell.Sheet.SheetConditionalFormatting.AddConditionalFormatting(regions, cfrs);
        }

        #endregion

        #region GetConditionalFormattingRules(获取条件格式规则)

        /// <summary>
        /// 获取条件格式规则
        /// </summary>
        /// <param name="cell">单元格</param>
        public static IConditionalFormattingRule[] GetConditionalFormattingRules(this ICell cell)
        {
            var cfrs = new List<IConditionalFormattingRule>();
            var scf = cell.Sheet.SheetConditionalFormatting;
            for (var i = 0; i < scf.NumConditionalFormattings; i++)
            {
                var cf = scf.GetConditionalFormattingAt(i);
                if (cell.ExistConditionalFormatting(cf))
                {
                    for (var j = 0; j < cf.NumberOfRules; j++)
                    {
                        cfrs.Add(cf.GetRule(j));
                    }
                }
            }

            return cfrs.ToArray();
        }

        /// <summary>
        /// 判断单元格是否存在条件格式
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="cf">条件格式</param>
        private static bool ExistConditionalFormatting(this ICell cell, IConditionalFormatting cf)
        {
            return cf.GetFormattingRanges().Any(cra =>
                cell.RowIndex >= cra.FirstRow && cell.RowIndex <= cra.LastRow && cell.ColumnIndex >= cra.FirstColumn &&
                cell.ColumnIndex <= cra.LastColumn);
        }

        #endregion
    }
}
