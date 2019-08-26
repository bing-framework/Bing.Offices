using System;
using System.Globalization;
using NPOI.SS.UserModel;

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
            if (cell == null)
                return string.Empty;
            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Error:
                        return cell.ErrorCellValue.ToString();
                    case CellType.Formula:
                        return cell.CellFormula;
                    case CellType.Numeric:
                        return DateUtil.IsCellDateFormatted(cell)
                            ? cell.DateCellValue.ToString(CultureInfo.InvariantCulture)
                            : cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                    default:
                        return cell.ToString();
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion
    }
}
