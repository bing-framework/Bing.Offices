using System.Globalization;
using Bing.Helpers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
/// </summary>
public static partial class CellExtensions
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
            return result?.Trim();
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
    public static void SetValue(this ICell cell, object value, byte? scale = null)
    {
        if (cell is null)
            throw new ArgumentNullException(nameof(cell));
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

    /// <summary>
    /// 设置单元格值
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="value">值</param>
    public static void SetCellValue(this ICell cell, object value) => cell.SetCellValue(value, null);

    /// <summary>
    /// 设置单元格值
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="value">值</param>
    /// <param name="formatter">格式化字符串</param>
    public static void SetCellValue(this ICell cell, object value, string formatter)
    {
        if (cell is null)
            throw new ArgumentNullException(nameof(cell));
        if (value is null || DBNull.Value == value)
        {
            cell.SetCellType(CellType.Blank);
            return;
        }

        if (value is DateTime time)
        {
            cell.SetCellValue(string.IsNullOrWhiteSpace(formatter)
                ? time.Date == time ? time.ToString("yyyy-MM-dd") : time.ToString("yyyy-MM-dd HH:mm:ss")
                : time.ToString(formatter));
            cell.SetCellType(CellType.String);
        }
        else
        {
            var type = value.GetType();
            if (type == typeof(double) || 
                type == typeof(int) || 
                type == typeof(long) || 
                type == typeof(float) ||
                type == typeof(decimal))
            {
                cell.SetCellValue(Convert.ToDouble(value));
                cell.SetCellType(CellType.Numeric);
            }
            else if (type == typeof(bool))
            {
                cell.SetCellValue((bool)value);
                cell.SetCellType(CellType.Boolean);
            }
            else if (type == typeof(byte[]) && value is byte[] bytes)
            {
                cell.Sheet.TryAddPicture(cell.RowIndex, cell.ColumnIndex, bytes);
            }
            else
            {
                cell.SetCellValue(value is IFormattable val && !string.IsNullOrWhiteSpace(formatter)
                    ? val.ToString(formatter, CultureInfo.CurrentCulture)
                    : value.ToString());
                cell.SetCellType(CellType.String);
            }
        }
    }

    #endregion
}
