using System.Globalization;
using System.Runtime.CompilerServices;
using Bing.Helpers;
using NPOI.SS.UserModel;

namespace Bing.Offices.Npoi.Extensions;

/// <summary>
/// NPOI单元格(<see cref="NPOI.SS.UserModel.ICell"/>) 扩展
/// </summary>
public static partial class CellExtensions
{
    /// <summary>
    /// IEEE 754 双精度可精确表示的最大整数。
    /// </summary>
    private const long MaxSafeInteger = 9007199254740991;

    /// <summary>
    /// 按工作簿保存的日期样式缓存，避免为每个日期单元格重复创建样式。
    /// </summary>
    private static readonly ConditionalWeakTable<IWorkbook, CellStyleCache> StyleCaches = new();

    #region GetStringValue(获取单元格的字符串值)

    /// <summary>
    /// 按公式缓存结果、日期格式和单元格类型将单元格转换为文本。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <returns>单元格文本；空单元格或空引用返回空字符串。</returns>
    public static string GetStringValue(this ICell cell)
    {
        if (cell == null)
            return string.Empty;

        var cellType = cell.CellType == CellType.Formula
            ? cell.CachedFormulaResultType
            : cell.CellType;
        switch (cellType)
        {
            case CellType.String:
                return cell.StringCellValue?.Trim();
            case CellType.Boolean:
                return cell.BooleanCellValue.ToString();
            case CellType.Error:
                return Convert.ToString(cell.ErrorCellValue, CultureInfo.InvariantCulture);
            case CellType.Numeric:
                return DateUtil.IsCellDateFormatted(cell)
                    ? Convert.ToDateTime(cell.DateCellValue, CultureInfo.InvariantCulture)
                        .ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
                    : Convert.ToString(cell.NumericCellValue, CultureInfo.InvariantCulture);
            case CellType.Blank:
                return string.Empty;
            default:
                return cell.ToString()?.Trim();
        }
    }

    #endregion

    #region SetValue(设置单元格值)

    /// <summary>
    /// 根据运行时类型将值写入单元格，并处理日期、枚举、图片和 Excel 精度边界。
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
            cell.SetCellType(CellType.Blank);
            return;
        }

        var type = value.GetType();
        if (!string.IsNullOrWhiteSpace(type.FullName) && type.FullName.Equals("System.Byte[]"))
        {
            var pictureBytes = value as byte[];
            var pictureIndex = cell.Sheet.Workbook.AddPicture(pictureBytes,
                PictureTypeResolver.Resolve(pictureBytes));
            var anchor = cell.Sheet.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = cell.ColumnIndex;
            anchor.Col2 = cell.ColumnIndex + cell.GetSpan().Item2;
            anchor.Row1 = cell.RowIndex;
            anchor.Row2 = cell.RowIndex + cell.GetSpan().Item1;
            var drawing = cell.Sheet.CreateDrawingPatriarch();
            var picture = drawing.CreatePicture(anchor, pictureIndex);
            return;
        }

        if (value is System.Enum enumValue)
        {
            cell.SetCellValue(enumValue.ToString());
            return;
        }

        if (value is Guid guidValue)
        {
            cell.SetCellValue(guidValue.ToString("D"));
            return;
        }

        if (value is Version versionValue)
        {
            cell.SetCellValue(versionValue.ToString());
            return;
        }

        if (value is long longValue && (longValue > MaxSafeInteger || longValue < -MaxSafeInteger))
        {
            cell.SetCellValue(longValue.ToString(CultureInfo.InvariantCulture));
            return;
        }

        if (value is ulong ulongValue && ulongValue > (ulong)MaxSafeInteger)
        {
            cell.SetCellValue(ulongValue.ToString(CultureInfo.InvariantCulture));
            return;
        }

        if (value is decimal decimalValue)
        {
            var scaledValue = scale.HasValue ? Math.Round(decimalValue, scale.Value) : decimalValue;
            if (!CanRoundTripDecimal(scaledValue))
            {
                cell.SetCellValue(scaledValue.ToString(CultureInfo.InvariantCulture));
                return;
            }
        }

        switch (Type.GetTypeCode(type))
        {
            case TypeCode.String:
                cell.SetCellValue(Convert.ToString(value));
                break;
            case TypeCode.DateTime:
                cell.SetDateTimeValue(Convert.ToDateTime(value, CultureInfo.InvariantCulture));
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
                cell.SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
                break;
        }
    }

    /// <summary>
    /// 以 Excel 日期数值写入单元格并设置显示格式。
    /// </summary>
    /// <param name="cell">单元格。</param>
    /// <param name="value">日期值。</param>
    /// <param name="formatter">日期格式。</param>
    private static void SetDateTimeValue(this ICell cell, DateTime value, string formatter = null)
    {
        cell.SetCellValue(value);
        var workbook = cell.Sheet.Workbook;
        var baseStyle = cell.CellStyle;
        var dataFormat = string.IsNullOrWhiteSpace(formatter)
            ? value.TimeOfDay == TimeSpan.Zero ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss"
            : formatter;
        var cacheKey = $"date:{baseStyle.Index}:{dataFormat}";
        var cache = StyleCaches.GetValue(workbook, _ => new CellStyleCache());
        cell.CellStyle = cache.GetOrAdd(cacheKey, () =>
        {
            var style = workbook.CreateCellStyle();
            style.CloneStyleFrom(baseStyle);
            style.DataFormat = workbook.CreateDataFormat().GetFormat(dataFormat);
            return style;
        });
    }

    /// <summary>
    /// 获取启用自动换行且不修改原始共享样式的单元格样式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <returns>启用自动换行的可复用单元格样式。</returns>
    public static ICellStyle GetStyleWithWrapText(this ICell cell)
    {
        if (cell.CellStyle.WrapText)
            return cell.CellStyle;
        return GetOrAddDerivedStyle(cell, $"wrap:{cell.CellStyle.Index}", style => style.WrapText = true);
    }

    /// <summary>
    /// 获取指定垂直对齐方式且不修改原始共享样式的单元格样式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="verticalAlignment">目标垂直对齐方式。</param>
    /// <returns>使用指定垂直对齐方式的可复用单元格样式。</returns>
    public static ICellStyle GetStyleWithVerticalAlignment(this ICell cell, VerticalAlignment verticalAlignment)
    {
        if (cell.CellStyle.VerticalAlignment == verticalAlignment)
            return cell.CellStyle;
        return GetOrAddDerivedStyle(cell, $"vertical:{cell.CellStyle.Index}:{verticalAlignment}",
            style => style.VerticalAlignment = verticalAlignment);
    }

    /// <summary>
    /// 从当前样式派生并缓存指定配置的单元格样式。
    /// </summary>
    /// <param name="cell">目标单元格。</param>
    /// <param name="cacheKey">样式缓存键。</param>
    /// <param name="configure">派生样式配置操作。</param>
    /// <returns>可复用的派生样式。</returns>
    private static ICellStyle GetOrAddDerivedStyle(ICell cell, string cacheKey, Action<ICellStyle> configure)
    {
        var workbook = cell.Sheet.Workbook;
        var baseStyle = cell.CellStyle;
        var cache = StyleCaches.GetValue(workbook, _ => new CellStyleCache());
        return cache.GetOrAdd(cacheKey, () =>
        {
            var style = workbook.CreateCellStyle();
            style.CloneStyleFrom(baseStyle);
            configure(style);
            return style;
        });
    }

    /// <summary>
    /// 单个工作簿的派生单元格样式缓存。
    /// </summary>
    private sealed class CellStyleCache
    {
        /// <summary>
        /// 已创建的日期样式。
        /// </summary>
        private readonly Dictionary<string, ICellStyle> _styles = new();

        /// <summary>
        /// 获取或创建指定基础样式和日期格式对应的单元格样式。
        /// </summary>
        /// <param name="key">样式缓存键。</param>
        /// <param name="styleFactory">创建样式的工厂。</param>
        /// <returns>可复用的日期单元格样式。</returns>
        public ICellStyle GetOrAdd(string key, Func<ICellStyle> styleFactory)
        {
            lock (_styles)
            {
                if (_styles.TryGetValue(key, out var style))
                    return style;
                style = styleFactory();
                _styles[key] = style;
                return style;
            }
        }
    }

    /// <summary>
    /// 检查 decimal 经过 Excel 双精度存储后能否精确还原。
    /// </summary>
    /// <param name="value">待检查的小数值。</param>
    private static bool CanRoundTripDecimal(decimal value)
    {
        var doubleValue = Convert.ToDouble(value, CultureInfo.InvariantCulture);
        return decimal.TryParse(doubleValue.ToString("R", CultureInfo.InvariantCulture), NumberStyles.Float,
                   CultureInfo.InvariantCulture, out var roundTrippedValue)
               && roundTrippedValue == value;
    }

    /// <summary>
    /// 获取单元格所在合并区域的行列跨度；非合并单元格返回（1，1）。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <returns>包含行跨度和列跨度的元组。</returns>
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
    /// 将值写入单元格，并使用默认格式处理日期、数字和特殊类型。
    /// </summary>
    /// <param name="cell">单元格</param>
    /// <param name="value">值</param>
    public static void SetCellValue(this ICell cell, object value) => cell.SetValue(value);

    /// <summary>
    /// 将值写入单元格，并在需要时使用指定格式创建派生样式。
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

        if (string.IsNullOrWhiteSpace(formatter))
        {
            cell.SetValue(value);
            return;
        }

        if (value is DateTime time)
        {
            cell.SetDateTimeValue(time, formatter);
        }
        else
        {
            if (IsNumeric(value))
            {
                cell.SetValue(value);
                cell.CellStyle = GetOrAddDerivedStyle(cell, $"format:{cell.CellStyle.Index}:{formatter}", style =>
                    style.DataFormat = cell.Sheet.Workbook.CreateDataFormat().GetFormat(formatter));
                return;
            }
            cell.SetCellValue(value is IFormattable formattable
                ? formattable.ToString(formatter, CultureInfo.InvariantCulture)
                : value.ToString());
        }
    }

    /// <summary>
    /// 判断值是否为 Excel 可写入的数值类型。
    /// </summary>
    private static bool IsNumeric(object value)
    {
        var typeCode = Type.GetTypeCode(value.GetType());
        return typeCode == TypeCode.Byte || typeCode == TypeCode.SByte || typeCode == TypeCode.Int16
               || typeCode == TypeCode.UInt16 || typeCode == TypeCode.Int32 || typeCode == TypeCode.UInt32
               || typeCode == TypeCode.Int64 || typeCode == TypeCode.UInt64 || typeCode == TypeCode.Single
               || typeCode == TypeCode.Double || typeCode == TypeCode.Decimal;
    }

    #endregion
}

/// <summary>
/// 根据图片文件签名解析 NPOI 图片类型。
/// </summary>
internal static class PictureTypeResolver
{
    /// <summary>
    /// 解析图片类型。
    /// </summary>
    /// <param name="pictureBytes">图片字节。</param>
    /// <returns>可识别类型；未知内容保持原有 PNG 回退。</returns>
    public static PictureType Resolve(byte[] pictureBytes)
    {
        if (pictureBytes?.Length >= 8
            && pictureBytes[0] == 0x89
            && pictureBytes[1] == 0x50
            && pictureBytes[2] == 0x4E
            && pictureBytes[3] == 0x47
            && pictureBytes[4] == 0x0D
            && pictureBytes[5] == 0x0A
            && pictureBytes[6] == 0x1A
            && pictureBytes[7] == 0x0A)
            return PictureType.PNG;
        if (pictureBytes?.Length >= 3
            && pictureBytes[0] == 0xFF
            && pictureBytes[1] == 0xD8
            && pictureBytes[2] == 0xFF)
            return PictureType.JPEG;
        if (pictureBytes?.Length >= 6
            && pictureBytes[0] == 0x47
            && pictureBytes[1] == 0x49
            && pictureBytes[2] == 0x46
            && pictureBytes[3] == 0x38
            && (pictureBytes[4] == 0x37 || pictureBytes[4] == 0x39)
            && pictureBytes[5] == 0x61)
            return PictureType.GIF;
        return PictureType.PNG;
    }
}
