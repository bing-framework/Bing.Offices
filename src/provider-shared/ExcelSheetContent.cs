using System.Globalization;

namespace Bing.Offices.Exports;

/// <summary>
/// 图片和原生校验的共享边界与地址转换。
/// </summary>
internal static class ExcelSheetContent
{
    /// <summary>
    /// 验证工作表图片和数据校验规则的格式边界。
    /// </summary>
    /// <param name="request">包含输出格式与工作表内容的导出请求。</param>
    internal static void Validate(ExcelWorkbookExportRequest request)
    {
        var rows = request.Format == ExcelFormat.Xls ? 65536 : 1048576;
        var columns = request.Format == ExcelFormat.Xls ? 256 : 16384;
        foreach (var sheet in request.Sheets)
        {
            foreach (var image in sheet.Images)
            {
                image.Validate();
                if (image.Row >= rows || image.Column >= columns)
                    throw new ArgumentOutOfRangeException(nameof(image), "图片锚点超出格式边界。");
            }
            foreach (var rule in sheet.DataValidations)
            {
                rule.Validate();
                if (rule.Range.EndRow >= rows || rule.Range.EndColumn >= columns)
                    throw new ArgumentOutOfRangeException(nameof(rule), "数据校验区域超出格式边界。");
            }
        }
    }

    /// <summary>
    /// 生成 A1 地址。
    /// </summary>
    /// <param name="row">零基行索引。</param>
    /// <param name="column">零基列索引。</param>
    /// <returns>不带工作表名称的单元格地址。</returns>
    internal static string Address(int row, int column)
    {
        var letters = string.Empty;
        for (var index = column + 1; index > 0; index /= 26)
        {
            index--;
            letters = (char)('A' + index % 26) + letters;
        }
        return letters + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 生成 A1 地址。
    /// </summary>
    /// <param name="range">包含起止位置的零基区域。</param>
    /// <returns>由冒号连接起止单元格的区域地址。</returns>
    internal static string Address(ExcelRangeDefinition range) =>
        Address(range.StartRow, range.StartColumn) + ":" + Address(range.EndRow, range.EndColumn);

    /// <summary>
    /// 将日期转换为工作簿日期序列值。
    /// </summary>
    /// <param name="value">待转换日期。</param>
    /// <param name="date1904">为 <see langword="true"/> 时使用 1904 日期系统，否则使用 1900 日期系统。</param>
    /// <returns>包含小数时间部分的日期序列值。</returns>
    internal static double DateSerial(DateTime value, bool date1904)
    {
        if (date1904) return (value - new DateTime(1904, 1, 1)).TotalDays;
        var serial = value.ToOADate();
        return value < new DateTime(1900, 3, 1) ? serial - 1 : serial;
    }

    /// <summary>
    /// 格式化数据校验规则的边界值。
    /// </summary>
    /// <param name="rule">包含数值或日期边界的校验规则。</param>
    /// <param name="second">为 <see langword="true"/> 时读取第二个边界，否则读取第一个边界。</param>
    /// <param name="date1904">是否使用 1904 日期系统；为 <see langword="false"/> 时使用 1900 日期系统。</param>
    /// <returns>采用固定区域性往返格式的边界值文本；第二个日期缺失时使用第一个日期。</returns>
    internal static string Value(ExcelDataValidationDefinition rule, bool second, bool date1904 = false) =>
        (rule.Type == ExcelDataValidationType.Date
            ? DateSerial((second ? rule.Date2 ?? rule.Date1 : rule.Date1).Value, date1904)
            : second ? rule.Value2 : rule.Value1).ToString("R", CultureInfo.InvariantCulture);
}
