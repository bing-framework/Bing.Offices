using System.Globalization;
using System.Text.RegularExpressions;

namespace Bing.Offices.Internals;

/// <summary>
/// 校验透视模板的字面来源区域。
/// </summary>
internal static class ExcelPivotTemplateSource
{
    /// <summary>
    /// XLSX 单格或矩形绝对、相对地址格式。
    /// </summary>
    private static readonly Regex RangePattern = new(@"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// 校验已有缓存的字面来源。
    /// </summary>
    /// <param name="sheet">来源工作表。</param>
    /// <param name="reference">字面区域。</param>
    /// <param name="sheets">实际工作表名称。</param>
    /// <param name="part">所属缓存部件。</param>
    internal static void Validate(string sheet, string reference, ISet<string> sheets, string part)
    {
        if (sheet == null || !sheets.Contains(sheet) || reference == null)
            throw new InvalidDataException($"缓存来源工作表或区域缺失: {part}");
        var match = RangePattern.Match(reference);
        if (!match.Success)
            throw new InvalidDataException($"缓存来源必须是有界区域: {part}/{reference}");
        var firstColumn = Column(match.Groups[1].Value);
        var firstRow = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        var lastColumn = match.Groups[3].Success ? Column(match.Groups[3].Value) : firstColumn;
        var lastRow = match.Groups[4].Success ? int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) : firstRow;
        if (lastRow < firstRow || lastColumn < firstColumn || lastRow > 1048576 || lastColumn > 16384)
            throw new InvalidDataException($"缓存来源超出 XLSX 区域边界: {part}/{reference}");
    }

    /// <summary>
    /// 将列字母转换为一基列号。
    /// </summary>
    /// <param name="letters">列字母。</param>
    /// <returns>列号。</returns>
    private static int Column(string letters)
    {
        var column = 0;
        foreach (var letter in letters)
            column = column * 26 + char.ToUpperInvariant(letter) - 'A' + 1;
        return column;
    }
}
