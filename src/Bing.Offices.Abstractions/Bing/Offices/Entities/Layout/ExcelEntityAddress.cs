using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities.Layout;

/// <summary>
/// 提供 Excel A1 单元格和矩形区域地址的解析与格式化。
/// </summary>
internal static class ExcelEntityAddress
{
    /// <summary>
    /// XLSX 工作表允许的最大行数。
    /// </summary>
    private const int MaxRows = 1048576;
    /// <summary>
    /// XLSX 工作表允许的最大列数。
    /// </summary>
    private const int MaxColumns = 16384;

    /// <summary>
    /// 解析并验证单个 A1 单元格地址。
    /// </summary>
    /// <param name="address">待解析的 A1 地址。</param>
    /// <returns>与地址对应的零基单元格坐标。</returns>
    internal static ExcelEntityCellReference Parse(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("单元格地址不能为空。", nameof(address));
        var text = address.Trim().Replace("$", string.Empty);
        var split = 0;
        while (split < text.Length && ((text[split] >= 'A' && text[split] <= 'Z')
            || (text[split] >= 'a' && text[split] <= 'z')))
            split++;
        if (split == 0 || split == text.Length || !int.TryParse(text.Substring(split),
                System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture,
                out var row) || row <= 0 || row > MaxRows)
            throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
        var column = 0;
        foreach (var character in text.Substring(0, split).ToUpperInvariant())
        {
            if (character < 'A' || character > 'Z')
                throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
            if (column > (MaxColumns - (character - 'A' + 1)) / 26)
                throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
            column = column * 26 + character - 'A' + 1;
        }
        if (column <= 0 || column > MaxColumns)
            throw new ArgumentException($"无效的 A1 地址: {address}", nameof(address));
        return new ExcelEntityCellReference(row - 1, column - 1);
    }

    /// <summary>
    /// 解析并验证 A1 单元格或矩形区域地址。
    /// </summary>
    /// <param name="range">待解析的 A1 地址。</param>
    /// <returns>与地址对应的零基矩形区域。</returns>
    internal static ExcelEntityCellRange ParseRange(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            throw new ArgumentException("区域地址不能为空。", nameof(range));
        var parts = range.Split(':');
        if (parts.Length > 2)
            throw new ArgumentException($"无效的 A1 区域地址: {range}", nameof(range));
        var first = Parse(parts[0]);
        var last = parts.Length == 1 ? first : Parse(parts[1]);
        return new ExcelEntityCellRange(first, last);
    }

    /// <summary>
    /// 将零基行列坐标格式化为 A1 单元格地址。
    /// </summary>
    /// <param name="row">从零开始的行索引。</param>
    /// <param name="column">从零开始的列索引。</param>
    /// <returns>格式化后的 A1 单元格地址。</returns>
    internal static string Format(int row, int column)
    {
        var value = column + 1;
        var letters = string.Empty;
        while (value > 0)
        {
            var remainder = (value - 1) % 26;
            letters = (char)('A' + remainder) + letters;
            value = (value - 1) / 26;
        }
        return letters + (row + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
    }
}
