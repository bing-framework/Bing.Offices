using System.Globalization;

namespace Bing.Offices.Conversions;

/// <summary>DOM 与 XML 行读取共用的标量文本契约。</summary>
internal static class NpoiScalarText
{
    /// <summary>日期文本沿用既有导入格式，不返回 Excel 日期序列号。</summary>
    internal static string Date(DateTime value) => value.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
}
