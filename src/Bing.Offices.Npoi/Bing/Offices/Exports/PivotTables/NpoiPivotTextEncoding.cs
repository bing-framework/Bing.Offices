using System.Globalization;
using System.Text;
using System.Xml;

namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 编码透视缓存中的 SpreadsheetML 文本属性。
/// </summary>
internal static class NpoiPivotTextEncoding
{
    /// <summary>
    /// 将原始文本编码为 ST_Xstring 属性值。
    /// </summary>
    /// <param name="text">尚未进行 Xstring 转义的原始文本。</param>
    /// <returns>保留控制字符和字面转义序列的文本属性值。</returns>
    /// <remarks>
    /// 一次编码回车、换行、制表符和 XML 非法字符，并保护字面 _xHHHH_ 的首下划线。
    /// 不重复编码本次生成的转义序列；合法 Unicode 代理对保持原样。
    /// </remarks>
    internal static string Encode(string text)
    {
        if (text == null)
            throw new ArgumentNullException(nameof(text));
        var encoded = new StringBuilder(text.Length);
        for (var index = 0; index < text.Length; index++)
        {
            var character = text[index];
            if (character == '_' && IsEscapeSequence(text, index))
                encoded.Append("_x005F_");
            else if (char.IsHighSurrogate(character) && index + 1 < text.Length
                && char.IsLowSurrogate(text[index + 1]))
            {
                encoded.Append(character).Append(text[++index]);
            }
            else if (character is '\r' or '\n' or '\t' || !XmlConvert.IsXmlChar(character))
                encoded.Append("_x").Append(((int)character).ToString("X4", CultureInfo.InvariantCulture)).Append('_');
            else
                encoded.Append(character);
        }
        return encoded.ToString();
    }

    /// <summary>
    /// 判断原始文本当前位置是否为字面 Xstring 转义序列。
    /// </summary>
    /// <param name="text">原始文本。</param>
    /// <param name="index">待检查的首下划线位置。</param>
    /// <returns>true 表示须保护该字面序列，false 表示普通文本。</returns>
    private static bool IsEscapeSequence(string text, int index) => index + 6 < text.Length
        && text[index + 1] == 'x' && text[index + 6] == '_'
        && Uri.IsHexDigit(text[index + 2]) && Uri.IsHexDigit(text[index + 3])
        && Uri.IsHexDigit(text[index + 4]) && Uri.IsHexDigit(text[index + 5]);
}
