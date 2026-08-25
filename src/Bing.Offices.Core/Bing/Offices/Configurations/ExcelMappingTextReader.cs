using System;
using System.IO;
using System.Text;

namespace Bing.Offices.Configurations;

internal static class ExcelMappingTextReader
{
    internal const int MaxDocumentBytes = 1024 * 1024;

    internal static string ReadLimitedText(TextReader reader)
    {
        var buffer = new char[8192];
        var builder = new StringBuilder();
        var total = 0;
        int count;
        while ((count = reader.Read(buffer, 0, buffer.Length)) > 0)
        {
            total += count;
            if (total > MaxDocumentBytes)
                throw new InvalidOperationException($"配置超过最大字符数: {MaxDocumentBytes}");
            builder.Append(buffer, 0, count);
        }
        return builder.ToString();
    }

    internal static void ValidateDocumentText(string text, string format)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException($"{format} 配置不能为空。", nameof(text));
        if (Encoding.UTF8.GetByteCount(text) > MaxDocumentBytes)
            throw new InvalidOperationException($"{format} 配置超过最大字节数: {MaxDocumentBytes}");
    }
}
