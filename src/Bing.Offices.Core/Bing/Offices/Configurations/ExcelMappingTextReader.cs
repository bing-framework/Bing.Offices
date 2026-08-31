using System;
using System.IO;
using System.Text;

namespace Bing.Offices.Configurations;

/// <summary>读取并限制映射配置文本大小的内部工具。</summary>
internal static class ExcelMappingTextReader
{
    /// <summary>映射配置允许读取的最大 UTF-8 字节数及最大字符数。</summary>
    internal const int MaxDocumentBytes = 1024 * 1024;

    /// <summary>从文本读取器读取配置，并在超过大小上限时立即终止。</summary>
    /// <param name="reader">配置文本读取器。</param>
    /// <returns>读取到的完整配置文本。</returns>
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

    /// <summary>验证配置文本非空且 UTF-8 编码后的大小未超过限制。</summary>
    /// <param name="text">待验证的配置文本。</param>
    /// <param name="format">配置格式名称，用于生成错误消息。</param>
    internal static void ValidateDocumentText(string text, string format)
    {
        if (string.IsNullOrWhiteSpace(text))
            throw new ArgumentException($"{format} 配置不能为空。", nameof(text));
        if (Encoding.UTF8.GetByteCount(text) > MaxDocumentBytes)
            throw new InvalidOperationException($"{format} 配置超过最大字节数: {MaxDocumentBytes}");
    }
}
