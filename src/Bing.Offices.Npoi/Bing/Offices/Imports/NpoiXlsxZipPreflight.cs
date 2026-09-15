using System.IO.Compression;
using System.Xml;
using Bing.Offices.Exceptions;

namespace Bing.Offices.Imports;

/// <summary>
/// 在 NPOI 创建 Workbook DOM 前检查 XLSX ZIP 资源预算。
/// </summary>
internal static class NpoiXlsxZipPreflight
{
    /// <summary>
    /// 在创建 Workbook DOM 前校验 XLSX ZIP 的条目、大小和 XML 安全预算。
    /// </summary>
    /// <param name="source">待预检且可定位的输入流。</param>
    /// <param name="limits">ZIP、工作表和 XML 资源限制。</param>
    /// <param name="cancellationToken">预检过程中检查的取消令牌。</param>
    /// <remarks>进入 ZIP 预检流程的可定位流在预检完成或失败后会恢复到开头。</remarks>
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken = default)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (limits == null || !source.CanSeek || source.Length < 4)
            return;
        if (!IsZip(source))
            return;
        source.Position = 0;
        try
        {
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var entryCount = archive.Entries.Count;
            if (limits.MaxZipEntries.HasValue && entryCount > limits.MaxZipEntries.Value)
                Throw($"XLSX ZIP entry 数量超过限制: {limits.MaxZipEntries.Value}");
            if (archive.GetEntry("xl/workbook.xml") == null)
                throw new BingOfficesImportException("XLSX ZIP 缺少 xl/workbook.xml。", null, "NPOI",
                    BingOfficesStage.Preflight);
            long totalUncompressed = 0;
            long totalWorksheetBytes = 0;
            var entryNames = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateEntryPath(entry.FullName);
                if (!entryNames.Add(entry.FullName))
                    Throw($"XLSX ZIP 存在重复 entry: {entry.FullName}");
                var uncompressed = entry.Length;
                var compressed = entry.CompressedLength;
                if (limits.MaxZipEntryUncompressedBytes.HasValue
                    && uncompressed > limits.MaxZipEntryUncompressedBytes.Value)
                    Throw($"XLSX ZIP entry 解压大小超过限制: {entry.FullName}");
                if (limits.MaxZipTotalUncompressedBytes.HasValue
                    && uncompressed > limits.MaxZipTotalUncompressedBytes.Value - totalUncompressed)
                    Throw($"XLSX ZIP 总解压大小超过限制: {limits.MaxZipTotalUncompressedBytes.Value}");
                totalUncompressed += uncompressed;
                if (limits.MaxZipCompressionRatio.HasValue && compressed == 0 && uncompressed > 0)
                    Throw($"XLSX ZIP entry 压缩比超过限制: {entry.FullName}");
                if (limits.MaxZipCompressionRatio.HasValue && compressed > 0
                    && (double)uncompressed / compressed > limits.MaxZipCompressionRatio.Value)
                    Throw($"XLSX ZIP entry 压缩比超过限制: {entry.FullName}");
                if (string.Equals(entry.FullName, "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase))
                    ValidateSize(uncompressed, limits.MaxSharedStringsBytes, "sharedStrings.xml");
                else if (string.Equals(entry.FullName, "xl/styles.xml", StringComparison.OrdinalIgnoreCase))
                    ValidateSize(uncompressed, limits.MaxStylesBytes, "styles.xml");
                else if (IsWorksheet(entry.FullName))
                {
                    ValidateSize(uncompressed, limits.MaxWorksheetBytes, entry.FullName);
                    if (limits.MaxTotalWorksheetBytes.HasValue
                        && uncompressed > limits.MaxTotalWorksheetBytes.Value - totalWorksheetBytes)
                        Throw($"XLSX worksheet XML 总大小超过限制: {limits.MaxTotalWorksheetBytes.Value}");
                    totalWorksheetBytes += uncompressed;
                }
                if (entry.Length > 0)
                    ValidateXmlSafety(entry, cancellationToken, limits.MaxXmlCharacters, limits.MaxXmlDepth);
            }
        }
        catch (BingOfficesResourceLimitException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new BingOfficesImportException("XLSX ZIP 预检失败。", exception, "NPOI",
                BingOfficesStage.Preflight);
        }
        catch (XmlException exception)
        {
            throw new BingOfficesImportException("XLSX XML 预检失败。", exception, "NPOI",
                BingOfficesStage.Preflight);
        }
        finally
        {
            source.Position = 0;
        }
    }

    /// <summary>
    /// 根据文件头判断输入流是否为 ZIP 格式。
    /// </summary>
    /// <param name="source">待判断且可定位的输入流。</param>
    /// <returns>文件头符合 ZIP 签名时为 true，否则为 false。</returns>
    private static bool IsZip(Stream source)
    {
        var originalPosition = source.Position;
        source.Position = 0;
        var header = new byte[4];
        var read = source.Read(header, 0, header.Length);
        source.Position = originalPosition;
        return read == 4 && header[0] == 0x50 && header[1] == 0x4B
            && ((header[2] == 0x03 && header[3] == 0x04)
                || (header[2] == 0x05 && header[3] == 0x06)
                || (header[2] == 0x07 && header[3] == 0x08));
    }

    /// <summary>
    /// 判断 ZIP 条目是否为工作表 XML 文件。
    /// </summary>
    /// <param name="name">ZIP 条目名称。</param>
    /// <returns>条目位于工作表目录且以 XML 扩展名结尾时为 true。</returns>
    private static bool IsWorksheet(string name) => name.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
        && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 校验 ZIP 条目路径不包含绝对路径或目录穿越片段。
    /// </summary>
    /// <param name="name">待校验的 ZIP 条目名称。</param>
    private static void ValidateEntryPath(string name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.StartsWith("/", StringComparison.Ordinal)
            || name.Contains("\\", StringComparison.Ordinal) || name.Contains("../", StringComparison.Ordinal)
            || name.Contains("..\\", StringComparison.Ordinal))
            Throw($"XLSX ZIP entry 路径无效: {name}");
    }

    /// <summary>
    /// 校验 XML 部件大小不超过配置上限。
    /// </summary>
    /// <param name="actual">实际字节数。</param>
    /// <param name="maximum">允许的最大字节数；为空时不限制。</param>
    /// <param name="name">超限部件名称。</param>
    private static void ValidateSize(long actual, long? maximum, string name)
    {
        if (maximum.HasValue && actual > maximum.Value)
            Throw($"XLSX XML 部件超过限制: {name}");
    }

    /// <summary>
    /// 读取 XML 条目并执行 DTD、字符数和嵌套深度安全校验。
    /// </summary>
    /// <param name="entry">待检查的 ZIP XML 条目。</param>
    /// <param name="cancellationToken">读取过程中检查的取消令牌。</param>
    /// <param name="maxCharacters">允许统计的最大字符数；为空时不限制。</param>
    /// <param name="maxDepth">允许的最大 XML 嵌套深度；为空时不限制。</param>
    private static void ValidateXmlSafety(ZipArchiveEntry entry, CancellationToken cancellationToken,
        long? maxCharacters, int? maxDepth)
    {
        if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            return;
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = entry.Open();
        using var reader = System.Xml.XmlReader.Create(stream, new System.Xml.XmlReaderSettings
        {
            DtdProcessing = System.Xml.DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersFromEntities = 0,
            MaxCharactersInDocument = 0
        });
        long characters = 0;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (maxDepth.HasValue && reader.Depth > maxDepth.Value)
                Throw($"XLSX XML 嵌套深度超过限制: {entry.FullName}");
            var nodeCharacters = (long)(reader.Name?.Length ?? 0) + (reader.Value?.Length ?? 0);
            if (reader.HasAttributes)
            {
                for (var index = 0; index < reader.AttributeCount; index++)
                {
                    reader.MoveToAttribute(index);
                    nodeCharacters += (reader.Name?.Length ?? 0) + (reader.Value?.Length ?? 0);
                }
                reader.MoveToElement();
            }
            if (maxCharacters.HasValue && nodeCharacters > maxCharacters.Value - characters)
                Throw($"XLSX XML 字符数量超过限制: {entry.FullName}");
            characters += nodeCharacters;
        }
    }

    /// <summary>
    /// 创建并抛出资源限制异常。
    /// </summary>
    /// <param name="message">资源限制原因。</param>
    private static void Throw(string message) => throw new BingOfficesResourceLimitException(message,
        provider: "NPOI", operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
}
