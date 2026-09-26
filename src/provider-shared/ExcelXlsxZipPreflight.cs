using System.IO.Compression;
using System.Xml;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;

namespace Bing.Offices.IO;

/// <summary>
/// 执行 XLSX ZIP/XML 资源预检。
/// </summary>
/// <remarks>
/// MiniExcel 与 NPOI 项目通过编译链接复用本实现。
/// </remarks>
internal static class ExcelXlsxZipPreflight
{
    /// <summary>
    /// 读取工作簿是否声明使用 Excel 1904 日期系统。
    /// </summary>
    /// <param name="source">可定位的 XLSX 流；不可定位或长度不足时返回 false。</param>
    /// <param name="cancellationToken">读取过程使用的取消令牌。</param>
    /// <returns>WorkbookPr 的 date1904 值为 1、true 或 on 时返回 true；缺失或为其他值时返回 false。</returns>
    /// <remarks>
    /// 方法从流头读取 xl/workbook.xml，并在读取完成后恢复调用前的流位置。
    /// </remarks>
    internal static bool GetDate1904(Stream source, CancellationToken cancellationToken = default)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (!source.CanSeek || source.Length < 4)
            return false;

        var originalPosition = source.Position;
        source.Position = 0;
        try
        {
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var workbook = archive.GetEntry("xl/workbook.xml");
            if (workbook == null)
                return false;
            using var workbookStream = workbook.Open();
            using var reader = XmlReader.Create(workbookStream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true
            });
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element
                    || !string.Equals(reader.LocalName, "workbookPr", StringComparison.Ordinal))
                    continue;
                var value = reader.GetAttribute("date1904");
                return string.Equals(value, "1", StringComparison.Ordinal)
                    || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 按资源限制检查工作簿的 XLSX ZIP 结构和 XML 内容。
    /// </summary>
    /// <param name="source">待检查的可定位工作簿流。</param>
    /// <param name="limits">ZIP/XML 条目、解压大小、压缩比和 XML 深度等资源限制；为 null 时跳过检查。</param>
    /// <param name="provider">用于异常上下文的提供程序名称。</param>
    /// <param name="requireZip">是否要求输入具有 XLSX ZIP 文件头；为 false 时非 ZIP 流直接跳过 ZIP 检查。</param>
    /// <param name="cancellationToken">用于取消预检的令牌。</param>
    /// <param name="includePictures">是否扫描并应用图片数量与大小限制；仅由支持图片资源 admission 的 Provider 开启。</param>
    /// <param name="enforceSheetLimit">是否在 ZIP 预检阶段执行工作表数量限制。</param>
    /// <param name="enforceCellLimits">是否在 ZIP 预检阶段执行物理列数和单元格数量限制。</param>
    /// <param name="requireWorkbookXml">是否要求 OOXML 工作簿部件；XLSB 等 ZIP 容器可关闭该要求。</param>
    /// <remarks>
    /// 检查 ZIP 条目数量、路径、重复项、解压大小、压缩比、指定 XML 部件大小以及 XML 字符数和嵌套深度；进入检查后可定位流会在结束时置于文件头。
    /// </remarks>
    internal static void Validate(Stream source, ExcelResourceLimits limits, string provider,
        bool requireZip, CancellationToken cancellationToken = default, bool includePictures = false,
        bool enforceSheetLimit = true, bool enforceCellLimits = true, bool requireWorkbookXml = true)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (limits == null || !source.CanSeek || source.Length < 4)
            return;
        source.Position = 0;
        try
        {
            if (!IsZip(source))
            {
                if (requireZip)
                    throw new BingOfficesImportException($"{provider} 仅支持 XLSX ZIP 工作簿。", null,
                        provider, BingOfficesStage.Preflight);
                return;
            }
            source.Position = 0;
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            if (limits.MaxZipEntries.HasValue && archive.Entries.Count > limits.MaxZipEntries.Value)
                Resource($"XLSX ZIP entry 数量超过限制: {limits.MaxZipEntries.Value}", provider);
            if (requireWorkbookXml && archive.GetEntry("xl/workbook.xml") == null)
                throw new BingOfficesImportException("XLSX ZIP 缺少 xl/workbook.xml。", null,
                    provider, BingOfficesStage.Preflight);

            long totalUncompressed = 0;
            long totalWorksheetBytes = 0;
            long totalPhysicalCells = 0;
            var logicalSheetCount = 0;
            var maximumColumnsPerSheet = 0;
            var pictureCount = 0;
            long totalPictureBytes = 0;
            var entryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in archive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ValidateEntryPath(entry.FullName, provider);
                if (!entryNames.Add(entry.FullName))
                    Resource($"XLSX ZIP 存在重复 entry: {entry.FullName}", provider);
                var uncompressed = entry.Length;
                var compressed = entry.CompressedLength;
                if (limits.MaxZipEntryUncompressedBytes.HasValue
                    && uncompressed > limits.MaxZipEntryUncompressedBytes.Value)
                    Resource($"XLSX ZIP entry 解压大小超过限制: {entry.FullName}", provider);
                if (limits.MaxZipTotalUncompressedBytes.HasValue
                    && uncompressed > limits.MaxZipTotalUncompressedBytes.Value - totalUncompressed)
                    Resource($"XLSX ZIP 总解压大小超过限制: {limits.MaxZipTotalUncompressedBytes.Value}", provider);
                totalUncompressed += uncompressed;
                if (limits.MaxZipCompressionRatio.HasValue && compressed == 0 && uncompressed > 0)
                    Resource($"XLSX ZIP entry 压缩比超过限制: {entry.FullName}", provider);
                if (limits.MaxZipCompressionRatio.HasValue && compressed > 0
                    && (double)uncompressed / compressed > limits.MaxZipCompressionRatio.Value)
                    Resource($"XLSX ZIP entry 压缩比超过限制: {entry.FullName}", provider);
                if (string.Equals(entry.FullName, "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase))
                    ValidateSize(uncompressed, limits.MaxSharedStringsBytes, "sharedStrings.xml", provider);
                else if (string.Equals(entry.FullName, "xl/styles.xml", StringComparison.OrdinalIgnoreCase))
                    ValidateSize(uncompressed, limits.MaxStylesBytes, "styles.xml", provider);
                else if (string.Equals(entry.FullName, "xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                {
                    logicalSheetCount = CountWorkbookSheets(entry, cancellationToken, provider);
                    if (enforceSheetLimit && limits.MaxSheets.HasValue
                        && logicalSheetCount > limits.MaxSheets.Value)
                        Resource($"XLSX Sheet 数量超过限制: {limits.MaxSheets.Value}", provider);
                }
                else if (IsWorksheet(entry.FullName))
                {
                    ValidateSize(uncompressed, limits.MaxWorksheetBytes, entry.FullName, provider);
                    if (limits.MaxTotalWorksheetBytes.HasValue
                        && uncompressed > limits.MaxTotalWorksheetBytes.Value - totalWorksheetBytes)
                        Resource($"XLSX worksheet XML 总大小超过限制: {limits.MaxTotalWorksheetBytes.Value}", provider);
                    totalWorksheetBytes += uncompressed;
                    if (enforceCellLimits && (limits.MaxColumnsPerSheet.HasValue || limits.MaxCells.HasValue))
                    {
                        var structure = ScanWorksheetStructure(entry, cancellationToken, provider,
                            limits.MaxColumnsPerSheet, limits.MaxCells, totalPhysicalCells);
                        totalPhysicalCells = structure.TotalCells;
                        maximumColumnsPerSheet = Math.Max(maximumColumnsPerSheet, structure.MaximumColumn);
                    }
                }
                else if (includePictures && IsPicture(entry.FullName))
                {
                    pictureCount++;
                    if (limits.MaxPictures.HasValue && pictureCount > limits.MaxPictures.Value)
                        Resource($"XLSX 图片数量超过限制: {limits.MaxPictures.Value}", provider);
                    if (limits.MaxPictureBytes.HasValue
                        && uncompressed > limits.MaxPictureBytes.Value)
                        Resource($"XLSX 图片部件超过限制: {entry.FullName}", provider);
                    if (limits.MaxTotalPictureBytes.HasValue
                        && uncompressed > limits.MaxTotalPictureBytes.Value - totalPictureBytes)
                        Resource($"XLSX 图片总大小超过限制: {limits.MaxTotalPictureBytes.Value}", provider);
                    totalPictureBytes += uncompressed;
                }
                if (entry.Length > 0)
                    ValidateXmlSafety(entry, cancellationToken, limits.MaxXmlCharacters,
                        limits.MaxXmlDepth, provider);
            }
        }
        catch (BingOfficesResourceLimitException)
        {
            throw;
        }
        catch (InvalidDataException exception)
        {
            throw new BingOfficesImportException("XLSX ZIP 预检失败。", exception, provider,
                BingOfficesStage.Preflight);
        }
        catch (XmlException exception)
        {
            throw new BingOfficesImportException("XLSX XML 预检失败。", exception, provider,
                BingOfficesStage.Preflight);
        }
        finally
        {
            // 解析器契约要求预检后的可定位 XLSX 流从文件头开始。
            source.Position = 0;
        }
    }

    /// <summary>
    /// 检查流头是否符合受支持的 ZIP 文件签名。
    /// </summary>
    /// <param name="source">待读取文件头的可定位流。</param>
    /// <returns>文件头匹配本地、空 ZIP 或跨卷 ZIP 签名时返回 true，否则返回 false。</returns>
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
    /// 判断 ZIP 条目是否为工作表 XML 部件。
    /// </summary>
    /// <param name="name">ZIP 条目名称。</param>
    /// <returns>名称位于 xl/worksheets/ 下且以 .xml 结尾时返回 true，否则返回 false。</returns>
    private static bool IsWorksheet(string name) => name.StartsWith("xl/worksheets/",
        StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 判断 ZIP 条目是否为工作簿图片部件。
    /// </summary>
    /// <param name="name">待判断的 ZIP 条目名称。</param>
    /// <returns>名称位于 xl/media/ 下时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool IsPicture(string name) => name.StartsWith("xl/media/",
        StringComparison.OrdinalIgnoreCase) && !name.EndsWith("/", StringComparison.Ordinal);

    /// <summary>
    /// 统计工作簿 XML 中声明的逻辑工作表数量。
    /// </summary>
    /// <param name="entry">工作簿 XML 的 ZIP 条目。</param>
    /// <param name="cancellationToken">用于取消 XML 扫描的令牌。</param>
    /// <param name="provider">用于异常上下文的提供程序名称。</param>
    /// <returns>工作簿中声明的工作表数量。</returns>
    private static int CountWorkbookSheets(ZipArchiveEntry entry, CancellationToken cancellationToken,
        string provider)
    {
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        var count = 0;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element
                && string.Equals(reader.LocalName, "sheet", StringComparison.Ordinal))
                count++;
        }
        return count;
    }

    /// <summary>
    /// 扫描工作表 XML 中的物理单元格和最大使用列。
    /// </summary>
    /// <param name="entry">工作表 XML 的 ZIP 条目。</param>
    /// <param name="cancellationToken">用于取消 XML 扫描的令牌。</param>
    /// <param name="provider">用于资源限制异常上下文的提供程序名称。</param>
    /// <param name="maximumColumns">允许的最大使用列；为 null 时不限制列数。</param>
    /// <param name="maximumCells">允许的累计物理单元格数；为 null 时不限制单元格数。</param>
    /// <param name="existingCells">扫描当前工作表前已统计的单元格数。</param>
    /// <returns>当前扫描后的累计单元格数和最大列号。</returns>
    private static WorksheetStructure ScanWorksheetStructure(ZipArchiveEntry entry,
        CancellationToken cancellationToken, string provider, int? maximumColumns,
        long? maximumCells, long existingCells)
    {
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        long totalCells = existingCells;
        var maximumColumn = 0;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "c", StringComparison.Ordinal))
                continue;
            totalCells++;
            if (maximumCells.HasValue && totalCells > maximumCells.Value)
                Resource($"XLSX 物理 Cell 数量超过限制: {maximumCells.Value}", provider);
            var reference = reader.GetAttribute("r");
            var column = ParseColumnIndex(reference);
            if (column > maximumColumn)
                maximumColumn = column;
            if (maximumColumns.HasValue && column > maximumColumns.Value)
                Resource($"XLSX Sheet 最大使用列超过限制: {maximumColumns.Value}", provider);
        }
        return new WorksheetStructure(totalCells, maximumColumn);
    }

    /// <summary>
    /// 从 A1 单元格引用中解析一基列号。
    /// </summary>
    /// <param name="reference">包含列字母的单元格引用。</param>
    /// <returns>解析出的列号；引用为空或不含列字母时返回 0。</returns>
    private static int ParseColumnIndex(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return 0;
        var column = 0;
        for (var index = 0; index < reference.Length; index++)
        {
            var character = reference[index];
            if (character < 'A' || character > 'Z')
                break;
            column = checked(column * 26 + character - 'A' + 1);
        }
        return column;
    }

    /// <summary>
    /// 创建启用安全 XML 设置的读取器。
    /// </summary>
    /// <param name="stream">待读取的 XML 流。</param>
    /// <returns>使用受限解析设置创建的 XML 读取器。</returns>
    private static XmlReader CreateReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0,
        MaxCharactersInDocument = 0
    });

    /// <summary>
    /// 保存工作表结构扫描结果。
    /// </summary>
    private readonly struct WorksheetStructure
    {
        /// <summary>
        /// 初始化一个 <see cref="WorksheetStructure" /> 类型的实例。
        /// </summary>
        /// <param name="totalCells">扫描后的累计物理单元格数。</param>
        /// <param name="maximumColumn">扫描到的最大列号。</param>
        internal WorksheetStructure(long totalCells, int maximumColumn)
        {
            TotalCells = totalCells;
            MaximumColumn = maximumColumn;
        }

        /// <summary>
        /// 获取累计物理单元格数。
        /// </summary>
        internal long TotalCells { get; }

        /// <summary>
        /// 获取扫描到的最大列号。
        /// </summary>
        internal int MaximumColumn { get; }
    }

    /// <summary>
    /// 校验 ZIP 条目路径不包含不受支持的路径形式。
    /// </summary>
    /// <param name="name">待校验的 ZIP 条目名称。</param>
    /// <param name="provider">用于资源限制异常上下文的提供程序名称。</param>
    private static void ValidateEntryPath(string name, string provider)
    {
        if (string.IsNullOrWhiteSpace(name) || name.StartsWith("/", StringComparison.Ordinal)
            || name.IndexOf('\\') >= 0 || name.IndexOf("../", StringComparison.Ordinal) >= 0
            || name.IndexOf("..\\", StringComparison.Ordinal) >= 0)
            Resource($"XLSX ZIP entry 路径无效: {name}", provider);
    }

    /// <summary>
    /// 将实际大小与可选的资源上限进行比较。
    /// </summary>
    /// <param name="actual">部件的实际大小，单位为字节。</param>
    /// <param name="maximum">允许的最大大小；为 null 时不检查。</param>
    /// <param name="name">用于异常消息的部件名称。</param>
    /// <param name="provider">用于资源限制异常上下文的提供程序名称。</param>
    private static void ValidateSize(long actual, long? maximum, string name, string provider)
    {
        if (maximum.HasValue && actual > maximum.Value)
            Resource($"XLSX XML 部件超过限制: {name}", provider);
    }

    /// <summary>
    /// 按给定上限检查 XML 条目的字符数量和嵌套深度。
    /// </summary>
    /// <param name="entry">待检查的 ZIP 条目。</param>
    /// <param name="cancellationToken">用于取消 XML 检查的令牌。</param>
    /// <param name="maxCharacters">允许的累计节点字符数量；为 null 时不限制字符数量。</param>
    /// <param name="maxDepth">允许的最大 XML 嵌套深度；为 null 时不限制嵌套深度。</param>
    /// <param name="provider">用于资源限制异常上下文的提供程序名称。</param>
    private static void ValidateXmlSafety(ZipArchiveEntry entry, CancellationToken cancellationToken,
        long? maxCharacters, int? maxDepth, string provider)
    {
        if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            return;
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
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
                Resource($"XLSX XML 嵌套深度超过限制: {entry.FullName}", provider);
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
                Resource($"XLSX XML 字符数量超过限制: {entry.FullName}", provider);
            characters += nodeCharacters;
        }
    }

    /// <summary>
    /// 抛出表示资源限制被超过的异常。
    /// </summary>
    /// <param name="message">资源限制异常消息。</param>
    /// <param name="provider">用于异常上下文的提供程序名称。</param>
    private static void Resource(string message, string provider) =>
        throw new BingOfficesResourceLimitException(message, provider: provider,
            operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
}
