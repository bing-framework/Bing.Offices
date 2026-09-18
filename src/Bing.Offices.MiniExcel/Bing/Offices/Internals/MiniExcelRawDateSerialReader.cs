using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// 从 MiniExcel 读取前的 XLSX 工作表 XML 中保留原始日期 serial。
/// </summary>
/// <remarks>
/// 读取结果按一基物理行列坐标索引，仅用于恢复日期单元格的原始数值。
/// </remarks>
internal static class MiniExcelRawDateSerialReader
{
    /// <summary>
    /// XLSX 工作表主 XML 命名空间。
    /// </summary>
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    /// <summary>
    /// XLSX 工作簿关系引用使用的 XML 命名空间。
    /// </summary>
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    /// <summary>
    /// XLSX 包关系文件使用的 XML 命名空间。
    /// </summary>
    private const string PackageRelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>
    /// 读取指定工作表的 numeric cell serial，键为一基物理行列号。
    /// </summary>
    /// <param name="source">已通过 XLSX 预检且可定位的工作簿流。</param>
    /// <param name="sheetName">工作表物理名称。</param>
    /// <param name="cancellationToken">读取过程使用的取消令牌。</param>
    /// <returns>原始 numeric serial 索引。</returns>
    internal static IReadOnlyDictionary<long, double> Read(Stream source, string sheetName,
        CancellationToken cancellationToken = default)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new ArgumentException("工作表名称不能为空。", nameof(sheetName));
        if (!source.CanSeek)
            throw new ArgumentException("读取原始日期 serial 需要可定位流。", nameof(source));

        var originalPosition = source.Position;
        source.Position = 0;
        try
        {
            using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
            var worksheetPath = ResolveWorksheetPath(archive, sheetName, cancellationToken);
            var worksheet = archive.GetEntry(worksheetPath);
            if (worksheet == null)
                throw new InvalidDataException($"XLSX 缺少工作表 XML: {worksheetPath}");

            var result = new Dictionary<long, double>();
            using var worksheetStream = worksheet.Open();
            using var reader = XmlReader.Create(worksheetStream, CreateReaderSettings());
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c")
                    continue;

                var reference = reader.GetAttribute("r");
                if (!TryParseReference(reference, out var row, out var column))
                    continue;

                var cellType = reader.GetAttribute("t");
                if (cellType != null && !string.Equals(cellType, "n", StringComparison.Ordinal))
                    continue;

                var serialText = ReadValueElement(reader, cancellationToken);
                if (double.TryParse(serialText, NumberStyles.Float, CultureInfo.InvariantCulture,
                        out var serial) && !double.IsNaN(serial) && !double.IsInfinity(serial))
                    result[CreateKey(row, column)] = serial;
            }
            return result;
        }
        finally
        {
            source.Position = originalPosition;
        }
    }

    /// <summary>
    /// 读取当前单元格的值元素文本。
    /// </summary>
    /// <param name="reader">定位在单元格元素上的 XML 读取器。</param>
    /// <param name="cancellationToken">用于取消读取的令牌。</param>
    /// <returns>值元素文本；不存在时返回 null。</returns>
    private static string ReadValueElement(XmlReader reader, CancellationToken cancellationToken)
    {
        using var cellReader = reader.ReadSubtree();
        while (cellReader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (cellReader.NodeType == XmlNodeType.Element && cellReader.LocalName == "v")
                return cellReader.ReadElementContentAsString();
        }
        return null;
    }

    /// <summary>
    /// 根据工作表名称解析对应的 ZIP 条目路径。
    /// </summary>
    /// <param name="archive">已打开的 XLSX ZIP 包。</param>
    /// <param name="sheetName">工作表物理名称。</param>
    /// <param name="cancellationToken">用于取消解析的令牌。</param>
    /// <returns>工作表 XML 的规范化 ZIP 路径。</returns>
    private static string ResolveWorksheetPath(ZipArchive archive, string sheetName,
        CancellationToken cancellationToken)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml");
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (workbookEntry == null || relationshipsEntry == null)
            throw new InvalidDataException("XLSX 缺少工作表关系描述。");

        XDocument workbook;
        using (var stream = workbookEntry.Open())
        using (var reader = XmlReader.Create(stream, CreateReaderSettings()))
            workbook = XDocument.Load(reader, LoadOptions.None);
        cancellationToken.ThrowIfCancellationRequested();

        var relationshipId = workbook.Root?
            .Element(XName.Get("sheets", MainNamespace))?
            .Elements(XName.Get("sheet", MainNamespace))
            .FirstOrDefault(sheet => string.Equals((string)sheet.Attribute("name"), sheetName,
                StringComparison.Ordinal))?
            .Attribute(XName.Get("id", RelationshipsNamespace))?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
            throw new InvalidDataException($"XLSX 缺少工作表关系: {sheetName}");

        XDocument relationships;
        using (var stream = relationshipsEntry.Open())
        using (var reader = XmlReader.Create(stream, CreateReaderSettings()))
            relationships = XDocument.Load(reader, LoadOptions.None);
        cancellationToken.ThrowIfCancellationRequested();

        var target = relationships.Root?
            .Elements(XName.Get("Relationship", PackageRelationshipsNamespace))
            .FirstOrDefault(item => string.Equals((string)item.Attribute("Id"), relationshipId,
                StringComparison.Ordinal))?
            .Attribute("Target")?.Value;
        if (string.IsNullOrWhiteSpace(target))
            throw new InvalidDataException($"XLSX 缺少工作表关系目标: {sheetName}");
        return NormalizeTarget(target);
    }

    /// <summary>
    /// 规范化关系目标并限制其位于 xl 目录下。
    /// </summary>
    /// <param name="target">关系文件中的目标路径。</param>
    /// <returns>规范化后的 ZIP 条目路径。</returns>
    private static string NormalizeTarget(string target)
    {
        var segments = new List<string>();
        foreach (var segment in target.Replace('\\', '/').Split('/'))
        {
            if (string.IsNullOrEmpty(segment) || segment == ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count == 0)
                    throw new InvalidDataException($"XLSX 工作表关系目标越界: {target}");
                segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        var normalized = string.Join("/", segments);
        if (!normalized.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            normalized = "xl/" + normalized.TrimStart('/');
        return normalized;
    }

    /// <summary>
    /// 将 Excel 单元格引用解析为一基行列号。
    /// </summary>
    /// <param name="reference">Excel A1 单元格引用。</param>
    /// <param name="row">解析得到的一基行号。</param>
    /// <param name="column">解析得到的一基列号。</param>
    /// <returns>引用格式有效且行列均为正数时返回 true，否则返回 false。</returns>
    private static bool TryParseReference(string reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
            return false;
        var index = 0;
        long parsedColumn = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            parsedColumn = parsedColumn * 26 + char.ToUpperInvariant(reference[index]) - 'A' + 1;
            if (parsedColumn > int.MaxValue)
                return false;
            index++;
        }
        if (index == 0 || index >= reference.Length || !int.TryParse(reference.Substring(index),
                NumberStyles.None, CultureInfo.InvariantCulture, out row)
            || row <= 0 || parsedColumn <= 0)
            return false;
        column = (int)parsedColumn;
        return true;
    }

    /// <summary>
    /// 将一基物理行列号编码为日期 serial 索引键。
    /// </summary>
    /// <param name="row">一基物理行号。</param>
    /// <param name="column">一基物理列号。</param>
    /// <returns>由行列坐标组成的索引键。</returns>
    internal static long CreateKey(int row, int column) => ((long)row << 32) | (uint)column;

    /// <summary>
    /// 创建禁止 DTD 和外部实体的 XML 读取设置。
    /// </summary>
    /// <returns>用于读取 XLSX XML 的安全设置。</returns>
    private static XmlReaderSettings CreateReaderSettings() => new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0
    };
}
