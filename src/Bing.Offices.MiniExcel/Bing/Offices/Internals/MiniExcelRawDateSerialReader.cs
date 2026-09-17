using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Bing.Offices.MiniExcel.Internals;

/// <summary>
/// 从 MiniExcel 读取前的 XLSX worksheet XML 保留原始 numeric serial。
/// </summary>
internal static class MiniExcelRawDateSerialReader
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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

    internal static long CreateKey(int row, int column) => ((long)row << 32) | (uint)column;

    private static XmlReaderSettings CreateReaderSettings() => new XmlReaderSettings
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersFromEntities = 0
    };
}
