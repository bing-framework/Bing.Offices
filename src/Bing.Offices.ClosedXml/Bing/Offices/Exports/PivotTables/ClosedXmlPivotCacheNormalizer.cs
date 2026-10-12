using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Bing.Offices.Internals;

namespace Bing.Offices.ClosedXml.Exports.PivotTables;

/// <summary>
/// 规范化 ClosedXML 生成的透视缓存字符串记录。
/// </summary>
/// <remarks>
/// ClosedXML 可能同时生成字符串共享项和直接字符串记录。按照 SpreadsheetML
/// 约定，存在共享项的字符串记录必须使用对应的索引，否则 Excel 在文本度量
/// 与多轴组合下可能无法计算 Count。
/// </remarks>
internal static class ClosedXmlPivotCacheNormalizer
{
    /// <summary>
    /// SpreadsheetML 主命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 将透视缓存中的共享字符串记录转换为索引记录并写入目标流。
    /// </summary>
    /// <param name="source">ClosedXML 已保存的 XLSX 流。</param>
    /// <param name="destination">规范化后的目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    internal static void Normalize(Stream source, Stream destination,
        CancellationToken cancellationToken)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (destination == null)
            throw new ArgumentNullException(nameof(destination));
        if (!source.CanRead)
            throw new ArgumentException("透视缓存源流不可读取。", nameof(source));
        if (!destination.CanWrite)
            throw new ArgumentException("透视缓存目标流不可写入。", nameof(destination));

        cancellationToken.ThrowIfCancellationRequested();
        var originalDestinationPosition = destination.CanSeek ? destination.Position : 0L;
        using var input = new MemoryStream();
        if (source.CanSeek)
            source.Position = 0;
        Copy(source, input, cancellationToken);
        input.Position = 0;

        using var normalized = new MemoryStream();
        using (var inputArchive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outputArchive = new ZipArchive(normalized, ZipArchiveMode.Create, leaveOpen: true))
        {
            var cacheParts = ReadCacheDefinitions(inputArchive, cancellationToken);
            foreach (var entry in inputArchive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var outputEntry = outputArchive.CreateEntry(entry.FullName, CompressionLevel.Fastest);
                using var output = outputEntry.Open();
                using var entryStream = entry.Open();
                var partPath = entry.FullName.Replace('\\', '/');
                if (cacheParts.DefinitionRecords.TryGetValue(partPath, out var recordsEntry))
                {
                    var document = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
                    using var recordsStream = recordsEntry.Open();
                    var records = XDocument.Load(recordsStream, LoadOptions.PreserveWhitespace);
                    NormalizeSharedItemFlags(document, records, cancellationToken);
                    Save(document, output);
                    continue;
                }

                if (cacheParts.RecordFields.TryGetValue(partPath, out var fields))
                {
                    var document = XDocument.Load(entryStream, LoadOptions.PreserveWhitespace);
                    NormalizeRecords(document, fields, cancellationToken);
                    Save(document, output);
                    continue;
                }

                Copy(entryStream, output, cancellationToken);
            }
        }

        normalized.Position = 0;
        if (destination.CanSeek)
            destination.Position = originalDestinationPosition;
        Copy(normalized, destination, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// 建立缓存记录部件到共享字符串索引的映射。
    /// </summary>
    /// <param name="archive">输入 XLSX 包。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>缓存定义部件与记录部件的规范化映射。</returns>
    private static (IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, int>>> RecordFields,
        IReadOnlyDictionary<string, ZipArchiveEntry> DefinitionRecords) ReadCacheDefinitions(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var entries = archive.Entries.ToDictionary(entry => entry.FullName.Replace('\\', '/'),
            StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, int>>>(
            StringComparer.OrdinalIgnoreCase);
        var definitionRecords = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in entries.Values.Where(IsCacheDefinition))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relationPath = RelationshipPart(definition.FullName);
            if (!entries.TryGetValue(relationPath, out var relationEntry))
                continue;
            using var relationStream = relationEntry.Open();
            var relation = XDocument.Load(relationStream, LoadOptions.PreserveWhitespace);
            foreach (var relationship in relation.Root?.Elements()
                         .Where(item => item.Name.LocalName == "Relationship") ?? Enumerable.Empty<XElement>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var type = (string)relationship.Attribute("Type");
                if (type == null || !type.EndsWith("/pivotCacheRecords", StringComparison.Ordinal))
                    continue;
                var target = ExcelPivotTemplateParts.ResolveTarget(relationPath,
                    (string)relationship.Attribute("Target"));
                if (!entries.ContainsKey(target))
                    continue;
                using var definitionStream = definition.Open();
                var document = XDocument.Load(definitionStream, LoadOptions.PreserveWhitespace);
                result[target] = ReadFields(document);
                definitionRecords[definition.FullName.Replace('\\', '/')] = entries[target];
            }
        }
        return (result, definitionRecords);
    }

    /// <summary>
    /// 根据共享项内容显式写入缓存字段类型标志。
    /// </summary>
    /// <param name="document">缓存定义文档。</param>
    /// <param name="records">缓存记录文档。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void NormalizeSharedItemFlags(XDocument document, XDocument records,
        CancellationToken cancellationToken)
    {
        var fields = document.Root?.Element(SpreadsheetNamespace + "cacheFields")?.Elements(
            SpreadsheetNamespace + "cacheField").ToArray() ?? Array.Empty<XElement>();
        var containsStrings = new bool[fields.Length];
        var containsNumbers = new bool[fields.Length];
        var containsDates = new bool[fields.Length];
        var containsBooleans = new bool[fields.Length];
        var containsBlanks = new bool[fields.Length];
        var containsNonDates = new bool[fields.Length];
        var containsLongText = new bool[fields.Length];
        var allNumbersAreIntegers = Enumerable.Repeat(true, fields.Length).ToArray();

        for (var index = 0; index < fields.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var sharedItems = fields[index].Element(SpreadsheetNamespace + "sharedItems");
            if (sharedItems == null)
            {
                sharedItems = new XElement(SpreadsheetNamespace + "sharedItems");
                fields[index].Add(sharedItems);
            }
            var items = sharedItems.Elements().ToArray();
            var numericItems = items.Where(item => item.Name == SpreadsheetNamespace + "n").ToArray();
            containsStrings[index] = items.Any(item => item.Name == SpreadsheetNamespace + "s");
            containsNumbers[index] = numericItems.Length > 0;
            containsDates[index] = items.Any(item => item.Name == SpreadsheetNamespace + "d");
            containsBooleans[index] = items.Any(item => item.Name == SpreadsheetNamespace + "b");
            containsBlanks[index] = items.Any(item => item.Name == SpreadsheetNamespace + "m");
            containsNonDates[index] = containsStrings[index] || containsNumbers[index] || containsBooleans[index];
            containsLongText[index] = items.Where(item => item.Name == SpreadsheetNamespace + "s")
                .Any(item => ((string)item.Attribute("v") ?? item.Value).Length > 255);
            allNumbersAreIntegers[index] = numericItems.All(IsIntegerSharedItem);
        }

        foreach (var record in records.Root?.Elements(SpreadsheetNamespace + "r") ?? Enumerable.Empty<XElement>())
        {
            cancellationToken.ThrowIfCancellationRequested();
            var index = 0;
            foreach (var item in record.Elements())
            {
                if (index >= fields.Length)
                    break;

                if (item.Name == SpreadsheetNamespace + "s")
                {
                    containsStrings[index] = true;
                    containsNonDates[index] = true;
                    containsLongText[index] |= ((string)item.Attribute("v") ?? item.Value).Length > 255;
                }
                else if (item.Name == SpreadsheetNamespace + "n")
                {
                    containsNumbers[index] = true;
                    containsNonDates[index] = true;
                    allNumbersAreIntegers[index] &= IsIntegerSharedItem(item);
                }
                else if (item.Name == SpreadsheetNamespace + "d")
                    containsDates[index] = true;
                else if (item.Name == SpreadsheetNamespace + "b")
                {
                    containsBooleans[index] = true;
                    containsNonDates[index] = true;
                }
                else if (item.Name == SpreadsheetNamespace + "m")
                    containsBlanks[index] = true;
                index++;
            }
        }

        for (var index = 0; index < fields.Length; index++)
        {
            var sharedItems = fields[index].Element(SpreadsheetNamespace + "sharedItems");
            var containsMixedTypes = new[]
            {
                containsStrings[index], containsNumbers[index], containsBooleans[index], containsDates[index]
            }.Count(value => value) > 1;
            SetFlag(sharedItems, "containsString", containsStrings[index]);
            SetFlag(sharedItems, "containsSemiMixedTypes", containsStrings[index]);
            SetFlag(sharedItems, "containsNonDate", containsNonDates[index]);
            SetFlag(sharedItems, "containsDate", containsDates[index]);
            SetFlag(sharedItems, "containsBlank", containsBlanks[index]);
            SetFlag(sharedItems, "containsMixedTypes", containsMixedTypes);
            SetFlag(sharedItems, "containsNumber", containsNumbers[index]);
            SetFlag(sharedItems, "containsInteger", containsNumbers[index] && !containsMixedTypes
                && allNumbersAreIntegers[index]);
            SetFlag(sharedItems, "longText", containsLongText[index]);
        }
    }

    /// <summary>
    /// 判断共享数值项是否为整数。
    /// </summary>
    /// <param name="item">数值共享项。</param>
    /// <returns>数值为有限整数时返回 true。</returns>
    private static bool IsIntegerSharedItem(XElement item)
    {
        var value = (string)item.Attribute("v") ?? item.Value;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
            && double.IsFinite(number) && Math.Truncate(number) == number;
    }

    /// <summary>
    /// 设置共享项布尔标志。
    /// </summary>
    /// <param name="sharedItems">共享项定义。</param>
    /// <param name="name">标志名称。</param>
    /// <param name="value">标志值。</param>
    private static void SetFlag(XElement sharedItems, string name, bool value)
    {
        sharedItems.SetAttributeValue(name, value ? "1" : "0");
    }

    /// <summary>
    /// 读取缓存字段中的字符串共享项。
    /// </summary>
    /// <param name="document">缓存定义文档。</param>
    /// <returns>按字段顺序排列的字符串到索引映射。</returns>
    private static IReadOnlyList<IReadOnlyDictionary<string, int>> ReadFields(XDocument document)
    {
        var fields = document.Root?.Element(SpreadsheetNamespace + "cacheFields")?.Elements(
            SpreadsheetNamespace + "cacheField") ?? Enumerable.Empty<XElement>();
        return fields.Select(field =>
        {
            var items = field.Element(SpreadsheetNamespace + "sharedItems")?
                .Elements() ?? Enumerable.Empty<XElement>();
            var indexes = new Dictionary<string, int>(StringComparer.Ordinal);
            var index = 0;
            foreach (var item in items)
            {
                if (item.Name == SpreadsheetNamespace + "s")
                {
                    var value = (string)item.Attribute("v") ?? item.Value;
                    if (!indexes.ContainsKey(value))
                        indexes.Add(value, index);
                }
                index++;
            }
            return indexes;
        }).ToArray();
    }

    /// <summary>
    /// 将存在共享项的直接字符串记录替换为索引记录。
    /// </summary>
    /// <param name="document">缓存记录文档。</param>
    /// <param name="fields">按字段顺序排列的共享字符串映射。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void NormalizeRecords(XDocument document,
        IReadOnlyList<IReadOnlyDictionary<string, int>> fields,
        CancellationToken cancellationToken)
    {
        foreach (var record in document.Root?.Elements(SpreadsheetNamespace + "r")
                     ?? Enumerable.Empty<XElement>())
        {
            var cells = record.Elements().ToArray();
            for (var index = 0; index < cells.Length && index < fields.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (cells[index].Name != SpreadsheetNamespace + "s"
                    || fields[index].Count == 0)
                    continue;
                var value = (string)cells[index].Attribute("v") ?? cells[index].Value;
                if (!fields[index].TryGetValue(value, out var sharedIndex))
                    continue;
                cells[index].ReplaceWith(new XElement(SpreadsheetNamespace + "x",
                    new XAttribute("v", sharedIndex)));
            }
        }
    }

    /// <summary>
    /// 判断部件是否为透视缓存定义。
    /// </summary>
    /// <param name="entry">ZIP 部件。</param>
    /// <returns>是透视缓存定义时返回 true。</returns>
    private static bool IsCacheDefinition(ZipArchiveEntry entry) =>
        entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        && entry.FullName.IndexOf("pivotCacheDefinition", StringComparison.OrdinalIgnoreCase) >= 0;

    /// <summary>
    /// 获取部件对应的关系部件路径。
    /// </summary>
    /// <param name="part">部件路径。</param>
    /// <returns>关系部件路径。</returns>
    private static string RelationshipPart(string part)
    {
        var normalized = part.Replace('\\', '/');
        var separator = normalized.LastIndexOf('/');
        return normalized[..(separator + 1)] + "_rels/" + normalized[(separator + 1)..] + ".rels";
    }

    /// <summary>
    /// 将 XML 文档写入 ZIP 部件。
    /// </summary>
    /// <param name="document">待保存的 XML 文档。</param>
    /// <param name="destination">目标部件流。</param>
    private static void Save(XDocument document, Stream destination)
    {
        using var writer = System.Xml.XmlWriter.Create(destination, new System.Xml.XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false,
            Indent = false
        });
        document.Save(writer);
    }

    /// <summary>
    /// 复制 ZIP 部件内容并响应取消。
    /// </summary>
    /// <param name="source">源流。</param>
    /// <param name="destination">目标流。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static void Copy(Stream source, Stream destination, CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
        {
            cancellationToken.ThrowIfCancellationRequested();
            destination.Write(buffer, 0, read);
        }
    }
}
