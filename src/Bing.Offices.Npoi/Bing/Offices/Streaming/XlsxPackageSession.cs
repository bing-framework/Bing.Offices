using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Bing.Offices.Imports;

namespace Bing.Offices.Streaming;

/// <summary>有界 XLSX 元数据和 XML 会话；不创建完整工作簿 DOM。</summary>
internal sealed class XlsxPackageSession : IDisposable
{
    internal const string MainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private readonly ZipArchive _archive;
    private readonly ExcelStreamingImportOptions _options;
    private readonly CancellationToken _token;
    private long _actualBytes;
    internal XlsxSharedStringStore Strings { get; }
    internal List<(string Name, string Path)> Sheets { get; } = new();
    internal List<int> StyleFormats { get; } = new();
    internal Dictionary<int, string> NumberFormats { get; } = new();
    internal bool Date1904 { get; private set; }
    internal long ActualBytesRead => _actualBytes;

    internal XlsxPackageSession(Stream source, ExcelStreamingImportOptions options,
        ExcelStreamingTempSession temporary, CancellationToken token)
    {
        _options = options; _token = token;
        _archive = new ZipArchive(source, ZipArchiveMode.Read, true);
        try
        {
            NpoiXlsxZipPreflight.ValidateMetadata(_archive, options.ResourceLimits, token, false);
            Strings = new XlsxSharedStringStore(temporary, options.SharedStringMemoryBytes, options.MaxCellCharacters);
        }
        catch { _archive.Dispose(); throw; }
    }

    internal XmlReader OpenXml(string path, bool asynchronous, long? partLimit = null)
    {
        var entry = _archive.GetEntry(path) ?? throw new InvalidDataException("XLSX 缺少部件: " + path);
        var limit = Math.Min(_options.ResourceLimits.MaxZipEntryUncompressedBytes.Value, partLimit ?? long.MaxValue);
        if (entry.Length > limit) throw new StreamingResourceLimitException("XLSX 部件超过额度: " + path);
        var stream = new StreamingReadLimit(entry.Open(), Math.Min(limit, entry.Length), _token, count =>
        {
            if (count > _options.ResourceLimits.MaxZipTotalUncompressedBytes.Value - _actualBytes)
                throw new StreamingResourceLimitException("XLSX 实际总解压字节超过限制。");
            _actualBytes += count;
        });
        try
        {
            return XmlReader.Create(stream, new XmlReaderSettings
            {
                Async = asynchronous, CloseInput = true, XmlResolver = null,
                DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true,
                MaxCharactersInDocument = _options.ResourceLimits.MaxXmlCharacters.Value
            });
        }
        catch { stream.Dispose(); throw; }
    }

    private ValueTask<bool> Read(XmlReader reader, bool asynchronous)
        => StreamingXml.Read(reader, asynchronous, _options.ResourceLimits.MaxXmlDepth.Value, _token);

    internal async ValueTask<bool> Initialize(bool asynchronous)
    {
        var rootRelationships = await Relationships("", "_rels/.rels", asynchronous).ConfigureAwait(false);
        var main = rootRelationships.Values.SingleOrDefault(item => item.Type == RelationshipNamespace + "/officeDocument");
        if (main.Path == null) throw new NotSupportedException("仅支持普通 Transitional XLSX。");
        var types = new Dictionary<string, string>(StringComparer.Ordinal);
        using (var reader = OpenXml("[Content_Types].xml", asynchronous))
            while (await Read(reader, asynchronous).ConfigureAwait(false))
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Override")
                {
                    if (types.Count >= _options.ResourceLimits.MaxZipEntries.Value)
                        throw new StreamingResourceLimitException("Content Types 数量超过限制。");
                    types.Add(Resolve("", reader.GetAttribute("PartName")), reader.GetAttribute("ContentType"));
                }
        if (!types.TryGetValue(main.Path, out var contentType)
            || contentType != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
            throw new NotSupportedException("不支持加密、宏或非 XLSX 工作簿。");
        var slash = main.Path.LastIndexOf('/');
        var directory = slash < 0 ? "" : main.Path.Substring(0, slash + 1);
        var fileName = main.Path.Substring(slash + 1);
        var relationships = await Relationships(main.Path, directory + "_rels/" + fileName + ".rels", asynchronous).ConfigureAwait(false);
        using (var reader = OpenXml(main.Path, asynchronous))
        {
            while (await Read(reader, asynchronous).ConfigureAwait(false))
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth == 0 && (reader.LocalName != "workbook" || reader.NamespaceURI != MainNamespace))
                    throw new NotSupportedException("不支持此工作簿 XML 命名空间。");
                if (reader.LocalName == "workbookPr") Date1904 = reader.GetAttribute("date1904") is "1" or "true";
                if (reader.LocalName != "sheet" || reader.NamespaceURI != MainNamespace) continue;
                var name = reader.GetAttribute("name");
                var id = reader.GetAttribute("id", RelationshipNamespace);
                if (string.IsNullOrEmpty(name) || name.Length > 31 || id == null
                    || !relationships.TryGetValue(id, out var relation)
                    || relation.Type != RelationshipNamespace + "/worksheet")
                    throw new InvalidDataException("工作表关系无效。");
                if (Sheets.Count >= _options.MaxSheets)
                    throw new StreamingResourceLimitException("工作表数量超过限制。");
                if (Sheets.Any(sheet => string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidDataException("工作表名称重复。");
                Sheets.Add((name, relation.Path));
            }
        }
        long sheetBytes = 0;
        foreach (var sheet in Sheets)
        {
            var size = (_archive.GetEntry(sheet.Path) ?? throw new InvalidDataException("工作表部件缺失。")).Length;
            if (size > (_options.ResourceLimits.MaxWorksheetBytes ?? long.MaxValue)
                || size > (_options.ResourceLimits.MaxTotalWorksheetBytes ?? long.MaxValue) - sheetBytes)
                throw new StreamingResourceLimitException("工作表 XML 字节数超过限制。");
            sheetBytes += size;
        }
        var styles = relationships.Values.SingleOrDefault(item => item.Type == RelationshipNamespace + "/styles");
        if (styles.Path != null) await ReadStyles(styles.Path, asynchronous).ConfigureAwait(false);
        if (StyleFormats.Count == 0) StyleFormats.Add(0);
        var strings = relationships.Values.SingleOrDefault(item => item.Type == RelationshipNamespace + "/sharedStrings");
        if (strings.Path != null) await ReadStrings(strings.Path, asynchronous).ConfigureAwait(false);
        return true;
    }

    private async ValueTask<Dictionary<string, (string Type, string Path)>> Relationships(string owner, string path, bool asynchronous)
    {
        var result = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
        using var reader = OpenXml(path, asynchronous);
        while (await Read(reader, asynchronous).ConfigureAwait(false))
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "Relationship") continue;
            if (reader.GetAttribute("TargetMode") == "External") continue;
            if (result.Count >= _options.ResourceLimits.MaxZipEntries.Value)
                throw new StreamingResourceLimitException("关系数量超过限制。");
            var id = reader.GetAttribute("Id") ?? throw new InvalidDataException("关系缺少 Id。");
            if (!result.TryAdd(id, (reader.GetAttribute("Type"), Resolve(owner, reader.GetAttribute("Target")))))
                throw new InvalidDataException("关系 Id 重复。");
        }
        return result;
    }

    /// <summary>允许合法的包内相对父路径，但禁止逃逸、外部 URI 和反斜线。</summary>
    internal static string Resolve(string owner, string target)
    {
        if (string.IsNullOrWhiteSpace(target) || target.Contains('\\') || target.Contains(':') || target.Contains('#') || target.Contains('?'))
            throw new InvalidDataException("XLSX 关系路径无效。");
        target = Uri.UnescapeDataString(target);
        if (target.Contains('\\') || target.Contains(':')) throw new InvalidDataException("XLSX 关系路径无效。");
        var parts = new List<string>();
        var prefix = target.StartsWith('/') ? "" : owner.Substring(0, owner.LastIndexOf('/') + 1);
        foreach (var part in (prefix + target).Split('/'))
        {
            if (part is "" or ".") continue;
            if (part == "..")
            {
                if (parts.Count == 0) throw new InvalidDataException("XLSX 关系越出包根目录。");
                parts.RemoveAt(parts.Count - 1);
            }
            else parts.Add(part);
        }
        if (parts.Count == 0) throw new InvalidDataException("XLSX 关系路径为空。");
        return string.Join("/", parts);
    }

    private async ValueTask<bool> ReadStyles(string path, bool asynchronous)
    {
        using var reader = OpenXml(path, asynchronous, _options.ResourceLimits.MaxStylesBytes);
        var cellFormatsDepth = -1;
        var fonts = 0;
        while (await Read(reader, asynchronous).ConfigureAwait(false))
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "cellXfs" && !reader.IsEmptyElement) cellFormatsDepth = reader.Depth;
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cellXfs") cellFormatsDepth = -1;
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "font" && ++fonts > _options.MaxFonts)
                throw new StreamingResourceLimitException("字体数量超过限制。");
            if (reader.LocalName == "numFmt")
            {
                if (NumberFormats.Count >= _options.MaxStyles) throw new StreamingResourceLimitException("数字格式数量超过限制。");
                var code = reader.GetAttribute("formatCode") ?? "";
                if (code.Length > _options.MaxCellCharacters) throw new StreamingResourceLimitException("数字格式文本超过限制。");
                NumberFormats.Add(int.Parse(reader.GetAttribute("numFmtId"), CultureInfo.InvariantCulture), code);
            }
            if (reader.LocalName == "xf" && cellFormatsDepth >= 0 && reader.Depth == cellFormatsDepth + 1)
            {
                if (StyleFormats.Count >= _options.MaxStyles) throw new StreamingResourceLimitException("样式数量超过限制。");
                StyleFormats.Add(int.Parse(reader.GetAttribute("numFmtId") ?? "0", CultureInfo.InvariantCulture));
            }
        }
        return true;
    }

    private async ValueTask<bool> ReadStrings(string path, bool asynchronous)
    {
        using var reader = OpenXml(path, asynchronous, _options.ResourceLimits.MaxSharedStringsBytes);
        StringBuilder current = null;
        var textDepth = -1;
        var phoneticDepth = -1;
        while (await Read(reader, asynchronous).ConfigureAwait(false))
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "si")
                {
                    current = new StringBuilder();
                    if (reader.IsEmptyElement)
                    {
                        if (asynchronous) await Strings.AddAsync("", _token).ConfigureAwait(false); else Strings.Add("");
                        current = null;
                    }
                }
                if (reader.LocalName == "rPh" && !reader.IsEmptyElement) phoneticDepth = reader.Depth;
                if (reader.LocalName == "t" && !reader.IsEmptyElement && phoneticDepth < 0) textDepth = reader.Depth;
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == textDepth) textDepth = -1;
                if (reader.Depth == phoneticDepth) phoneticDepth = -1;
                if (reader.LocalName == "si" && current != null)
                {
                    var value = Decode(current.ToString());
                    if (asynchronous) await Strings.AddAsync(value, _token).ConfigureAwait(false); else Strings.Add(value);
                    current = null;
                }
            }
            else if (current != null && textDepth >= 0 && phoneticDepth < 0
                && reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace or XmlNodeType.Whitespace)
                current.Append(await StreamingXml.Text(reader, asynchronous, _options.MaxCellCharacters - current.Length, _token).ConfigureAwait(false));
        }
        return true;
    }

    internal static string Decode(string value) => Regex.Replace(value, "_x([0-9a-fA-F]{4})_",
        match => ((char)int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());

    public void Dispose()
    {
        try { Strings?.Dispose(); }
        finally { _archive.Dispose(); }
    }
}
