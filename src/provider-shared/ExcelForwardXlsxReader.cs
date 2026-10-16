using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Validations;

namespace Bing.Offices;

/// <summary>
/// 读取受限 Forward 实体路径需要的 XLSX 工作簿部件。
/// </summary>
internal static class ExcelForwardXlsxReader
{
    /// <summary>
    /// 读取单 Sheet 工作簿的名称、部件路径和日期系统。
    /// </summary>
    internal static ExcelForwardXlsxWorkbook ReadWorkbook(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new BingOfficesImportException("XLSX 缺少 workbook.xml。", provider: "Excel",
                stage: BingOfficesStage.Preflight);
        var relationEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new BingOfficesImportException("XLSX 缺少工作簿关系部件。", provider: "Excel",
                stage: BingOfficesStage.Preflight);
        var sheets = new List<(string Name, string RelationId)>();
        var definedNames = new List<ExcelForwardXlsxDefinedName>();
        var date1904 = false;
        using (var stream = workbookEntry.Open())
        using (var reader = CreateReader(stream))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element)
                    continue;
                if (string.Equals(reader.LocalName, "workbookPr", StringComparison.Ordinal))
                {
                    var value = reader.GetAttribute("date1904");
                    date1904 = string.Equals(value, "1", StringComparison.Ordinal)
                        || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);
                }
                else if (string.Equals(reader.LocalName, "sheet", StringComparison.Ordinal))
                {
                    var relationId = reader.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                        ?? reader.GetAttribute("r:id");
                    sheets.Add((reader.GetAttribute("name"), relationId));
                }
                else if (string.Equals(reader.LocalName, "definedName", StringComparison.Ordinal))
                {
                    var name = reader.GetAttribute("name");
                    var localSheetId = -1;
                    var localSheetText = reader.GetAttribute("localSheetId");
                    if (!string.IsNullOrWhiteSpace(localSheetText)
                        && (!int.TryParse(localSheetText, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out localSheetId)
                            || localSheetId < 0))
                        localSheetId = -2;
                    var formula = ReadElementText(reader);
                    if (!string.IsNullOrWhiteSpace(name))
                        definedNames.Add(new ExcelForwardXlsxDefinedName(name, localSheetId,
                            formula));
                }
            }
        }

        var relationships = new Dictionary<string, string>(StringComparer.Ordinal);
        using (var stream = relationEntry.Open())
        using (var reader = CreateReader(stream))
        {
            while (reader.Read())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (reader.NodeType != XmlNodeType.Element
                    || !string.Equals(reader.LocalName, "Relationship", StringComparison.Ordinal))
                    continue;
                var id = reader.GetAttribute("Id");
                var target = reader.GetAttribute("Target");
                if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(target))
                    relationships[id] = NormalizePartPath(target);
            }
        }
        if (sheets.Count == 0)
            throw new BingOfficesImportException("XLSX 未声明工作表。", provider: "Excel",
                stage: BingOfficesStage.Preflight);
        var workbookSheets = new List<ExcelForwardXlsxSheet>(sheets.Count);
        foreach (var sheet in sheets)
        {
            if (string.IsNullOrWhiteSpace(sheet.Name) || string.IsNullOrWhiteSpace(sheet.RelationId)
                || !relationships.TryGetValue(sheet.RelationId, out var path))
                throw new BingOfficesImportException("XLSX 工作表关系无效。", provider: "Excel",
                    stage: BingOfficesStage.Preflight, sheetName: sheet.Name);
            workbookSheets.Add(new ExcelForwardXlsxSheet(sheet.Name, path));
        }
        return new ExcelForwardXlsxWorkbook(workbookSheets, date1904, definedNames);
    }

    /// <summary>
    /// 读取当前元素的文本内容并保持父级读取器在当前元素上。
    /// </summary>
    private static string ReadElementText(XmlReader reader)
    {
        var builder = new StringBuilder();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            if (subtree.NodeType == XmlNodeType.Text || subtree.NodeType == XmlNodeType.CDATA)
                builder.Append(subtree.Value);
        }
        return builder.ToString();
    }

    /// <summary>
    /// 解析前向工作簿中的单格命名锚点。
    /// </summary>
    /// <param name="workbook">前向工作簿描述。</param>
    /// <param name="sheetName">布局声明的工作表名称。</param>
    /// <param name="anchorName">命名锚点名称。</param>
    /// <returns>命名锚点解析出的零基单元格坐标。</returns>
    internal static ExcelEntityCellReference ResolveSingleCellAnchor(
        ExcelForwardXlsxWorkbook workbook, string sheetName, string anchorName)
    {
        if (workbook == null)
            throw new ArgumentNullException(nameof(workbook));
        var sheetIndex = workbook.GetSheetIndex(sheetName);
        if (sheetIndex < 0)
            throw new BingOfficesConfigurationException(
                $"Forward 实体布局要求 Sheet {sheetName}，工作簿中不存在该工作表。",
                stage: BingOfficesStage.Plan);
        var candidates = workbook.DefinedNames.Where(item =>
            string.Equals(item.Name, anchorName, StringComparison.OrdinalIgnoreCase)).ToArray();
        var scoped = candidates.Where(item => item.LocalSheetId == sheetIndex).ToArray();
        var global = candidates.Where(item => item.LocalSheetId == -1).ToArray();
        var matches = scoped.Length > 0 ? scoped : global;
        if (matches.Length == 0)
            throw new BingOfficesConfigurationException(
                $"模板缺少命名锚点: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);
        if (matches.Length > 1)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点不唯一: {sheetName}!{anchorName}", stage: BingOfficesStage.Plan);

        var match = SingleCellFormula.Match(matches[0].Formula ?? string.Empty);
        if (!match.Success)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个有界单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        var targetSheet = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value.Replace("''", "'")
            : match.Groups["plain"].Success ? match.Groups["plain"].Value : sheetName;
        if (!match.Groups["quoted"].Success && !match.Groups["plain"].Success
            && matches[0].LocalSheetId != sheetIndex)
            throw new BingOfficesConfigurationException(
                $"模板命名锚点缺少 Sheet 引用: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (!string.Equals(targetSheet, sheetName, StringComparison.OrdinalIgnoreCase))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点引用了其他 Sheet: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        if (match.Groups["endColumn"].Success
            && (!string.Equals(match.Groups["startColumn"].Value,
                    match.Groups["endColumn"].Value, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(match.Groups["startRow"].Value,
                    match.Groups["endRow"].Value, StringComparison.Ordinal)))
            throw new BingOfficesConfigurationException(
                $"模板命名锚点必须引用单个单元格: {sheetName}!{anchorName}",
                stage: BingOfficesStage.Plan);
        try
        {
            return ExcelEntityCellReference.Parse(match.Groups["startColumn"].Value
                + match.Groups["startRow"].Value);
        }
        catch (ArgumentException exception)
        {
            throw new BingOfficesConfigurationException(
                $"模板命名锚点地址超出工作表边界: {sheetName}!{anchorName}", exception,
                BingOfficesStage.Plan);
        }
    }

    private static readonly Regex SingleCellFormula = new Regex(
        @"^\s*=?\s*(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^'!]+))?!\$?(?<startColumn>[A-Za-z]{1,3})\$?(?<startRow>[1-9][0-9]*)(?::\$?(?<endColumn>[A-Za-z]{1,3})\$?(?<endRow>[1-9][0-9]*))?\s*$",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// 读取共享字符串表。
    /// </summary>
    internal static ExcelForwardXlsxSharedStrings ReadSharedStrings(ZipArchive archive,
        CancellationToken cancellationToken, bool diskBacked = false,
        ExcelTemporaryStorageBudget temporaryStorage = null)
    {
        var entry = archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null)
            return ExcelForwardXlsxSharedStrings.FromValues(Array.Empty<string>());
        if (diskBacked)
            return ExcelForwardXlsxSharedStrings.ReadDiskBacked(entry, cancellationToken,
                Path.GetTempPath(), temporaryStorage);
        var values = new List<string>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "si", StringComparison.Ordinal))
                continue;
            var builder = new StringBuilder();
            using (var subtree = reader.ReadSubtree())
            {
                while (subtree.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (subtree.NodeType == XmlNodeType.Element
                        && string.Equals(subtree.LocalName, "t", StringComparison.Ordinal))
                        builder.Append(subtree.ReadElementContentAsString());
                }
            }
            values.Add(builder.ToString());
        }
        return ExcelForwardXlsxSharedStrings.FromValues(values);
    }

    /// <summary>
    /// 读取单元格样式及其数字格式。
    /// </summary>
    internal static ExcelForwardXlsxStyles ReadStyles(ZipArchive archive,
        CancellationToken cancellationToken)
    {
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null)
            return ExcelForwardXlsxStyles.Empty;

        var customFormats = new Dictionary<int, string>();
        var cellFormatIds = new List<int>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        var numberFormatsDepth = -1;
        var cellFormatsDepth = -1;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (string.Equals(reader.LocalName, "numFmts", StringComparison.Ordinal))
                    numberFormatsDepth = reader.Depth;
                else if (numberFormatsDepth >= 0 && reader.Depth == numberFormatsDepth + 1
                    && string.Equals(reader.LocalName, "numFmt", StringComparison.Ordinal))
                {
                    if (TryParseInt(reader.GetAttribute("numFmtId"), out var formatId))
                        customFormats[formatId] = reader.GetAttribute("formatCode") ?? string.Empty;
                }
                else if (string.Equals(reader.LocalName, "cellXfs", StringComparison.Ordinal))
                    cellFormatsDepth = reader.Depth;
                else if (cellFormatsDepth >= 0 && reader.Depth == cellFormatsDepth + 1
                    && string.Equals(reader.LocalName, "xf", StringComparison.Ordinal))
                {
                    cellFormatIds.Add(TryParseInt(reader.GetAttribute("numFmtId"), out var formatId)
                        ? formatId : 0);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (reader.Depth == numberFormatsDepth
                    && string.Equals(reader.LocalName, "numFmts", StringComparison.Ordinal))
                    numberFormatsDepth = -1;
                if (reader.Depth == cellFormatsDepth
                    && string.Equals(reader.LocalName, "cellXfs", StringComparison.Ordinal))
                    cellFormatsDepth = -1;
            }
        }
        return new ExcelForwardXlsxStyles(cellFormatIds, customFormats);
    }

    /// <summary>
    /// 按 worksheet XML 行顺序读取单元格值。
    /// </summary>
    internal static IEnumerable<ExcelForwardXlsxRow> ReadRows(Stream stream,
        IReadOnlyList<string> sharedStrings, ExcelForwardXlsxStyles styles, bool date1904,
        CancellationToken cancellationToken)
    {
        using var reader = CreateReader(stream);
        var nextRow = 0;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "row", StringComparison.Ordinal))
                continue;
            var rowIndex = ParseRowIndex(reader.GetAttribute("r"), nextRow);
            nextRow = rowIndex + 1;
            yield return new ExcelForwardXlsxRow(rowIndex,
                ReadRowCells(reader, sharedStrings, styles, date1904, cancellationToken));
        }
    }

    /// <summary>
    /// 读取 worksheet 声明的合并区域。
    /// </summary>
    /// <param name="entry">worksheet XML 部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>按 XML 顺序排列的零基合并区域。</returns>
    /// <remarks>
    /// Forward 只读取合并元数据用于布局预检，不创建工作簿 DOM，也不复制未声明的合并区域。
    /// </remarks>
    internal static IReadOnlyList<ExcelEntityCellRange> ReadMergeRanges(ZipArchiveEntry entry,
        CancellationToken cancellationToken)
    {
        if (entry == null)
            throw new ArgumentNullException(nameof(entry));

        var ranges = new List<ExcelEntityCellRange>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "mergeCell", StringComparison.Ordinal))
                continue;

            var address = reader.GetAttribute("ref");
            if (string.IsNullOrWhiteSpace(address))
                throw new BingOfficesImportException("XLSX 合并区域缺少地址。", provider: "Excel",
                    stage: BingOfficesStage.Preflight);
            try
            {
                ranges.Add(ExcelEntityCellRange.Parse(address));
            }
            catch (ArgumentException exception)
            {
                throw new BingOfficesImportException("XLSX 合并区域地址无效。", exception,
                    provider: "Excel", stage: BingOfficesStage.Preflight);
            }
        }
        return Array.AsReadOnly(ranges.ToArray());
    }

    /// <summary>
    /// 读取 worksheet 声明的显式列展示元数据。
    /// </summary>
    /// <param name="entry">worksheet XML 部件。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    /// <returns>零基列号到列宽和隐藏状态的映射。</returns>
    /// <remarks>
    /// 只保留有限范围内的显式宽度；缺失或超出格式边界的元数据忽略，不影响实体数据读取。
    /// </remarks>
    internal static IReadOnlyDictionary<int, (double Width, bool HasWidth, bool Hidden)>
        ReadColumnMetadata(ZipArchiveEntry entry,
        CancellationToken cancellationToken)
    {
        if (entry == null)
            throw new ArgumentNullException(nameof(entry));

        var metadata = new Dictionary<int, (double Width, bool HasWidth, bool Hidden)>();
        using var stream = entry.Open();
        using var reader = CreateReader(stream);
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element
                || !string.Equals(reader.LocalName, "col", StringComparison.Ordinal))
                continue;
            if (!TryParseInt(reader.GetAttribute("min"), out var first)
                || !TryParseInt(reader.GetAttribute("max"), out var last)
                || first < 1 || last < first)
                continue;
            var hasWidth = double.TryParse(reader.GetAttribute("width"), NumberStyles.Float,
                CultureInfo.InvariantCulture, out var width)
                && width > 0 && !double.IsNaN(width) && !double.IsInfinity(width);
            var hidden = IsTrue(reader.GetAttribute("hidden"));
            if (!hasWidth && !hidden)
                continue;
            if (first > 16_384)
                continue;
            first = Math.Min(first, 16_384);
            last = Math.Min(last, 16_384);
            for (var column = first; column <= last; column++)
                metadata[column - 1] = (width, hasWidth, hidden);
        }
        return metadata;
    }

    /// <summary>
    /// 解析 XML 中的布尔属性。
    /// </summary>
    /// <param name="value">属性文本。</param>
    /// <returns>是否为 true。</returns>
    private static bool IsTrue(string value) => string.Equals(value, "1", StringComparison.Ordinal)
        || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 读取当前 XML 行中的单元格集合。
    /// </summary>
    private static IReadOnlyDictionary<int, ExcelCellValue> ReadRowCells(XmlReader reader,
        IReadOnlyList<string> sharedStrings, ExcelForwardXlsxStyles styles, bool date1904,
        CancellationToken cancellationToken)
    {
        var cells = new Dictionary<int, ExcelCellValue>();
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (subtree.NodeType != XmlNodeType.Element
                || !string.Equals(subtree.LocalName, "c", StringComparison.Ordinal))
                continue;
            var cell = ReadCell(subtree, sharedStrings, styles, date1904);
            if (cell.ColumnIndex >= 0)
                cells[cell.ColumnIndex] = cell.Value;
        }
        return cells;
    }

    /// <summary>
    /// 读取单个 worksheet 单元格。
    /// </summary>
    private static (int ColumnIndex, ExcelCellValue Value) ReadCell(XmlReader reader,
        IReadOnlyList<string> sharedStrings, ExcelForwardXlsxStyles styles, bool date1904)
    {
        var reference = reader.GetAttribute("r");
        var type = reader.GetAttribute("t");
        var styleIndex = TryParseInt(reader.GetAttribute("s"), out var parsedStyleIndex)
            ? parsedStyleIndex : 0;
        var value = (string)null;
        var formula = (string)null;
        var inline = new StringBuilder();
        using (var subtree = reader.ReadSubtree())
        {
            while (subtree.Read())
            {
                if (subtree.NodeType != XmlNodeType.Element)
                    continue;
                if (string.Equals(subtree.LocalName, "f", StringComparison.Ordinal))
                    formula = subtree.ReadElementContentAsString();
                else if (string.Equals(subtree.LocalName, "v", StringComparison.Ordinal))
                    value = subtree.ReadElementContentAsString();
                else if (string.Equals(subtree.LocalName, "t", StringComparison.Ordinal))
                    inline.Append(subtree.ReadElementContentAsString());
            }
        }
        if (!string.IsNullOrWhiteSpace(formula))
            throw new BingOfficesConfigurationException("Excel Forward 实体批次暂不支持公式单元格。",
                stage: BingOfficesStage.Plan);
        return (ParseColumnIndex(reference), CreateCellValue(type, value, inline.ToString(), sharedStrings,
            styles, styleIndex, date1904));
    }

    /// <summary>
    /// 将 OOXML 单元格内容转换为公共单元格值。
    /// </summary>
    private static ExcelCellValue CreateCellValue(string type, string value, string inline,
        IReadOnlyList<string> sharedStrings, ExcelForwardXlsxStyles styles, int styleIndex,
        bool date1904)
    {
        if (string.Equals(type, "s", StringComparison.Ordinal))
        {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var index)
                || index < 0 || index >= sharedStrings.Count)
                throw new BingOfficesImportException("XLSX shared string 索引无效.", provider: "Excel",
                    stage: BingOfficesStage.Read);
            var text = sharedStrings[index] ?? string.Empty;
            return new ExcelCellValue(text, text, ExcelCellKind.Text, isDate1904: date1904);
        }
        if (string.Equals(type, "inlineStr", StringComparison.Ordinal)
            || string.Equals(type, "str", StringComparison.Ordinal))
        {
            var text = inline ?? value ?? string.Empty;
            return new ExcelCellValue(text, text, ExcelCellKind.Text, isDate1904: date1904);
        }
        if (string.Equals(type, "b", StringComparison.Ordinal))
        {
            var boolean = string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
            return new ExcelCellValue(boolean, boolean ? "TRUE" : "FALSE", ExcelCellKind.Boolean,
                isDate1904: date1904);
        }
        if (string.Equals(type, "e", StringComparison.Ordinal))
            return new ExcelCellValue(null, value ?? string.Empty, ExcelCellKind.Error,
                isDate1904: date1904);
        if (string.IsNullOrEmpty(value))
            return new ExcelCellValue(null, string.Empty, ExcelCellKind.Empty, isDate1904: date1904);
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return new ExcelCellValue(value, value, ExcelCellKind.Text, isDate1904: date1904);

        var formatIndex = styles.GetFormatIndex(styleIndex);
        if (styles.IsDateFormat(styleIndex)
            && new DateTimeExcelValidationRule().TryParseValue(
                new ExcelCellValue(number, value, ExcelCellKind.DateTime, formatIndex: formatIndex,
                    isDate1904: date1904), value, typeof(DateTime), CultureInfo.InvariantCulture, null,
                out var dateValue)
            && dateValue is DateTime date)
        {
            return new ExcelCellValue(date, date.ToString("O", CultureInfo.InvariantCulture),
                ExcelCellKind.DateTime, formatIndex: formatIndex, isDate1904: date1904);
        }
        return new ExcelCellValue(number, value, ExcelCellKind.Number, formatIndex: formatIndex,
            isDate1904: date1904);
    }

    /// <summary>
    /// 创建启用安全设置的 XML 读取器。
    /// </summary>
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
    /// 规范化 ZIP 部件路径。
    /// </summary>
    private static string NormalizePartPath(string target)
    {
        var value = (target ?? string.Empty).Replace('\\', '/');
        while (value.StartsWith("../", StringComparison.Ordinal))
            value = value.Substring(3);
        value = value.TrimStart('/');
        return value.StartsWith("xl/", StringComparison.OrdinalIgnoreCase) ? value : "xl/" + value;
    }

    /// <summary>
    /// 解析 worksheet 单元格引用的零基列号。
    /// </summary>
    private static int ParseColumnIndex(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return -1;
        var column = 0;
        foreach (var character in reference)
        {
            if (character < 'A' && character < 'a' || character > 'Z' && character > 'z')
                break;
            column = checked(column * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        }
        return column - 1;
    }

    /// <summary>
    /// 解析 worksheet 行号。
    /// </summary>
    private static int ParseRowIndex(string reference, int fallback)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return fallback;
        var index = 0;
        while (index < reference.Length && !char.IsDigit(reference[index]))
            index++;
        return index < reference.Length && int.TryParse(reference.Substring(index),
            NumberStyles.Integer, CultureInfo.InvariantCulture, out var row)
            ? Math.Max(0, row - 1) : fallback;
    }

    /// <summary>
    /// 解析非负整数属性。
    /// </summary>
    private static bool TryParseInt(string value, out int result) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result)
        && result >= 0;
}

/// <summary>
/// 保存 Forward XLSX 工作簿的工作表信息。
/// </summary>
internal sealed class ExcelForwardXlsxWorkbook
{
    internal ExcelForwardXlsxWorkbook(IReadOnlyList<ExcelForwardXlsxSheet> sheets, bool date1904,
        IReadOnlyList<ExcelForwardXlsxDefinedName> definedNames)
    {
        Sheets = Array.AsReadOnly((sheets ?? Array.Empty<ExcelForwardXlsxSheet>()).ToArray());
        if (Sheets.Count == 0)
            throw new ArgumentException("工作簿至少需要一个工作表。", nameof(sheets));
        Date1904 = date1904;
        DefinedNames = Array.AsReadOnly((definedNames ?? Array.Empty<ExcelForwardXlsxDefinedName>()).ToArray());
    }

    internal IReadOnlyList<ExcelForwardXlsxSheet> Sheets { get; }
    internal string SheetName => Sheets[0].Name;
    internal string SheetPath => Sheets[0].Path;
    internal bool Date1904 { get; }
    internal IReadOnlyList<ExcelForwardXlsxDefinedName> DefinedNames { get; }

    internal int GetSheetIndex(string name)
    {
        for (var index = 0; index < Sheets.Count; index++)
        {
            if (string.Equals(Sheets[index].Name, name, StringComparison.OrdinalIgnoreCase))
                return index;
        }
        return -1;
    }

    internal ExcelForwardXlsxSheet FindSheet(string name) => Sheets.FirstOrDefault(sheet =>
        string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase));
}

/// <summary>
/// 保存 Forward XLSX 工作簿中的命名范围定义。
/// </summary>
internal sealed class ExcelForwardXlsxDefinedName
{
    internal ExcelForwardXlsxDefinedName(string name, int localSheetId, string formula)
    {
        Name = name;
        LocalSheetId = localSheetId;
        Formula = formula;
    }

    internal string Name { get; }
    internal int LocalSheetId { get; }
    internal string Formula { get; }
}

/// <summary>
/// 保存 Forward XLSX 工作表的名称和部件路径。
/// </summary>
internal sealed class ExcelForwardXlsxSheet
{
    internal ExcelForwardXlsxSheet(string name, string path)
    {
        Name = name;
        Path = path;
    }

    internal string Name { get; }
    internal string Path { get; }
}

/// <summary>
/// 保存 Forward XLSX 的样式和数字格式信息。
/// </summary>
internal sealed class ExcelForwardXlsxStyles
{
    private readonly IReadOnlyList<int> _formatIds;
    private readonly IReadOnlyDictionary<int, string> _customFormats;

    internal ExcelForwardXlsxStyles(IReadOnlyList<int> formatIds,
        IReadOnlyDictionary<int, string> customFormats)
    {
        _formatIds = formatIds ?? Array.Empty<int>();
        _customFormats = customFormats ?? new Dictionary<int, string>();
    }

    internal static ExcelForwardXlsxStyles Empty { get; } = new(Array.Empty<int>(),
        new Dictionary<int, string>());

    internal int GetFormatIndex(int styleIndex) => styleIndex >= 0 && styleIndex < _formatIds.Count
        ? _formatIds[styleIndex] : 0;

    internal bool IsDateFormat(int styleIndex)
    {
        var formatIndex = GetFormatIndex(styleIndex);
        if ((formatIndex >= 14 && formatIndex <= 22)
            || (formatIndex >= 27 && formatIndex <= 36)
            || (formatIndex >= 45 && formatIndex <= 47))
            return true;
        if (!_customFormats.TryGetValue(formatIndex, out var formatCode)
            || string.IsNullOrWhiteSpace(formatCode))
            return false;
        var quoted = false;
        var escaped = false;
        for (var index = 0; index < formatCode.Length; index++)
        {
            var character = formatCode[index];
            if (escaped)
            {
                escaped = false;
                continue;
            }
            if (character == '\\')
            {
                escaped = true;
                continue;
            }
            if (character == '"')
            {
                quoted = !quoted;
                continue;
            }
            if (quoted)
                continue;
            if (character == '[')
            {
                var end = formatCode.IndexOf(']', index + 1);
                if (end < 0)
                    break;
                var bracket = formatCode.Substring(index + 1, end - index - 1);
                if (bracket.IndexOfAny(new[] { 'y', 'Y', 'm', 'M', 'd', 'D', 'h', 'H', 's', 'S' }) >= 0)
                    return true;
                index = end;
                continue;
            }
            if (character == 'y' || character == 'Y' || character == 'm' || character == 'M'
                || character == 'd' || character == 'D' || character == 'h' || character == 'H'
                || character == 's' || character == 'S')
                return true;
        }
        return false;
    }
}

/// <summary>
/// 保存 Forward XLSX 的物理行和单元格值。
/// </summary>
internal sealed class ExcelForwardXlsxRow
{
    internal ExcelForwardXlsxRow(int rowIndex, IReadOnlyDictionary<int, ExcelCellValue> cells)
    {
        RowIndex = rowIndex;
        Cells = cells ?? throw new ArgumentNullException(nameof(cells));
    }

    internal int RowIndex { get; }
    internal IReadOnlyDictionary<int, ExcelCellValue> Cells { get; }
}
