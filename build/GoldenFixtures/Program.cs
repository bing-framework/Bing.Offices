using System.Globalization;
using System.IO.Compression;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

const string GeneratorVersion = "golden-fixtures-v1";
const string InitialBaselineUpdateReason = "Initial provider-neutral contract fixture baseline";
var repositoryRoot = Directory.GetCurrentDirectory();
var root = Path.Combine(repositoryRoot, "tests", "Bing.Offices.Testing", "Resources", "Golden");
Directory.CreateDirectory(root);
foreach (var directory in Directory.GetDirectories(root, "*", SearchOption.AllDirectories))
    Directory.CreateDirectory(directory);

var entries = new List<ManifestEntry>();

AddWorkbook("formula/cached-number.xlsx", "formula text with numeric cached value", "Formula!A2 has 1+1 and cached value 2",
    Sheet("Formula", Rows(
        Row(1, CellString("A1", "Value")),
        Row(2, CellFormula("A2", "1+1", "2"))), "A1:A2"));
AddWorkbook("formula/cached-bool.xlsx", "formula text with boolean cached value", "Formula!A2 has TRUE and cached value 1",
    Sheet("Formula", Rows(
        Row(1, CellString("A1", "Value")),
        Row(2, CellFormula("A2", "1=1", "1", "b"))), "A1:A2"));
AddWorkbook("formula/cached-string.xlsx", "formula text with string cached value", "Formula!A2 has CONCAT and cached value text",
    Sheet("Formula", Rows(
        Row(1, CellString("A1", "Value")),
        Row(2, CellFormula("A2", "CONCAT(\"A\",\"B\")", "AB", "str"))), "A1:A2"));
AddWorkbook("formula/no-cache.xlsx", "formula without cached value", "Formula!A2 has formula text and an empty cached result", 
    Sheet("Formula", Rows(
        Row(1, CellString("A1", "Value")),
        Row(2, CellFormula("A2", "1+1", string.Empty))), "A1:A2"));

var date1900 = DateTime.ParseExact("2026-09-04", "yyyy-MM-dd", CultureInfo.InvariantCulture);
var date1900Serial = date1900.ToOADate().ToString(CultureInfo.InvariantCulture);
var date1904Serial = (date1900.ToOADate() - 1462d).ToString(CultureInfo.InvariantCulture);
AddWorkbook("date/date1900.xlsx", "1900 date-system serial", "Data!A2 uses a date style and 1900-system serial", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Date")),
        Row(2, CellNumber("A2", date1900Serial, 1))), "A1:A2"));
AddWorkbookWithDate1904("date/date1904.xlsx", "1904 date-system serial", "Data!A2 uses a date style and workbookPr date1904",
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Date")),
        Row(2, CellNumber("A2", date1904Serial, 1))), "A1:A2"));

AddWorkbook("validation/list-invalid.xlsx", "invalid explicit list value", "Data!A2 violates east/west list", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "invalid"), CellNumber("B2", "1"))),
        "A1:B2", Validation("list", "A2", "\"east,west\"")));
AddWorkbook("validation/whole-between-valid.xlsx", "valid whole-number between rule", "Data!B2 is 5 and must be between 1 and 10", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "OK"), CellNumber("B2", "5"))),
        "A1:B2", Validation("whole", "B2", "1", "10", "between")));
AddWorkbook("validation/whole-less-than-invalid.xlsx", "invalid whole-number comparison", "Data!B2 is 5 and must be less than 5", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "OK"), CellNumber("B2", "5"))),
        "A1:B2", Validation("whole", "B2", "5", null, "lessThan")));
AddWorkbook("validation/decimal-between-valid.xlsx", "valid decimal between rule", "Data!A2 is 5.5 and must be between 1 and 10", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "5.5"), CellNumber("B2", "1"))),
        "A1:B2", Validation("decimal", "A2", "1", "10", "between")));
AddWorkbook("validation/date-between-valid.xlsx", "valid date between rule", "Data!A2 is 2026-09-04 and has a date validation", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "2026-09-04"), CellNumber("B2", "1"))),
        "A1:B2", Validation("date", "A2", date1900Serial, (date1900.ToOADate() + 1).ToString(CultureInfo.InvariantCulture), "between")));
AddWorkbook("validation/time-between-valid.xlsx", "valid time between rule", "Data!A2 is 12:00 and has a time validation", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "12:00:00"), CellNumber("B2", "1"))),
        "A1:B2", Validation("time", "A2", "0.25", "0.75", "between")));
AddWorkbook("validation/text-length-invalid.xlsx", "invalid text length rule", "Data!A2 has six characters and must be shorter than five", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "abcdef"), CellNumber("B2", "1"))),
        "A1:B2", Validation("textLength", "A2", "5", null, "lessThan")));
AddWorkbook("validation/custom-unsupported.xlsx", "unsupported custom formula", "Data!A2 uses INDIRECT and must follow explicit unsupported policy", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Code"), CellString("B1", "Quantity")),
        Row(2, CellString("A2", "OK"), CellNumber("B2", "1"))),
        "A1:B2", Validation("custom", "A2", "INDIRECT(A1)")));

AddWorkbook("relations/parent-child.xlsx", "parent-child relation input", "Parents and Children sheets share OrderNo keys", 
    Sheet("Parents", Rows(
        Row(1, CellString("A1", "OrderNo")),
        Row(2, CellString("A2", "P-001"))), "A1:A2"),
    Sheet("Children", Rows(
        Row(1, CellString("A1", "OrderNo"), CellString("B1", "Name")),
        Row(2, CellString("A2", "P-001"), CellString("B2", "Line 1"))), "A1:B2"));
AddWorkbook("template/styles.xlsx", "basic style template", "Data!A1 is bold and B2 uses a date number format", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Header", 3), CellString("B1", "Date", 3)),
        Row(2, CellString("A2", "Value"), CellNumber("B2", date1900Serial, 1))), "A1:B2"));
AddWorkbook("template/comments.xlsx", "basic comment part", "Data!A1 has a source comment authored by GoldenFixtures", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Header")),
        Row(2, CellString("A2", "Value"))), "A1:A2", null, null, null, true));
AddWorkbook("template/conditional-formatting.xlsx", "basic conditional formatting", "Data!A2:A3 has a positive-value expression rule", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Value")),
        Row(2, CellNumber("A2", "1")),
        Row(3, CellNumber("A3", "-1"))), "A1:A3", null, null,
        "<conditionalFormatting sqref=\"A2:A3\"><cfRule type=\"expression\" dxfId=\"0\" priority=\"1\"><formula>A2&gt;0</formula></cfRule></conditionalFormatting>"));
AddWorkbook("template/merged-cells.xlsx", "basic merged-cell template", "Data!A1:B1 is a merged range", 
    Sheet("Data", Rows(
        Row(1, CellString("A1", "Merged")),
        Row(2, CellString("A2", "Value"), CellString("B2", "Second"))), "A1:B2", null, "A1:B1"));

WriteBinary("security/malformed.zip", Encoding.UTF8.GetBytes("not a zip archive"), "malformed ZIP preflight input", "Input is intentionally not a ZIP package");
var oversized = Path.Combine(root, "security", "oversized-entry.xlsx");
WriteXlsxCore(oversized, new[] { Sheet("Data", Rows(Row(1, CellString("A1", "Value"))), "A1:A1") },
    extraEntryBytes: 1024 * 1024);
Register(oversized, "oversized ZIP entry preflight input", "Package includes xl/media/oversized.bin for resource admission checks");

var manifest = new ManifestFile(GeneratorVersion, DateTime.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), entries);
var manifestPath = Path.Combine(root, "manifest.json");
File.WriteAllText(manifestPath, JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }) + "\n",
    new UTF8Encoding(false));
File.WriteAllText(Path.Combine(root, "README.md"), "# Golden XLSX Fixtures\n\nThese files are provider-neutral Open XML inputs authored by the offline `GoldenFixtures` generator. Runtime contract tests verify the manifest hashes and never regenerate these inputs through a Provider. Update a fixture only with a reviewed contract reason and a regenerated SHA-256 entry.\n",
    new UTF8Encoding(false));

void AddWorkbook(string relativePath, string purpose, string expected, params string[] sheets)
{
    var path = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));
    WriteXlsx(path, sheets);
    Register(path, purpose, expected);
}

void AddWorkbookWithDate1904(string relativePath, string purpose, string expected,
    params string[] sheets)
{
    var path = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));
    WriteXlsxCore(path, sheets, date1904: true);
    Register(path, purpose, expected);
}

void Register(string path, string purpose, string expected)
{
    var bytes = File.ReadAllBytes(path);
    entries.Add(new ManifestEntry(
        Path.GetRelativePath(root, path).Replace(Path.DirectorySeparatorChar, '/'), purpose, expected,
        "Manually authored Open XML by GoldenFixtures", GeneratorVersion,
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), "1", InitialBaselineUpdateReason));
}

void WriteBinary(string relativePath, byte[] bytes, string purpose, string expected)
{
    var path = Path.Combine(root, relativePath.Replace('/', Path.DirectorySeparatorChar));
    Directory.CreateDirectory(Path.GetDirectoryName(path));
    File.WriteAllBytes(path, bytes);
    Register(path, purpose, expected);
}

static void WriteXlsx(string path, params string[] sheets) => WriteXlsxCore(path, sheets, false, null, null, 0);

static void WriteXlsxCore(string path, string[] sheets, bool date1904 = false, string unused = null,
    string unused2 = null, int extraEntryBytes = 0)
{
    Directory.CreateDirectory(Path.GetDirectoryName(path));
    var workbookSheets = new StringBuilder();
    var relationships = new StringBuilder();
    for (var index = 0; index < sheets.Length; index++)
    {
        var name = ExtractSheetName(sheets[index]);
        workbookSheets.Append($"<sheet name=\"{Xml(name)}\" sheetId=\"{index + 1}\" r:id=\"rId{index + 1}\"/>");
        relationships.Append($"<Relationship Id=\"rId{index + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{index + 1}.xml\"/>");
    }
    var hasComments = sheets.Any(sheet => SheetXml(sheet).Contains("<legacyDrawing", StringComparison.Ordinal));
    var contentTypes = new StringBuilder("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
    if (hasComments)
    {
        contentTypes.Append("<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/><Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>");
    }
    for (var index = 0; index < sheets.Length; index++)
        contentTypes.Append($"<Override PartName=\"/xl/worksheets/sheet{index + 1}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
    contentTypes.Append("</Types>");
    relationships.Append($"<Relationship Id=\"rId{sheets.Length + 1}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
    using var file = File.Create(path);
    using var archive = new ZipArchive(file, ZipArchiveMode.Create);
    WriteEntry(archive, "[Content_Types].xml", contentTypes.ToString());
    WriteEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
    WriteEntry(archive, "xl/workbook.xml", $"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><workbookPr date1904=\"{(date1904 ? "1" : "0")}\"/><sheets>{workbookSheets}</sheets></workbook>");
    WriteEntry(archive, "xl/_rels/workbook.xml.rels", $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{relationships}</Relationships>");
    WriteEntry(archive, "xl/styles.xml", Styles());
    for (var index = 0; index < sheets.Length; index++)
    {
        WriteEntry(archive, $"xl/worksheets/sheet{index + 1}.xml", SheetXml(sheets[index]));
        if (SheetXml(sheets[index]).Contains("<legacyDrawing", StringComparison.Ordinal))
        {
            WriteEntry(archive, $"xl/worksheets/_rels/sheet{index + 1}.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing1.vml\"/></Relationships>");
            WriteEntry(archive, "xl/comments1.xml",
                "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>GoldenFixtures</author></authors><commentList><comment ref=\"A1\" authorId=\"0\"><text><t>Source comment</t></text></comment></commentList></comments>");
            WriteEntry(archive, "xl/drawings/vmlDrawing1.vml",
                "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype><v:shape id=\"_x0000_s1025\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 16, 5, 5</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>");
        }
    }
    if (extraEntryBytes > 0)
    {
        var entry = archive.CreateEntry("xl/media/oversized.bin", CompressionLevel.NoCompression);
        using var stream = entry.Open();
        stream.Write(new byte[extraEntryBytes]);
    }
}

static void WriteEntry(ZipArchive archive, string path, string content)
{
    var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
    using var stream = entry.Open();
    var bytes = new UTF8Encoding(false).GetBytes(content);
    stream.Write(bytes);
}

static string Sheet(string name, string rows, string dimension, string validation = null,
    string merge = null, string conditionalFormatting = null, bool comments = false)
{
    var extras = new StringBuilder();
    if (!string.IsNullOrWhiteSpace(merge))
        extras.Append($"<mergeCells count=\"1\"><mergeCell ref=\"{merge}\"/></mergeCells>");
    if (!string.IsNullOrWhiteSpace(validation))
        extras.Append(validation);
    if (!string.IsNullOrWhiteSpace(conditionalFormatting))
        extras.Append(conditionalFormatting);
    if (comments)
        extras.Append("<legacyDrawing r:id=\"rId2\"/>");
    return name + "\u001f" + $"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><dimension ref=\"{dimension}\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetData>{rows}</sheetData>{extras}</worksheet>";
}

static string Rows(params string[] rows) => string.Concat(rows);
static string Row(int number, params string[] cells) => $"<row r=\"{number}\">{string.Concat(cells)}</row>";
static string CellString(string reference, string value, int style = 0) => $"<c r=\"{reference}\"{StyleAttribute(style)} t=\"inlineStr\"><is><t xml:space=\"preserve\">{Xml(value)}</t></is></c>";
static string CellNumber(string reference, string value, int style = 0) => $"<c r=\"{reference}\"{StyleAttribute(style)}><v>{Xml(value)}</v></c>";
static string CellFormula(string reference, string formula, string cached, string type = null) => $"<c r=\"{reference}\"{(type == null ? string.Empty : $" t=\"{type}\"")}><f>{Xml(formula)}</f><v>{Xml(cached)}</v></c>";
static string StyleAttribute(int style) => style == 0 ? string.Empty : $" s=\"{style}\"";
static string Validation(string type, string sqref, string first, string second = null, string operatorName = null)
{
    var builder = new StringBuilder($"<dataValidations count=\"1\"><dataValidation type=\"{type}\"");
    if (!string.IsNullOrWhiteSpace(operatorName)) builder.Append($" operator=\"{operatorName}\"");
    builder.Append($" allowBlank=\"0\" sqref=\"{sqref}\"><formula1>{Xml(first)}</formula1>");
    if (second != null) builder.Append($"<formula2>{Xml(second)}</formula2>");
    return builder.Append("</dataValidation></dataValidations>").ToString();
}

static string Styles() => "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/><numFmt numFmtId=\"165\" formatCode=\"hh:mm:ss\"/></numFmts><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs><cellXfs count=\"4\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/><xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" applyFont=\"1\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles><dxfs count=\"1\"><dxf><font><color rgb=\"FFFF0000\"/></font></dxf></dxfs><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleMedium9\"/></styleSheet>";

static string ExtractSheetName(string sheet)
{
    var separator = sheet.IndexOf('\u001f');
    return separator < 0 ? string.Empty : sheet[..separator];
}

static string SheetXml(string sheet)
{
    var separator = sheet.IndexOf('\u001f');
    return separator < 0 ? sheet : sheet[(separator + 1)..];
}

static string Xml(string value) => SecurityElement.Escape(value ?? string.Empty) ?? string.Empty;

/// <summary>
/// 黄金样本生成清单。
/// </summary>
/// <param name="Generator">样本生成器标识。</param>
/// <param name="GeneratedOn">样本生成日期。</param>
/// <param name="Files">生成文件的清单项。</param>
sealed record ManifestFile(string Generator, string GeneratedOn, IReadOnlyList<ManifestEntry> Files);
/// <summary>
/// 单个黄金样本的来源与校验信息。
/// </summary>
/// <param name="Path">样本文件路径。</param>
/// <param name="Purpose">样本用途。</param>
/// <param name="Expected">预期验证结果。</param>
/// <param name="Provenance">样本来源说明。</param>
/// <param name="Generator">样本生成器标识。</param>
/// <param name="Sha256">样本内容的 SHA-256 摘要。</param>
/// <param name="ContractVersion">样本契约版本。</param>
/// <param name="UpdateReason">样本更新原因。</param>
sealed record ManifestEntry(string Path, string Purpose, string Expected, string Provenance,
    string Generator, string Sha256, string ContractVersion, string UpdateReason);
