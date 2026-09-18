using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using System.Xml;
using Bing.Offices.Exceptions;

if (args.Length >= 2 && string.Equals(args[0], "--staging-matrix", StringComparison.OrdinalIgnoreCase))
    return StagingResourceMatrix.Run(args[1], args.Length >= 3 ? int.Parse(args[2]) : 100000,
        args.Length >= 4 ? args[3] : null, args.Length >= 5 ? args[4] : null);
if (args.Length >= 2 && string.Equals(args[0], "--staging-entrypoints", StringComparison.OrdinalIgnoreCase))
    return StagingEntrypointMatrix.Run(args[1]);
if (args.Length >= 2 && string.Equals(args[0], "--staging-entrypoint-failure-probe", StringComparison.OrdinalIgnoreCase))
    return StagingEntrypointMatrix.RunFailureProbe(args[1]);
if (args.Length >= 6 && string.Equals(args[0], "--staging-scenario", StringComparison.OrdinalIgnoreCase))
    return StagingResourceMatrix.RunScenario(args[1], args[2], args[3], int.Parse(args[4]), int.Parse(args[5]),
        args.Length >= 7 ? args[6] : null);
if (args.Length != 2)
    return 2;

var inputPath = Path.GetFullPath(args[0]);
var mode = args[1];
if (!new[] { "zip", "dom", "dom-limit", "shared-strings", "styles", "drawings", "ole",
    "zip-total-limit", "zip-ratio-limit", "shared-strings-limit", "styles-limit", "worksheet-limit",
    "xml-depth-limit", "xml-character-limit" }
        .Contains(mode, StringComparer.Ordinal))
    return 2;

var limits = mode switch
{
    "dom-limit" => new ExcelResourceLimits { MaxRows = 100 },
    "zip-total-limit" => new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 1 },
    "zip-ratio-limit" => new ExcelResourceLimits { MaxZipCompressionRatio = 1 },
    "shared-strings-limit" => new ExcelResourceLimits { MaxSharedStringsBytes = 32 },
    "styles-limit" => new ExcelResourceLimits { MaxStylesBytes = 32 },
    "worksheet-limit" => new ExcelResourceLimits { MaxWorksheetBytes = 32 },
    "xml-depth-limit" => new ExcelResourceLimits { MaxXmlDepth = 2 },
    "xml-character-limit" => new ExcelResourceLimits { MaxXmlCharacters = 32 },
    _ => null
};
var request = ExcelImport.Workbook<ProbeWorkbook>(builder => builder
    .ResourceLimits(limits)
    .Sheet("Data", root => root.Rows));
var stopwatch = Stopwatch.StartNew();
var inputBytes = new FileInfo(inputPath).Length;

try
{
    var metrics = IsPreflightLimitMode(mode)
        ? new PreflightMetrics(inputBytes, -1, -1, -1, -1, -1, -1, -1, true)
        : ReadPreflightMetrics(inputPath, inputBytes);

    using var input = File.OpenRead(inputPath);
    var serviceCollection = new ServiceCollection();
    serviceCollection.AddBingOfficesNpoi();
    using var services = serviceCollection.BuildServiceProvider();
    var result = services.GetRequiredService<IExcelImporter>().Import(input, request);
    stopwatch.Stop();
    var resourceLimit = result.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit);
    var status = resourceLimit ? "resource-limit" : result.IsSuccess ? "success" : "errors";
    var importedRows = result.Sheets.Sum(sheet => sheet.SourceRows.Count);
    Console.WriteLine($"mode={mode};status={status};rejectStage=none;inputBytes={metrics.InputBytes};sheets={metrics.Sheets};rows={metrics.Rows};"
        + $"importedRows={importedRows};columns={metrics.Columns};cells={metrics.Cells};"
        + $"sharedStrings={metrics.SharedStrings};styles={metrics.Styles};pictures={metrics.Pictures};"
        + $"elapsedMs={stopwatch.ElapsedMilliseconds};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64};"
        + $"errors={result.Errors.Count}");
    return 0;
}
catch (BingOfficesResourceLimitException exception)
{
    stopwatch.Stop();
    Console.WriteLine($"mode={mode};status=resource-limit;rejectStage={exception.Stage};inputBytes={inputBytes};"
        + "sheets=-1;rows=-1;importedRows=0;columns=-1;cells=-1;sharedStrings=-1;styles=-1;pictures=-1;"
        + $"elapsedMs={stopwatch.ElapsedMilliseconds};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64};errors=0");
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"mode={mode};exception={exception.GetType().Name};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64}");
    return 1;
}

static PreflightMetrics ReadPreflightMetrics(string inputPath, long inputBytes)
{
    if (!string.Equals(Path.GetExtension(inputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        return new PreflightMetrics(inputBytes, -1, -1, -1, -1, -1, -1, -1, false);

    using var input = File.OpenRead(inputPath);
    using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: false);
    var workbookEntry = archive.GetEntry("xl/workbook.xml");
    if (workbookEntry == null)
        return new PreflightMetrics(inputBytes, -1, -1, -1, -1, -1, -1, -1, false);

    var sheets = CountElements(workbookEntry, "sheet");
    var rows = 0;
    var columns = 0;
    var cells = 0;
    foreach (var entry in archive.Entries.Where(entry => entry.FullName.StartsWith("xl/worksheets/sheet",
                 StringComparison.OrdinalIgnoreCase) && entry.FullName.EndsWith(".xml",
                 StringComparison.OrdinalIgnoreCase)))
    {
        var worksheet = ReadWorksheetMetrics(entry);
        rows += worksheet.Rows;
        columns = Math.Max(columns, worksheet.Columns);
        cells += worksheet.Cells;
    }

    var sharedStrings = CountEntryElements(archive.GetEntry("xl/sharedStrings.xml"), "si");
    var styles = CountEntryElements(archive.GetEntry("xl/styles.xml"), "xf");
    var pictures = archive.Entries
        .Where(entry => entry.FullName.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase)
                        && entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
        .Sum(entry => CountElements(entry, "pic"));
    return new PreflightMetrics(inputBytes, sheets, rows, columns, cells, sharedStrings, styles, pictures, true);
}

static bool IsPreflightLimitMode(string mode) => mode == "zip-total-limit"
    || mode == "zip-ratio-limit"
    || mode == "shared-strings-limit"
    || mode == "styles-limit"
    || mode == "worksheet-limit"
    || mode == "xml-depth-limit"
    || mode == "xml-character-limit";

static WorksheetMetrics ReadWorksheetMetrics(ZipArchiveEntry entry)
{
    var rows = 0;
    var columns = 0;
    var cells = 0;
    using var stream = entry.Open();
    using var reader = CreateXmlReader(stream);
    var rowCells = 0;
    while (reader.Read())
    {
        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
        {
            rows++;
            rowCells = 0;
        }
        else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
        {
            cells++;
            rowCells++;
            columns = Math.Max(columns, rowCells);
        }
    }
    return new WorksheetMetrics(rows, columns, cells);
}

static int CountEntryElements(ZipArchiveEntry entry, string localName) => entry == null
    ? -1
    : CountElements(entry, localName);

static int CountElements(ZipArchiveEntry entry, string localName)
{
    var count = 0;
    using var stream = entry.Open();
    using var reader = CreateXmlReader(stream);
    while (reader.Read())
    {
        if (reader.NodeType == XmlNodeType.Element && reader.LocalName == localName)
            count++;
    }
    return count;
}

static XmlReader CreateXmlReader(Stream stream) => XmlReader.Create(stream, new XmlReaderSettings
{
    DtdProcessing = DtdProcessing.Prohibit,
    XmlResolver = null,
    IgnoreComments = true,
    IgnoreWhitespace = true,
    MaxCharactersFromEntities = 0
});

/// <summary>
/// 记录测试场景的资源或性能指标。
/// </summary>
internal sealed class PreflightMetrics
{
    /// <summary>
    /// 初始化一个 <see cref="PreflightMetrics" /> 类型的实例。
    /// </summary>
    /// <param name="inputBytes">输入文件字节数。</param>
    /// <param name="sheets">工作表数量。</param>
    /// <param name="rows">统计到的数据行数。</param>
    /// <param name="columns">统计到的列数。</param>
    /// <param name="cells">统计到的单元格数。</param>
    /// <param name="sharedStrings">共享字符串数量。</param>
    /// <param name="styles">样式数量。</param>
    /// <param name="pictures">图片数量。</param>
    /// <param name="isZipMetadata">是否由 ZIP 元数据预检得到统计；false 表示通过工作簿对象模型统计。</param>
    public PreflightMetrics(long inputBytes, int sheets, int rows, int columns, int cells, int sharedStrings,
        int styles, int pictures, bool isZipMetadata)
    {
        InputBytes = inputBytes;
        Sheets = sheets;
        Rows = rows;
        Columns = columns;
        Cells = cells;
        SharedStrings = sharedStrings;
        Styles = styles;
        Pictures = pictures;
        IsZipMetadata = isZipMetadata;
    }

    /// <summary>
    /// 获取输入数据字节数。
    /// </summary>
    public long InputBytes { get; }
    /// <summary>
    /// 获取工作表集合。
    /// </summary>
    public int Sheets { get; }
    /// <summary>
    /// 获取数据行集合。
    /// </summary>
    public int Rows { get; }
    /// <summary>
    /// 获取列集合。
    /// </summary>
    public int Columns { get; }
    /// <summary>
    /// 获取单元格集合。
    /// </summary>
    public int Cells { get; }
    /// <summary>
    /// 获取共享字符串集合。
    /// </summary>
    public int SharedStrings { get; }
    /// <summary>
    /// 获取样式集合。
    /// </summary>
    public int Styles { get; }
    /// <summary>
    /// 获取图片集合。
    /// </summary>
    public int Pictures { get; }
    /// <summary>
    /// 获取是否为 ZIP 元数据。
    /// </summary>
    public bool IsZipMetadata { get; }
}

/// <summary>
/// 记录测试场景的资源或性能指标。
/// </summary>
internal sealed class WorksheetMetrics
{
    /// <summary>
    /// 初始化一个 <see cref="WorksheetMetrics" /> 类型的实例。
    /// </summary>
    /// <param name="rows">工作表行数。</param>
    /// <param name="columns">工作表列数。</param>
    /// <param name="cells">工作表单元格数。</param>
    public WorksheetMetrics(int rows, int columns, int cells)
    {
        Rows = rows;
        Columns = columns;
        Cells = cells;
    }

    /// <summary>
    /// 获取数据行集合。
    /// </summary>
    public int Rows { get; }
    /// <summary>
    /// 获取列集合。
    /// </summary>
    public int Columns { get; }
    /// <summary>
    /// 获取单元格集合。
    /// </summary>
    public int Cells { get; }
}

/// <summary>
/// 表示 Excel 测试使用的工作簿数据模型。
/// </summary>
public sealed class ProbeWorkbook
{
    /// <summary>
    /// 获取或设置数据行集合。
    /// </summary>
    public List<ProbeRow> Rows { get; set; } = new List<ProbeRow>();
}

/// <summary>
/// 表示测试使用的一行数据模型。
/// </summary>
public sealed class ProbeRow
{
    /// <summary>
    /// 获取或设置名称。
    /// </summary>
    public string Name { get; set; }
}
