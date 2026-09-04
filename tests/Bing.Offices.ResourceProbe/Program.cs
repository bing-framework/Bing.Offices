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

if (args.Length != 2)
    return 2;

var inputPath = Path.GetFullPath(args[0]);
var mode = args[1];
if (!new[] { "zip", "dom", "dom-limit", "shared-strings", "styles", "drawings", "ole" }
        .Contains(mode, StringComparer.Ordinal))
    return 2;

var request = ExcelImport.Workbook<ProbeWorkbook>(builder => builder
    .ResourceLimits(mode == "dom-limit" ? new ExcelResourceLimits { MaxRows = 100 } : null)
    .Sheet("Data", root => root.Rows));

try
{
    var stopwatch = Stopwatch.StartNew();
    var inputBytes = new FileInfo(inputPath).Length;
    var metrics = ReadPreflightMetrics(inputPath, inputBytes);

    using var input = File.OpenRead(inputPath);
    var serviceCollection = new ServiceCollection();
    serviceCollection.AddBingOfficesNpoi();
    using var services = serviceCollection.BuildServiceProvider();
    var result = services.GetRequiredService<IExcelImporter>().Import(input, request);
    stopwatch.Stop();
    var resourceLimit = result.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit);
    var status = resourceLimit ? "resource-limit" : result.IsSuccess ? "success" : "errors";
    var importedRows = result.Sheets.Sum(sheet => sheet.SourceRows.Count);
    Console.WriteLine($"mode={mode};status={status};inputBytes={metrics.InputBytes};sheets={metrics.Sheets};rows={metrics.Rows};"
        + $"importedRows={importedRows};columns={metrics.Columns};cells={metrics.Cells};"
        + $"sharedStrings={metrics.SharedStrings};styles={metrics.Styles};pictures={metrics.Pictures};"
        + $"elapsedMs={stopwatch.ElapsedMilliseconds};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64};"
        + $"errors={result.Errors.Count}");
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

internal sealed class PreflightMetrics
{
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

    public long InputBytes { get; }
    public int Sheets { get; }
    public int Rows { get; }
    public int Columns { get; }
    public int Cells { get; }
    public int SharedStrings { get; }
    public int Styles { get; }
    public int Pictures { get; }
    public bool IsZipMetadata { get; }
}

internal sealed class WorksheetMetrics
{
    public WorksheetMetrics(int rows, int columns, int cells)
    {
        Rows = rows;
        Columns = columns;
        Cells = cells;
    }

    public int Rows { get; }
    public int Columns { get; }
    public int Cells { get; }
}

public sealed class ProbeWorkbook
{
    public List<ProbeRow> Rows { get; set; } = new List<ProbeRow>();
}

public sealed class ProbeRow
{
    public string Name { get; set; }
}
