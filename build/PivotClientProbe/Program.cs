using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Tools.PivotClientProbe;
using ClosedXML.Excel;

Console.InputEncoding = Encoding.UTF8;
Console.OutputEncoding = Encoding.UTF8;
var arguments = Enumerable.Range(0, args.Length / 2)
    .ToDictionary(index => args[index * 2], index => args[index * 2 + 1], StringComparer.Ordinal);
if (!arguments.TryGetValue("--output", out var output))
    throw new ArgumentException("Supply --output followed by a generated XLSX path.");
var asynchronous = arguments.GetValueOrDefault("--async", "false") == "true";
var providerName = arguments.GetValueOrDefault("--provider", "ClosedXML");
if (arguments.GetValueOrDefault("--number-format-probe", "false") == "true")
{
    if (!providerName.Equals("ClosedXML", StringComparison.OrdinalIgnoreCase))
        throw new ArgumentException("The engine number-format diagnostic supports only ClosedXML.");
    AdvancedPivotNumberFormatProbe.Run(output);
    return;
}
IExcelExporter exporter = providerName.ToUpperInvariant() switch
{
    "CLOSEDXML" => new ClosedXmlExcelExporter(),
    "NPOI" => new NpoiExcelExporter(),
    _ => throw new ArgumentException("Use ClosedXML or NPOI for --provider.")
};
if (arguments.GetValueOrDefault("--advanced", "false") == "true")
{
    await AdvancedPivotProbe.RunAsync(output, asynchronous, exporter);
    return;
}
var rowCount = int.Parse(arguments.GetValueOrDefault("--rows", "3"), CultureInfo.InvariantCulture);
var pivotCount = int.Parse(arguments.GetValueOrDefault("--pivots", "1"), CultureInfo.InvariantCulture);
var separateSources = arguments.GetValueOrDefault("--different-sources", "false") == "true";
var cancelAtRow = int.Parse(arguments.GetValueOrDefault("--cancel-at-row", "0"), CultureInfo.InvariantCulture);
if (rowCount < 3 || rowCount > 10000 || pivotCount < 1 || pivotCount > 4)
    throw new ArgumentOutOfRangeException(nameof(args), "Probe supports 3..10000 rows and 1..4 pivots.");
if (cancelAtRow < 0 || cancelAtRow >= rowCount)
    throw new ArgumentOutOfRangeException(nameof(args), "Cancellation row must be inside the source.");
output = Path.GetFullPath(output);
Directory.CreateDirectory(Path.GetDirectoryName(output));

var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
var stopwatch = Stopwatch.StartNew();
using var cancellation = new CancellationTokenSource();
double? cancellationRequestedMilliseconds = null;
var enumeratedRows = 0;
var rows = Enumerable.Range(0, rowCount).Select(index => new PivotProbeRow
{
    Warehouse = index % 2 == 0 ? "North" : "South",
    Amount = rowCount == 3 ? (index + 1) * 10 : index % 100 + 1
}).ToArray();
var request = ExcelExport.Workbook(builder =>
{
    var sourceCount = separateSources ? pivotCount : 1;
    for (var index = 0; index < sourceCount; index++)
        builder.AddSheet(index == 0 ? "Data" : $"Data{index + 1}", EnumerateRows());
    for (var index = 0; index < pivotCount; index++)
        builder.PivotTable(new ExcelPivotTableDefinition
        {
            Name = index == 0 ? "SalesByWarehouse" : $"SalesByWarehouse{index + 1}",
            SourceSheet = !separateSources || index == 0 ? "Data" : $"Data{index + 1}",
            SourceRange = new ExcelRangeDefinition
            {
                StartRow = 0, StartColumn = 0, EndRow = rowCount, EndColumn = 1
            },
            TargetSheet = index == 0 ? "Summary" : $"Summary{index + 1}",
            TargetRow = 1, TargetColumn = 2, RowField = "Warehouse", ValueField = "Amount"
        });
});
var original = new byte[] { 7, 4, 2, 9, 1 };
if (cancelAtRow > 0)
{
    if (File.Exists(output))
        throw new ArgumentException("Use a new generated path for a cancellation probe.");
    File.WriteAllBytes(output, original);
}
var filesBefore = Directory.GetFiles(Path.GetDirectoryName(output)).OrderBy(path => path, StringComparer.Ordinal).ToArray();
try
{
    if (asynchronous)
        await exporter.ExportToFileAsync(request, output, cancellation.Token);
    else
        exporter.ExportToFile(request, output, cancellation.Token);
}
catch (OperationCanceledException exception) when (cancelAtRow > 0)
{
    stopwatch.Stop();
    Require(exception.CancellationToken == cancellation.Token, "Cancellation token differs.");
    Require(cancellationRequestedMilliseconds.HasValue, "Cancellation did not occur during source enumeration.");
    Require(File.ReadAllBytes(output).SequenceEqual(original), "Cancellation damaged the existing target.");
    Require(Directory.GetFiles(Path.GetDirectoryName(output)).OrderBy(path => path, StringComparer.Ordinal)
        .SequenceEqual(filesBefore), "Cancellation left a staging file.");
    var canceledResult = new
    {
        tool = "Bing.Offices.PivotClientProbe", tfm = RuntimeInformation.FrameworkDescription,
        output, provider = exporter.GetType().Name, asynchronous, rowCount, pivotCount, separateSources, cancelAtRow, enumeratedRows,
        canceled = true, originalTargetPreserved = true, stagingCleaned = true,
        cancellationResponseMilliseconds = stopwatch.Elapsed.TotalMilliseconds - cancellationRequestedMilliseconds.Value
    };
    File.WriteAllText(output + ".probe.json", JsonSerializer.Serialize(canceledResult,
        new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n", StringComparison.Ordinal)
        + "\n", new UTF8Encoding(false));
    Console.WriteLine($"pivot-probe-cancellation-ok rows={rowCount} at={cancelAtRow} async={asynchronous}");
    return;
}
Require(cancelAtRow == 0, "Requested cancellation was not observed.");
stopwatch.Stop();
var allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
using var process = Process.GetCurrentProcess();
var peakWorkingSet = process.PeakWorkingSet64;

using var workbook = new XLWorkbook(output);
var sourceNames = Enumerable.Range(0, separateSources ? pivotCount : 1)
    .Select(index => index == 0 ? "Data" : $"Data{index + 1}").ToArray();
var targetNames = Enumerable.Range(0, pivotCount)
    .Select(index => index == 0 ? "Summary" : $"Summary{index + 1}").ToArray();
Require(workbook.Worksheets.Select(sheet => sheet.Name).SequenceEqual(sourceNames.Concat(targetNames)),
    "Complete worksheet order differs.");
foreach (var name in sourceNames)
{
    var sheet = workbook.Worksheet(name);
    Require(sheet.Range("A1:B1").Cells().Select(cell => cell.GetString())
        .SequenceEqual(new[] { "Warehouse", "Amount" }), "Complete source headers differ.");
    Require(sheet.LastRowUsed().RowNumber() == rowCount + 1, "Source row count differs.");
    for (var index = 0; index < rows.Length; index++)
    {
        Require(sheet.Cell(index + 2, 1).GetString() == rows[index].Warehouse, "Source group differs.");
        Require(sheet.Cell(index + 2, 2).GetDouble() == rows[index].Amount, "Source amount differs.");
    }
}
using var archive = ZipFile.OpenRead(output);
XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var cacheEntries = archive.Entries.Where(entry => entry.FullName.Contains("pivotCache/")
    && entry.FullName.Contains("pivotCacheDefinition") && entry.FullName.EndsWith(".xml")).ToArray();
var recordEntries = archive.Entries.Where(entry => entry.FullName.Contains("pivotCache/")
    && entry.FullName.Contains("pivotCacheRecords") && entry.FullName.EndsWith(".xml")).ToArray();
Require(cacheEntries.Length >= sourceNames.Length && cacheEntries.Length <= pivotCount
    && recordEntries.Length == cacheEntries.Length, "Cache and record parts are incomplete.");
var cacheFieldsByRecord = new Dictionary<string, XElement[]>(StringComparer.Ordinal);
var cachedSourceNames = new HashSet<string>(StringComparer.Ordinal);
foreach (var entry in cacheEntries)
{
    using var stream = entry.Open();
    var root = XDocument.Load(stream).Root;
    Require(root.Element(ns + "cacheFields").Elements().Select(field => (string)field.Attribute("name"))
        .SequenceEqual(new[] { "Warehouse", "Amount" }), "Cache fields differ.");
    Require((string)root.Element(ns + "cacheSource").Element(ns + "worksheetSource").Attribute("ref")
        == $"A1:B{rowCount + 1}", "Cache source range differs.");
    cachedSourceNames.Add((string)root.Element(ns + "cacheSource")
        .Element(ns + "worksheetSource").Attribute("sheet"));
    cacheFieldsByRecord.Add(entry.FullName.Replace("pivotCacheDefinition", "pivotCacheRecords"),
        root.Element(ns + "cacheFields").Elements().ToArray());
}
Require(cachedSourceNames.SetEquals(sourceNames), "Complete cache sources differ.");
foreach (var entry in recordEntries)
{
    using var stream = entry.Open();
    var records = XDocument.Load(stream).Root.Elements(ns + "r").ToArray();
    Require(records.Length == rowCount, "Cache record count differs.");
    for (var index = 0; index < records.Length; index++)
    {
        var cells = records[index].Elements().ToArray();
        Require(cells.Length == 2, "Cache record width differs.");
        var fields = cacheFieldsByRecord[entry.FullName];
        var group = Decode(cells[0], fields[0], ns);
        var amount = Decode(cells[1], fields[1], ns);
        Require(group.Name.LocalName == "s"
            && (string)group.Attribute("v") == rows[index].Warehouse,
            "Complete cache group record differs.");
        Require(amount.Name.LocalName == "n" && double.Parse((string)amount.Attribute("v"),
            CultureInfo.InvariantCulture) == rows[index].Amount, "Complete cache amount record differs.");
    }
}
var visibleSummaryCells = workbook.Worksheet("Summary").CellsUsed()
    .Select(cell => new { address = cell.Address.ToString(), value = cell.GetString() }).ToArray();
var result = new
{
    tool = "Bing.Offices.PivotClientProbe", tfm = RuntimeInformation.FrameworkDescription,
    output, provider = exporter.GetType().Name, asynchronous, rowCount, pivotCount, separateSources, enumeratedRows,
    generationMilliseconds = stopwatch.Elapsed.TotalMilliseconds, allocatedBytes, peakWorkingSet,
    fileBytes = new FileInfo(output).Length, cacheCount = cacheEntries.Length,
    sourceNames, visibleSummaryCells,
    expected = rows.GroupBy(row => row.Warehouse).ToDictionary(group => group.Key, group => group.Sum(row => row.Amount))
};
File.WriteAllText(output + ".probe.json", JsonSerializer.Serialize(result,
    new JsonSerializerOptions { WriteIndented = true }).Replace("\r\n", "\n", StringComparison.Ordinal)
    + "\n", new UTF8Encoding(false));
Console.WriteLine($"pivot-probe-ok rows={rowCount} pivots={pivotCount} async={asynchronous}");

static void Require(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

static XElement Decode(XElement value, XElement field, XNamespace ns) => value.Name.LocalName == "x"
    ? field.Element(ns + "sharedItems").Elements().ElementAt(int.Parse((string)value.Attribute("v"),
        CultureInfo.InvariantCulture))
    : value;

IEnumerable<PivotProbeRow> EnumerateRows()
{
    for (var index = 0; index < rows.Length; index++)
    {
        if (cancelAtRow > 0 && index == cancelAtRow)
        {
            cancellationRequestedMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
            cancellation.Cancel();
        }
        enumeratedRows++;
        yield return rows[index];
    }
}
