using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

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
    ProbeMetrics metrics;
    var inputBytes = new FileInfo(inputPath).Length;
    using (var metricsInput = File.OpenRead(inputPath))
    using (var metricsWorkbook = WorkbookFactory.Create(metricsInput))
        metrics = ReadMetrics(inputBytes, metricsWorkbook);

    using var input = File.OpenRead(inputPath);
    var serviceCollection = new ServiceCollection();
    serviceCollection.AddNpoi();
    using var services = serviceCollection.BuildServiceProvider();
    var result = services.GetRequiredService<IExcelImporter>().Import(input, request);
    stopwatch.Stop();
    var resourceLimit = result.Errors.Any(error => error.Code == ExcelImportErrorCode.ResourceLimit);
    var status = resourceLimit ? "resource-limit" : result.IsSuccess ? "success" : "errors";
    Console.WriteLine($"mode={mode};status={status};inputBytes={metrics.InputBytes};sheets={metrics.Sheets};rows={metrics.Rows};"
        + $"columns={metrics.Columns};cells={metrics.Cells};sharedStrings={metrics.SharedStrings};styles={metrics.Styles};"
        + $"pictures={metrics.Pictures};elapsedMs={stopwatch.ElapsedMilliseconds};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64};"
        + $"errors={result.Errors.Count}");
    return 0;
}
catch (Exception exception)
{
    Console.Error.WriteLine($"mode={mode};exception={exception.GetType().Name};peakWorkingSet={Process.GetCurrentProcess().PeakWorkingSet64}");
    return 1;
}

static ProbeMetrics ReadMetrics(long inputBytes, IWorkbook workbook)
{
    var rows = 0;
    var columns = 0;
    var cells = 0;
    var pictures = 0;
    for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
    {
        var sheet = workbook.GetSheetAt(sheetIndex);
        rows += Math.Max(0, sheet.LastRowNum - sheet.FirstRowNum + 1);
        for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
                continue;
            columns = Math.Max(columns, row.LastCellNum);
            cells += row.Cells.Count;
        }
        pictures += CountPictures(sheet);
    }

    var sharedStrings = new HashSet<string>(StringComparer.Ordinal);
    for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
    {
        var sheet = workbook.GetSheetAt(sheetIndex);
        for (var rowIndex = sheet.FirstRowNum; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
                continue;
            foreach (var cell in row.Cells)
            {
                if (cell.CellType == CellType.String)
                    sharedStrings.Add(cell.StringCellValue);
            }
        }
    }

    return new ProbeMetrics(inputBytes, workbook.NumberOfSheets, rows, columns, cells,
        sharedStrings.Count, workbook.NumCellStyles, pictures);
}

static int CountPictures(ISheet sheet)
{
    if (sheet is HSSFSheet hssfSheet && hssfSheet.DrawingPatriarch is HSSFShapeContainer hssfDrawing)
        return hssfDrawing.Children.OfType<HSSFPicture>().Count();
    if (sheet is XSSFSheet xssfSheet)
        return xssfSheet.GetRelations().OfType<XSSFDrawing>().SelectMany(drawing => drawing.GetShapes())
            .OfType<XSSFPicture>().Count();
    return 0;
}

internal sealed class ProbeMetrics
{
    public ProbeMetrics(long inputBytes, int sheets, int rows, int columns, int cells, int sharedStrings,
        int styles, int pictures)
    {
        InputBytes = inputBytes;
        Sheets = sheets;
        Rows = rows;
        Columns = columns;
        Cells = cells;
        SharedStrings = sharedStrings;
        Styles = styles;
        Pictures = pictures;
    }

    public long InputBytes { get; }
    public int Sheets { get; }
    public int Rows { get; }
    public int Columns { get; }
    public int Cells { get; }
    public int SharedStrings { get; }
    public int Styles { get; }
    public int Pictures { get; }
}

public sealed class ProbeWorkbook
{
    public List<ProbeRow> Rows { get; set; } = new List<ProbeRow>();
}

public sealed class ProbeRow
{
    public string Name { get; set; }
}
