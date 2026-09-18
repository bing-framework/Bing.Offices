using System.Text;
using Bing.Offices;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.UserModel;

var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Consumer", Guid.NewGuid().ToString("N"));
var packageVersion = Environment.GetEnvironmentVariable("BING_OFFICES_PACKAGE_VERSION");
if (string.IsNullOrWhiteSpace(packageVersion))
    throw new InvalidOperationException("BING_OFFICES_PACKAGE_VERSION must identify the packages under test.");
Directory.CreateDirectory(directory);

try
{
    var services = new ServiceCollection();
    services.AddBingOfficesNpoi();
    using var provider = services.BuildServiceProvider();
    var csvExporter = provider.GetRequiredService<ICsvExporter>();
    var csvImporter = provider.GetRequiredService<ICsvImporter>();
    var excelExporter = provider.GetRequiredService<IExcelExporter>();
    var excelImporter = provider.GetRequiredService<IExcelImporter>();
    var rows = new[] { new ConsumerRow { Code = "consumer,net6", Count = 7 } };
    var csvOptions = new CsvExportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) };

    var csvSyncBytes = CsvStreamExtensions.ExportToBytes(csvExporter, rows, csvOptions);
    var csvSyncResult = CsvStreamExtensions.ImportFromBytes<ConsumerRow>(csvImporter, csvSyncBytes,
        new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });
    var csvAsyncBytes = await CsvStreamExtensions.ExportToBytesAsync(csvExporter, rows, csvOptions);
    var csvAsyncResult = await CsvStreamExtensions.ImportFromBytesAsync<ConsumerRow>(csvImporter,
        csvAsyncBytes, new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });

    var csvSyncPath = Path.Combine(directory, "sync.csv");
    csvExporter.ExportToFile(rows, csvSyncPath, csvOptions);
    var csvFileResult = CsvStreamExtensions.ImportFromFile<ConsumerRow>(csvImporter, csvSyncPath,
        new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });
    var csvAsyncPath = Path.Combine(directory, "async.csv");
    await csvExporter.ExportToFileAsync(rows, csvAsyncPath, csvOptions);
    var csvAsyncFileResult = await CsvStreamExtensions.ImportFromFileAsync<ConsumerRow>(csvImporter,
        csvAsyncPath, new CsvImportOptions<ConsumerRow> { Encoding = new UTF8Encoding(false, true) });

    var workbookRequest = ExcelExport.Workbook(workbook => workbook
        .Format(ExcelFormat.Xlsx)
        .AddSheet("Data", rows));
    var importRequest = ExcelImport.Workbook<ConsumerWorkbook>(workbook =>
        workbook.Sheet("Data", root => root.Rows));
    var excelSyncBytes = ExcelStreamExtensions.ExportToBytes(excelExporter, workbookRequest);
    var excelSyncResult = ExcelStreamExtensions.ImportFromBytes(excelImporter, excelSyncBytes, importRequest);
    var excelAsyncBytes = await ExcelStreamExtensions.ExportToBytesAsync(excelExporter, workbookRequest);
    var excelAsyncResult = await ExcelStreamExtensions.ImportFromBytesAsync(excelImporter,
        excelAsyncBytes, importRequest);

    var excelSyncPath = Path.Combine(directory, "sync.xlsx");
    excelExporter.ExportToFile(workbookRequest, excelSyncPath);
    var excelFileResult = ExcelStreamExtensions.ImportFromFile<ConsumerWorkbook>(excelImporter,
        excelSyncPath, importRequest);
    var excelAsyncPath = Path.Combine(directory, "async.xlsx");
    await excelExporter.ExportToFileAsync(workbookRequest, excelAsyncPath);
    var excelAsyncFileResult = await ExcelStreamExtensions.ImportFromFileAsync<ConsumerWorkbook>(
        excelImporter, excelAsyncPath, importRequest);

    var directExporter = new Bing.Offices.Exports.NpoiExcelExporter();
    var directImporter = new Bing.Offices.Imports.NpoiExcelImporter();
    var directBytes = await directExporter.ExportToBytesAsync(workbookRequest);
    var directResult = await directImporter.ImportFromBytesAsync(directBytes, importRequest);

    using var miniProvider = new ServiceCollection()
        .AddBingOfficesMiniExcel()
        .BuildServiceProvider();
    var miniExporter = miniProvider.GetRequiredService<IExcelExporter>();
    var miniImporter = miniProvider.GetRequiredService<IExcelImporter>();
    var miniBytes = await ExcelStreamExtensions.ExportToBytesAsync(miniExporter, workbookRequest);
    var miniResult = await ExcelStreamExtensions.ImportFromBytesAsync(miniImporter, miniBytes, importRequest);

    using var extensionWorkbook = new XSSFWorkbook();
    var extensionSheet = extensionWorkbook.CreateSheet("Extensions");
    extensionSheet.CreateRow(0).Value(0, "extension");

    Ensure(csvSyncResult.Items.Count == 1 && csvAsyncResult.Items.Count == 1,
        "CSV byte sync/async verification failed.");
    Ensure(csvFileResult.Items.Count == 1 && csvAsyncFileResult.Items.Count == 1,
        "CSV file sync/async verification failed.");
    Ensure(excelSyncResult.Workbook.Rows.Count == 1 && excelAsyncResult.Workbook.Rows.Count == 1,
        "Excel byte sync/async verification failed.");
    Ensure(excelFileResult.Workbook.Rows.Count == 1 && excelAsyncFileResult.Workbook.Rows.Count == 1,
        "Excel file sync/async verification failed.");
    Ensure(directResult.Workbook.Rows.Count == 1 && directBytes.Length > 0,
        "Direct NPOI provider verification failed.");
    Ensure(miniResult.Workbook.Rows.Count == 1 && miniBytes.Length > 0,
        "MiniExcel provider verification failed.");
    Ensure(extensionWorkbook.GetExcelFormat() == ExcelFormat.Xlsx
        && extensionSheet.GetRow(0).GetCell(0).GetStringValue() == "extension",
        "NPOI extension verification failed.");

    Console.WriteLine($"package-consumer-ok tfm={System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription} packages=Bing.Offices.Npoi+Bing.Offices.MiniExcel/{packageVersion} csvBytes={csvAsyncBytes.Length} excelBytes={excelAsyncBytes.Length} miniExcelBytes={miniBytes.Length} npoiExtensions=ok");
}
finally
{
    if (Directory.Exists(directory))
        Directory.Delete(directory, true);
}

static void Ensure(bool condition, string message)
{
    if (!condition)
        throw new InvalidOperationException(message);
}

/// <summary>
/// 表示消费者程序用于往返验证的单行数据。
/// </summary>
public sealed class ConsumerRow
{
    /// <summary>
    /// 获取或设置行的编码。
    /// </summary>
    public string Code { get; set; }

    /// <summary>
    /// 获取或设置行中的数量。
    /// </summary>
    public int Count { get; set; }
}

/// <summary>
/// 表示消费者程序用于 Excel 往返验证的工作簿。
/// </summary>
public sealed class ConsumerWorkbook
{
    /// <summary>
    /// 获取工作簿中的数据行集合。
    /// </summary>
    public List<ConsumerRow> Rows { get; } = new();
}
