using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using BenchmarkDotNet.Attributes;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 使用真实同目录文件提交路径测量 CSV 和 Excel 导出；与 MemoryStream 基准分开统计。
/// </summary>
[MemoryDiagnoser]
public class RealIoBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    private IServiceProvider _serviceProvider = null!;
    private ICsvExporter _csvExporter = null!;
    private IExcelExporter _excelExporter = null!;
    private IReadOnlyList<RealIoRow> _rows = Array.Empty<RealIoRow>();
    private ExcelWorkbookExportRequest _excelRequest = null!;
    private string _directory = string.Empty;

    [GlobalSetup]
    public void Setup()
    {
        _directory = Path.Combine(Path.GetTempPath(), $"bing-offices-real-io-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_directory);
        _serviceProvider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        _csvExporter = _serviceProvider.GetRequiredService<ICsvExporter>();
        _excelExporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _rows = Enumerable.Range(0, RowCount).Select(index => new RealIoRow
        {
            Code = $"REAL-{index:D6}",
            Quantity = index,
            Description = $"file-io-{index}"
        }).ToArray();
        _excelRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", _rows));
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        (_serviceProvider as IDisposable)?.Dispose();
        if (Directory.Exists(_directory))
            Directory.Delete(_directory, true);
    }

    [Benchmark]
    public long CsvExportFile()
    {
        var path = Path.Combine(_directory, "sync.csv");
        DeleteIfExists(path);
        _csvExporter.ExportToFile(_rows, path);
        return new FileInfo(path).Length;
    }

    [Benchmark]
    public async Task<long> CsvExportFileAsync()
    {
        var path = Path.Combine(_directory, "async.csv");
        DeleteIfExists(path);
        await _csvExporter.ExportToFileAsync(_rows, path).ConfigureAwait(false);
        return new FileInfo(path).Length;
    }

    [Benchmark]
    public long ExcelExportFile()
    {
        var path = Path.Combine(_directory, "sync.xlsx");
        DeleteIfExists(path);
        _excelExporter.ExportToFile(_excelRequest, path);
        return new FileInfo(path).Length;
    }

    [Benchmark]
    public async Task<long> ExcelExportFileAsync()
    {
        var path = Path.Combine(_directory, "async.xlsx");
        DeleteIfExists(path);
        await _excelExporter.ExportToFileAsync(_excelRequest, path).ConfigureAwait(false);
        return new FileInfo(path).Length;
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    private sealed class RealIoRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }
}
