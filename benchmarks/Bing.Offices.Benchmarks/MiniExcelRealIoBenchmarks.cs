using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using BenchmarkDotNet.Attributes;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 以真实 XLSX 字节往返测量 MiniExcel P0 管线，并覆盖 100K 行受控场景。
/// </summary>
[MemoryDiagnoser]
public class MiniExcelRealIoBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    private IServiceProvider _serviceProvider = null!;
    private IExcelExporter _exporter = null!;
    private IExcelImporter _importer = null!;
    private IReadOnlyList<MiniExcelBenchRow> _rows = Array.Empty<MiniExcelBenchRow>();
    private ExcelWorkbookExportRequest _exportRequest = null!;
    private ExcelWorkbookImportRequest<MiniExcelBenchWorkbook> _importRequest = null!;

    [GlobalSetup]
    public void Setup()
    {
        _serviceProvider = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        _exporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _importer = _serviceProvider.GetRequiredService<IExcelImporter>();
        _rows = Enumerable.Range(0, RowCount).Select(index => new MiniExcelBenchRow
        {
            Code = $"MINI-{index:D6}",
            Quantity = index,
            Description = $"mini-excel-{index}"
        }).ToArray();
        _exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", _rows));
        _importRequest = ExcelImport.Workbook<MiniExcelBenchWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows));
    }

    [GlobalCleanup]
    public void Cleanup() => (_serviceProvider as IDisposable)?.Dispose();

    [Benchmark]
    public int ExportImportSync()
    {
        var content = ExcelStreamExtensions.ExportToBytes(_exporter, _exportRequest);
        var result = ExcelStreamExtensions.ImportFromBytes(_importer, content, _importRequest);
        return result.Workbook.Rows.Count;
    }

    [Benchmark]
    public async Task<int> ExportImportAsync()
    {
        var content = await ExcelStreamExtensions.ExportToBytesAsync(_exporter, _exportRequest)
            .ConfigureAwait(false);
        var result = await ExcelStreamExtensions.ImportFromBytesAsync(_importer, content, _importRequest)
            .ConfigureAwait(false);
        return result.Workbook.Rows.Count;
    }

    private sealed class MiniExcelBenchRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    private sealed class MiniExcelBenchWorkbook
    {
        public List<MiniExcelBenchRow> Rows { get; } = new();
    }
}
