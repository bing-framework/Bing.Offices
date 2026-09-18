using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using BenchmarkDotNet.Attributes;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 使用相同数据和请求比较 NPOI 与 MiniExcel 的 XLSX 同步/异步往返。
/// </summary>
[MemoryDiagnoser]
public class ProviderComparisonBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    private IServiceProvider _npoiServices = null!;
    private IServiceProvider _miniExcelServices = null!;
    private IExcelExporter _npoiExporter = null!;
    private IExcelImporter _npoiImporter = null!;
    private IExcelExporter _miniExcelExporter = null!;
    private IExcelImporter _miniExcelImporter = null!;
    private ExcelWorkbookExportRequest _exportRequest = null!;
    private ExcelWorkbookImportRequest<ComparisonWorkbook> _importRequest = null!;

    [GlobalSetup]
    public void Setup()
    {
        _npoiServices = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        _miniExcelServices = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        _npoiExporter = _npoiServices.GetRequiredService<IExcelExporter>();
        _npoiImporter = _npoiServices.GetRequiredService<IExcelImporter>();
        _miniExcelExporter = _miniExcelServices.GetRequiredService<IExcelExporter>();
        _miniExcelImporter = _miniExcelServices.GetRequiredService<IExcelImporter>();
        var rows = Enumerable.Range(0, RowCount).Select(index => new ComparisonRow
        {
            Code = $"COMPARE-{index:D6}",
            Quantity = index,
            Description = $"same-workload-{index}"
        }).ToArray();
        _exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
        _importRequest = ExcelImport.Workbook<ComparisonWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows));
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        (_npoiServices as IDisposable)?.Dispose();
        (_miniExcelServices as IDisposable)?.Dispose();
    }

    [Benchmark(Baseline = true)]
    public int NpoiExportImportSync() => ExportImportSync(_npoiExporter, _npoiImporter);

    [Benchmark]
    public int MiniExcelExportImportSync() => ExportImportSync(_miniExcelExporter, _miniExcelImporter);

    [Benchmark]
    public Task<int> NpoiExportImportAsync() => ExportImportAsync(_npoiExporter, _npoiImporter);

    [Benchmark]
    public Task<int> MiniExcelExportImportAsync() => ExportImportAsync(_miniExcelExporter, _miniExcelImporter);

    private int ExportImportSync(IExcelExporter exporter, IExcelImporter importer)
    {
        using var stream = new MemoryStream();
        exporter.Export(_exportRequest, stream);
        stream.Position = 0;
        return importer.Import(stream, _importRequest).Workbook.Rows.Count;
    }

    private async Task<int> ExportImportAsync(IExcelExporter exporter, IExcelImporter importer)
    {
        using var stream = new MemoryStream();
        await exporter.ExportAsync(_exportRequest, stream).ConfigureAwait(false);
        stream.Position = 0;
        var result = await importer.ImportAsync(stream, _importRequest).ConfigureAwait(false);
        return result.Workbook.Rows.Count;
    }

    private sealed class ComparisonRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    private sealed class ComparisonWorkbook
    {
        public List<ComparisonRow> Rows { get; } = new();
    }
}
