using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using BenchmarkDotNet.Attributes;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 测量受控延迟和限速目标流下的异步导出；与真实 FileStream 基准分轨。
/// </summary>
[MemoryDiagnoser]
public class RealIoPipelineBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    private IServiceProvider _serviceProvider = null!;
    private ICsvExporter _csvExporter = null!;
    private IExcelExporter _excelExporter = null!;
    private IReadOnlyList<RealIoPipelineRow> _rows = Array.Empty<RealIoPipelineRow>();
    private ExcelWorkbookExportRequest _excelRequest = null!;

    [GlobalSetup]
    public void Setup()
    {
        _serviceProvider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        _csvExporter = _serviceProvider.GetRequiredService<ICsvExporter>();
        _excelExporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _rows = Enumerable.Range(0, RowCount).Select(index => new RealIoPipelineRow
        {
            Code = $"CONTROLLED-{index:D6}",
            Quantity = index,
            Description = $"controlled-io-{index}"
        }).ToArray();
        _excelRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", _rows));
    }

    [GlobalCleanup]
    public void Cleanup() => (_serviceProvider as IDisposable)?.Dispose();

    [Benchmark]
    public Task<long> CsvExportDelayedAsync() => ExportCsvAsync(new DelayedAsyncWriteStream(
        TimeSpan.FromMilliseconds(1)));

    [Benchmark]
    public Task<long> CsvExportThrottledAsync() => ExportCsvAsync(new ThrottledAsyncWriteStream(
        64 * 1024, TimeSpan.FromMilliseconds(1)));

    [Benchmark]
    public Task<long> ExcelExportDelayedAsync() => ExportExcelAsync(new DelayedAsyncWriteStream(
        TimeSpan.FromMilliseconds(1)));

    [Benchmark]
    public Task<long> ExcelExportThrottledAsync() => ExportExcelAsync(new ThrottledAsyncWriteStream(
        64 * 1024, TimeSpan.FromMilliseconds(1)));

    private async Task<long> ExportCsvAsync(AsyncWriteProbeStream destination)
    {
        await using (destination.ConfigureAwait(false))
        {
            await _csvExporter.ExportAsync(_rows, destination).ConfigureAwait(false);
            EnsureAsyncWrites(destination);
            return destination.Length;
        }
    }

    private async Task<long> ExportExcelAsync(AsyncWriteProbeStream destination)
    {
        await using (destination.ConfigureAwait(false))
        {
            await _excelExporter.ExportAsync(_excelRequest, destination).ConfigureAwait(false);
            EnsureAsyncWrites(destination);
            return destination.Length;
        }
    }

    private static void EnsureAsyncWrites(AsyncWriteProbeStream destination)
    {
        if (destination.AsyncWriteCount == 0 || destination.AsyncBytesWritten != destination.Length)
        {
            throw new InvalidOperationException(
                $"受控异步流没有记录完整异步写入: count={destination.AsyncWriteCount}, "
                + $"asyncBytes={destination.AsyncBytesWritten}, length={destination.Length}");
        }
    }

    private sealed class RealIoPipelineRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }
}
