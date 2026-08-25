using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// Stream-first 导入与单工作簿导出性能基准。
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 1, warmupCount: 2, iterationCount: 3)]
public class StreamPipelineBenchmarks
{
    /// <summary>
    /// 测量的数据行数。
    /// </summary>
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    /// <summary>
    /// NPOI 导入器。
    /// </summary>
    private IExcelImporter _importer = null!;

    /// <summary>
    /// NPOI 导出器。
    /// </summary>
    private IExcelExporter _exporter = null!;

    private IServiceProvider _serviceProvider = null!;

    /// <summary>
    /// 基准行集合。
    /// </summary>
    private IReadOnlyList<BenchmarkRow> _rows = Array.Empty<BenchmarkRow>();

    private ExcelWorkbookExportRequest _exportRequest = null!;

    private ExcelWorkbookImportRequest<BenchmarkWorkbook> _importRequest = null!;

    /// <summary>
    /// 预生成的导入工作簿字节。
    /// </summary>
    private byte[] _sourceBytes = Array.Empty<byte>();

    /// <summary>
    /// 可复用的导出目标流。
    /// </summary>
    private MemoryStream? _destination;

    /// <summary>
    /// 历史导出路径使用的中间工作簿缓冲流，仅用于对照基准。
    /// </summary>
    private MemoryStream? _legacyDestinationBuffer;

    /// <summary>
    /// 初始化基准数据和导入工作簿。
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        var services = new ServiceCollection();
        services.AddNpoi();
        _serviceProvider = services.BuildServiceProvider();
        _importer = _serviceProvider.GetRequiredService<IExcelImporter>();
        _exporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _rows = Enumerable.Range(0, RowCount)
            .Select(index => new BenchmarkRow
            {
                Code = $"CODE-{index:D6}",
                Quantity = index,
                Amount = index + 0.25m,
                OccurredAt = new DateTime(2026, 8, 13).AddMinutes(index)
            })
            .ToArray();
        _exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", _rows));
        _importRequest = ExcelImport.Workbook<BenchmarkWorkbook>(workbook =>
            workbook.Sheet("Sheet1", root => root.Items));

        using var source = new MemoryStream();
        _exporter.Export(_exportRequest, source);
        _sourceBytes = source.ToArray();
        _destination = new MemoryStream();
        _legacyDestinationBuffer = new MemoryStream();
    }

    /// <summary>
    /// 释放基准目标流。
    /// </summary>
    [GlobalCleanup]
    public void Cleanup()
    {
        _destination?.Dispose();
        _legacyDestinationBuffer?.Dispose();
        (_serviceProvider as IDisposable)?.Dispose();
    }

    /// <summary>
    /// 测量从 XLSX 输入流导入全部行的成本。
    /// </summary>
    /// <returns>成功导入的行数。</returns>
    [Benchmark]
    public int Import()
    {
        using var source = new MemoryStream(_sourceBytes, writable: false);
        return _importer.Import(source, _importRequest).Workbook.Items.Count;
    }

    /// <summary>
    /// 对照历史导入路径额外保留一份输入缓冲，再交给当前 NPOI 导入器。
    /// </summary>
    /// <returns>成功导入的行数。</returns>
    [Benchmark]
    public int ImportWithLegacySourceBuffer()
    {
        using var source = new MemoryStream(_sourceBytes, writable: false);
        using var staging = new MemoryStream();
        source.CopyTo(staging);
        staging.Position = 0;
        return _importer.Import(staging, _importRequest).Workbook.Items.Count;
    }

    /// <summary>
    /// 测量向 XLSX 目标流导出全部行的成本。
    /// </summary>
    /// <returns>生成工作簿的字节数。</returns>
    [Benchmark]
    public long Export()
    {
        _destination!.Position = 0;
        _destination.SetLength(0);
        _exporter.Export(_exportRequest, _destination);
        return _destination.Length;
    }

    /// <summary>
    /// 对照历史导出路径先写入中间工作簿缓冲，再复制到调用方目标流。
    /// </summary>
    /// <returns>生成工作簿的字节数。</returns>
    [Benchmark]
    public long ExportWithLegacyDestinationBuffer()
    {
        _legacyDestinationBuffer!.Position = 0;
        _legacyDestinationBuffer.SetLength(0);
        _destination!.Position = 0;
        _destination.SetLength(0);
        _exporter.Export(_exportRequest, _legacyDestinationBuffer);
        _legacyDestinationBuffer.Position = 0;
        _legacyDestinationBuffer.CopyTo(_destination);
        return _destination.Length;
    }

    /// <summary>
    /// 基准行模型。
    /// </summary>
    private sealed class BenchmarkRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;

        /// <summary>
        /// 数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 发生时间。
        /// </summary>
        public DateTime OccurredAt { get; set; }
    }

    private sealed class BenchmarkWorkbook
    {
        public List<BenchmarkRow> Items { get; } = new();
    }
}
