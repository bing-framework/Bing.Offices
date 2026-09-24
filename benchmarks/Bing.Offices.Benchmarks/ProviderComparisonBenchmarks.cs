using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using BenchmarkDotNet.Attributes;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 使用相同数据和请求比较 NPOI、MiniExcel 与 ClosedXML 的 XLSX 同步/异步往返。
/// </summary>
[MemoryDiagnoser]
public class ProviderComparisonBenchmarks
{
    /// <summary>
    /// 获取或设置基准数据行数。
    /// </summary>
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    /// <summary>
    /// NPOI 基准使用的服务容器。
    /// </summary>
    private IServiceProvider _npoiServices = null!;

    /// <summary>
    /// MiniExcel 基准使用的服务容器。
    /// </summary>
    private IServiceProvider _miniExcelServices = null!;
    /// <summary>
    /// ClosedXML 基准使用的服务容器。
    /// </summary>
    private IServiceProvider _closedXmlServices = null!;
    /// <summary>
    /// NPOI 导出器。
    /// </summary>
    private IExcelExporter _npoiExporter = null!;

    /// <summary>
    /// NPOI 导入器。
    /// </summary>
    private IExcelImporter _npoiImporter = null!;

    /// <summary>
    /// MiniExcel 导出器。
    /// </summary>
    private IExcelExporter _miniExcelExporter = null!;

    /// <summary>
    /// MiniExcel 导入器。
    /// </summary>
    private IExcelImporter _miniExcelImporter = null!;
    /// <summary>
    /// ClosedXML 导出器。
    /// </summary>
    private IExcelExporter _closedXmlExporter = null!;

    /// <summary>
    /// ClosedXML 导入器。
    /// </summary>
    private IExcelImporter _closedXmlImporter = null!;
    /// <summary>
    /// 统一的工作簿导出请求。
    /// </summary>
    private ExcelWorkbookExportRequest _exportRequest = null!;

    /// <summary>
    /// 统一的工作簿导入请求。
    /// </summary>
    private ExcelWorkbookImportRequest<ComparisonWorkbook> _importRequest = null!;

    /// <summary>
    /// 初始化各 Provider 服务和共享基准请求。
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        _npoiServices = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        _miniExcelServices = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        _closedXmlServices = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        _npoiExporter = _npoiServices.GetRequiredService<IExcelExporter>();
        _npoiImporter = _npoiServices.GetRequiredService<IExcelImporter>();
        _miniExcelExporter = _miniExcelServices.GetRequiredService<IExcelExporter>();
        _miniExcelImporter = _miniExcelServices.GetRequiredService<IExcelImporter>();
        _closedXmlExporter = _closedXmlServices.GetRequiredService<IExcelExporter>();
        _closedXmlImporter = _closedXmlServices.GetRequiredService<IExcelImporter>();
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

    /// <summary>
    /// 释放各 Provider 的服务容器。
    /// </summary>
    [GlobalCleanup]
    public void Cleanup()
    {
        (_npoiServices as IDisposable)?.Dispose();
        (_miniExcelServices as IDisposable)?.Dispose();
        (_closedXmlServices as IDisposable)?.Dispose();
    }

    /// <summary>
    /// 测量 NPOI 同步导出和导入往返。
    /// </summary>
    /// <returns>导入后的行数。</returns>
    [Benchmark(Baseline = true)]
    public int NpoiExportImportSync() => ExportImportSync(_npoiExporter, _npoiImporter);

    /// <summary>
    /// 测量 MiniExcel 同步导出和导入往返。
    /// </summary>
    /// <returns>导入后的行数。</returns>
    [Benchmark]
    public int MiniExcelExportImportSync() => ExportImportSync(_miniExcelExporter, _miniExcelImporter);

    /// <summary>
    /// 测量 ClosedXML 同步导出和导入往返。
    /// </summary>
    /// <returns>导入后的行数。</returns>
    [Benchmark]
    public int ClosedXmlExportImportSync() => ExportImportSync(_closedXmlExporter, _closedXmlImporter);

    /// <summary>
    /// 测量 NPOI 异步导出和导入往返。
    /// </summary>
    /// <returns>最终完成后导入的行数。</returns>
    [Benchmark]
    public Task<int> NpoiExportImportAsync() => ExportImportAsync(_npoiExporter, _npoiImporter);

    /// <summary>
    /// 测量 MiniExcel 异步导出和导入往返。
    /// </summary>
    /// <returns>最终完成后导入的行数。</returns>
    [Benchmark]
    public Task<int> MiniExcelExportImportAsync() => ExportImportAsync(_miniExcelExporter, _miniExcelImporter);

    /// <summary>
    /// 测量 ClosedXML 异步导出和导入往返。
    /// </summary>
    /// <returns>最终完成后导入的行数。</returns>
    [Benchmark]
    public Task<int> ClosedXmlExportImportAsync() => ExportImportAsync(_closedXmlExporter, _closedXmlImporter);

    /// <summary>
    /// 使用指定 Provider 同步执行共享导出/导入请求。
    /// </summary>
    /// <param name="exporter">Provider 导出器。</param>
    /// <param name="importer">Provider 导入器。</param>
    /// <returns>导入后的行数。</returns>
    private int ExportImportSync(IExcelExporter exporter, IExcelImporter importer)
    {
        using var stream = new MemoryStream();
        exporter.Export(_exportRequest, stream);
        stream.Position = 0;
        return importer.Import(stream, _importRequest).Workbook.Rows.Count;
    }

    /// <summary>
    /// 使用指定 Provider 异步执行共享导出/导入请求。
    /// </summary>
    /// <param name="exporter">Provider 导出器。</param>
    /// <param name="importer">Provider 导入器。</param>
    /// <returns>最终完成后导入的行数。</returns>
    private async Task<int> ExportImportAsync(IExcelExporter exporter, IExcelImporter importer)
    {
        using var stream = new MemoryStream();
        await exporter.ExportAsync(_exportRequest, stream).ConfigureAwait(false);
        stream.Position = 0;
        var result = await importer.ImportAsync(stream, _importRequest).ConfigureAwait(false);
        return result.Workbook.Rows.Count;
    }

    /// <summary>
    /// Provider 对照使用的行模型。
    /// </summary>
    private sealed class ComparisonRow
    {
        /// <summary>
        /// 获取或设置行编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置描述文本。
        /// </summary>
        public string Description { get; set; } = string.Empty;
    }

    /// <summary>
    /// Provider 对照使用的导入工作簿模型。
    /// </summary>
    private sealed class ComparisonWorkbook
    {
        /// <summary>
        /// 获取导入的数据行集合。
        /// </summary>
        public List<ComparisonRow> Rows { get; } = new();
    }
}
