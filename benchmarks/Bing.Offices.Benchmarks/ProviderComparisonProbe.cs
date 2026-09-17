using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 以相同 workload 采集 NPOI 与 MiniExcel 的可重复 XLSX 往返样本。
/// </summary>
internal static class ProviderComparisonProbe
{
    public static async Task RunAsync(string artifactPath, int rowCount, int repetitions, string phase)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (repetitions < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitions), "正式对照至少需要三次重复。");
        if (string.IsNullOrWhiteSpace(phase))
            throw new ArgumentException("benchmark phase 不能为空。", nameof(phase));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var rows = Enumerable.Range(0, rowCount).Select(index => new ComparisonRow
        {
            Code = $"COMPARE-{index:D6}",
            Quantity = index,
            Description = $"same-workload-{index}"
        }).ToArray();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
        var importRequest = ExcelImport.Workbook<ComparisonWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows));

        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "provider-comparison-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            phase,
            rowCount,
            repetitions,
            framework = RuntimeInformation.FrameworkDescription,
            os = RuntimeInformation.OSDescription,
            processorCount = Environment.ProcessorCount,
            beforeAfterStatus = string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
                ? "baseline-capture"
                : "after-only-current-worktree",
            workload = "same rows and request; export plus import roundtrip",
            baselineNote = "before data must be captured from the pre-change revision; this run does not synthesize it."
        }));

        foreach (var provider in new[] { "npoi", "miniexcel" })
        {
            using var services = string.Equals(provider, "npoi", StringComparison.Ordinal)
                ? new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider()
                : new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
            var exporter = services.GetRequiredService<IExcelExporter>();
            var importer = services.GetRequiredService<IExcelImporter>();

            await MeasureAndWriteAsync(writer, provider, "sync", repetitions, rowCount,
                () => MeasureSync(exporter, importer, exportRequest, importRequest, rowCount)).ConfigureAwait(false);
            await MeasureAndWriteAsync(writer, provider, "async", repetitions, rowCount,
                () => MeasureAsync(exporter, importer, exportRequest, importRequest, rowCount)).ConfigureAwait(false);
        }

        Console.WriteLine($"PROVIDER_COMPARISON_PROBE artifact={fullPath} providerCount=2 "
            + $"rows={rowCount} repetitions={repetitions} phase={phase} status=passed");
    }

    private static async Task MeasureAndWriteAsync(StreamWriter writer, string provider, string mode,
        int repetitions, int rowCount, Func<Task<Sample>> measure)
    {
        await measure().ConfigureAwait(false);
        for (var repetition = 1; repetition <= repetitions; repetition++)
        {
            var sample = await measure().ConfigureAwait(false);
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "provider-comparison-sample",
                provider,
                mode,
                repetition,
                rowCount,
                sample.elapsedMilliseconds,
                sample.allocatedBytes,
                sample.outputBytes,
                sample.rowsPerSecond,
                sample.gen0Collections,
                sample.gen1Collections,
                sample.gen2Collections,
                sample.peakWorkingSetBytes
            }));
            writer.Flush();
        }
    }

    private static Task<Sample> MeasureSync(IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest, ExcelWorkbookImportRequest<ComparisonWorkbook> importRequest,
        int expectedRowCount)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var workingSetBefore = process.WorkingSet64;
        var stopwatch = Stopwatch.StartNew();
        using var stream = new MemoryStream();
        exporter.Export(exportRequest, stream);
        var outputBytes = checked((int)stream.Length);
        stream.Position = 0;
        var result = importer.Import(stream, importRequest);
        stopwatch.Stop();
        Ensure(result.Workbook.Rows.Count == expectedRowCount, "provider comparison row count mismatch");
        process.Refresh();
        return Task.FromResult(new Sample(stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, outputBytes,
            result.Workbook.Rows.Count / stopwatch.Elapsed.TotalSeconds,
            GC.CollectionCount(0) - gen0Before, GC.CollectionCount(1) - gen1Before,
            GC.CollectionCount(2) - gen2Before, Math.Max(workingSetBefore, process.WorkingSet64)));
    }

    private static async Task<Sample> MeasureAsync(IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest, ExcelWorkbookImportRequest<ComparisonWorkbook> importRequest,
        int expectedRowCount)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var workingSetBefore = process.WorkingSet64;
        var stopwatch = Stopwatch.StartNew();
        using var stream = new MemoryStream();
        await exporter.ExportAsync(exportRequest, stream).ConfigureAwait(false);
        var outputBytes = checked((int)stream.Length);
        stream.Position = 0;
        var result = await importer.ImportAsync(stream, importRequest).ConfigureAwait(false);
        stopwatch.Stop();
        Ensure(result.Workbook.Rows.Count == expectedRowCount, "provider comparison row count mismatch");
        process.Refresh();
        return new Sample(stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, outputBytes,
            result.Workbook.Rows.Count / stopwatch.Elapsed.TotalSeconds,
            GC.CollectionCount(0) - gen0Before, GC.CollectionCount(1) - gen1Before,
            GC.CollectionCount(2) - gen2Before, Math.Max(workingSetBefore, process.WorkingSet64));
    }

    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
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

    private sealed record Sample(double elapsedMilliseconds, long allocatedBytes, int outputBytes,
        double rowsPerSecond, int gen0Collections, int gen1Collections, int gen2Collections,
        long peakWorkingSetBytes);
}
