using System.Diagnostics;
using System.Text;
using System.Text.Json;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 不依赖 BenchmarkDotNet 自动生成工程的 MiniExcel 受控往返探针。
/// </summary>
internal static class MiniExcelProbe
{
    public static void Run(string artifactPath, int rowCount)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        using var services = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        var exporter = services.GetRequiredService<IExcelExporter>();
        var importer = services.GetRequiredService<IExcelImporter>();
        var rows = Enumerable.Range(0, rowCount).Select(index => new ProbeRow
        {
            Code = $"PROBE-{index:D6}",
            Quantity = index,
            Description = $"mini-excel-{index}"
        }).ToArray();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
        var importRequest = ExcelImport.Workbook<ProbeWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows));

        var sync = Measure(rowCount, () =>
        {
            var content = ExcelStreamExtensions.ExportToBytes(exporter, exportRequest);
            var result = ExcelStreamExtensions.ImportFromBytes(importer, content, importRequest);
            Ensure(result.Workbook.Rows.Count == rowCount, "MiniExcel sync probe row count mismatch.");
            return content.Length;
        });
        var asynchronous = Measure(rowCount, () =>
        {
            var content = ExcelStreamExtensions.ExportToBytesAsync(exporter, exportRequest)
                .GetAwaiter().GetResult();
            var result = ExcelStreamExtensions.ImportFromBytesAsync(importer, content, importRequest)
                .GetAwaiter().GetResult();
            Ensure(result.Workbook.Rows.Count == rowCount, "MiniExcel async probe row count mismatch.");
            return content.Length;
        });

        var document = new
        {
            kind = "miniexcel-controlled-probe",
            schema = 2,
            generatedUtc = DateTimeOffset.UtcNow,
            framework = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            rowCount,
            sync,
            asynchronous
        };
        File.WriteAllText(fullPath, JsonSerializer.Serialize(document), new UTF8Encoding(false));
        Console.WriteLine($"MINIEXCEL_PROBE artifact={fullPath} rows={rowCount} "
            + $"syncMs={sync.elapsedMilliseconds:F2} asyncMs={asynchronous.elapsedMilliseconds:F2}");
    }

    private static ProbeResult Measure(int rowCount, Func<int> operation)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        using var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var workingSetBefore = process.WorkingSet64;
        var stopwatch = Stopwatch.StartNew();
        var outputBytes = operation();
        stopwatch.Stop();
        process.Refresh();
        var workingSetAfter = process.WorkingSet64;
        return new ProbeResult(
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore,
            outputBytes,
            rowCount / stopwatch.Elapsed.TotalSeconds,
            GC.CollectionCount(0) - gen0Before,
            GC.CollectionCount(1) - gen1Before,
            GC.CollectionCount(2) - gen2Before,
            Math.Max(workingSetBefore, workingSetAfter),
            process.PeakWorkingSet64);
    }

    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    private sealed class ProbeRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    private sealed class ProbeWorkbook
    {
        public List<ProbeRow> Rows { get; } = new();
    }

    private sealed record ProbeResult(
        double elapsedMilliseconds,
        long allocatedBytes,
        int outputBytes,
        double rowsPerSecond,
        int gen0Collections,
        int gen1Collections,
        int gen2Collections,
        long workingSetBytes,
        long peakWorkingSetBytes);
}
