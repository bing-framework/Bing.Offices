using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
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
    public static async Task RunAsync(string artifactPath, int rowCount, int repetitions, string phase,
        string? providerFilter = null)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (repetitions < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitions), "正式对照至少需要三次重复。");
        if (string.IsNullOrWhiteSpace(phase))
            throw new ArgumentException("benchmark phase 不能为空。", nameof(phase));
        if (!string.IsNullOrWhiteSpace(providerFilter) &&
            !string.Equals(providerFilter, "npoi", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(providerFilter, "miniexcel", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("provider 必须是 npoi 或 miniexcel。", nameof(providerFilter));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var providers = string.IsNullOrWhiteSpace(providerFilter)
            ? new[] { "npoi", "miniexcel" }
            : new[] { providerFilter.ToLowerInvariant() };

        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        var candidateAssemblies = GetCandidateAssemblies();
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
            processId = Environment.ProcessId,
            processorCount = Environment.ProcessorCount,
            candidateIdentity = new
            {
                source = "production-and-benchmark-assembly-sha256",
                assemblies = candidateAssemblies,
                gitCommit = Environment.GetEnvironmentVariable("GITHUB_SHA") ?? "working-tree"
            },
            beforeAfterStatus = string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
                ? "baseline-capture"
                : "after-only-current-worktree",
            providerFilter = providerFilter ?? "all",
            providerCount = providers.Length,
            workload = "same rows and request; export plus import roundtrip",
            baselineNote = "before data must be captured from the pre-change revision; this run does not synthesize it."
        }));

        foreach (var provider in providers)
        {
            foreach (var mode in new[] { "sync", "async" })
            {
                var workerPath = Path.Combine(Path.GetDirectoryName(fullPath)!,
                    $".{Path.GetFileName(fullPath)}.{provider}.{mode}.{Guid.NewGuid():N}.jsonl");
                try
                {
                    var samples = await RunIsolatedModeAsync(workerPath, rowCount, repetitions,
                        provider, mode).ConfigureAwait(false);
                    foreach (var sample in samples)
                        writer.WriteLine(sample.RawJson);
                    writer.WriteLine(JsonSerializer.Serialize(new
                    {
                        kind = "provider-comparison-summary",
                        provider,
                        mode,
                        rowCount,
                        repetitions = samples.Count,
                        medianElapsedMilliseconds = Median(samples.Select(sample => sample.ElapsedMilliseconds)),
                        medianAllocatedBytes = Median(samples.Select(sample => sample.AllocatedBytes)),
                        medianPeakWorkingSetBytes = Median(samples.Select(sample => sample.PeakWorkingSetBytes)),
                        medianRowsPerSecond = Median(samples.Select(sample => sample.RowsPerSecond))
                    }));
                    writer.Flush();
                }
                finally
                {
                    if (File.Exists(workerPath))
                        File.Delete(workerPath);
                }
            }
        }

        Console.WriteLine($"PROVIDER_COMPARISON_PROBE artifact={fullPath} providerCount={providers.Length} "
            + $"rows={rowCount} repetitions={repetitions} phase={phase} status=passed");
    }

    internal static async Task RunWorkerAsync(string artifactPath, int rowCount, int repetitions,
        string provider, string mode)
    {
        if ((provider != "npoi" && provider != "miniexcel") || (mode != "sync" && mode != "async"))
            throw new ArgumentException("worker provider 或 mode 无效。");

        var rows = Enumerable.Range(0, rowCount).Select(index => new ComparisonRow
        {
            Code = $"COMPARE-{index:D6}",
            Quantity = index,
            Description = $"same-workload-{index}"
        }).ToArray();
        var exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
        var importRequest = ExcelImport.Workbook<ComparisonWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows));
        using var services = provider == "npoi"
            ? new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider()
            : new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        var exporter = services.GetRequiredService<IExcelExporter>();
        var importer = services.GetRequiredService<IExcelImporter>();
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(artifactPath))!);
        using var writer = new StreamWriter(artifactPath, false, new UTF8Encoding(false));
        var measure = mode == "sync"
            ? new Func<Task<Sample>>(() => MeasureSync(exporter, importer, exportRequest, importRequest, rowCount))
            : () => MeasureAsync(exporter, importer, exportRequest, importRequest, rowCount);
        await MeasureAndWriteAsync(writer, provider, mode, repetitions, rowCount, measure)
            .ConfigureAwait(false);
    }

    private static async Task<IReadOnlyList<SampleDocument>> RunIsolatedModeAsync(
        string artifactPath, int rowCount, int repetitions, string provider, string mode)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = Environment.ProcessPath ?? throw new InvalidOperationException("无法定位当前进程。"),
            UseShellExecute = false,
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };
        var processName = Path.GetFileNameWithoutExtension(startInfo.FileName);
        if (string.Equals(processName, "dotnet", StringComparison.OrdinalIgnoreCase))
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        startInfo.ArgumentList.Add("--provider-comparison-worker");
        startInfo.ArgumentList.Add(artifactPath);
        startInfo.ArgumentList.Add(rowCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
        startInfo.ArgumentList.Add(repetitions.ToString(System.Globalization.CultureInfo.InvariantCulture));
        startInfo.ArgumentList.Add(provider);
        startInfo.ArgumentList.Add(mode);

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("无法启动性能采样子进程。");
        var errorTask = process.StandardError.ReadToEndAsync();
        var outputTask = process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync().ConfigureAwait(false);
        var error = await errorTask.ConfigureAwait(false);
        var output = await outputTask.ConfigureAwait(false);
        if (process.ExitCode != 0)
            throw new InvalidOperationException($"性能采样子进程失败: {error}{output}");

        var samples = File.ReadLines(artifactPath, Encoding.UTF8)
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(line => new SampleDocument(line))
            .ToArray();
        if (samples.Length != repetitions)
            throw new InvalidOperationException($"性能采样子进程返回 {samples.Length} 个样本，期望 {repetitions} 个。");
        return samples;
    }

    private static IReadOnlyDictionary<string, string> GetCandidateAssemblies()
    {
        var assemblies = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [Path.GetFileName(typeof(ProviderComparisonProbe).Assembly.Location)] =
                HashAssembly(typeof(ProviderComparisonProbe).Assembly.Location),
            [Path.GetFileName(typeof(IExcelExporter).Assembly.Location)] =
                HashAssembly(typeof(IExcelExporter).Assembly.Location),
            [Path.GetFileName(typeof(Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider).Assembly.Location)] =
                HashAssembly(typeof(Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider).Assembly.Location),
            [Path.GetFileName(typeof(Bing.Offices.Exports.NpoiExcelExporter).Assembly.Location)] =
                HashAssembly(typeof(Bing.Offices.Exports.NpoiExcelExporter).Assembly.Location),
            [Path.GetFileName(typeof(Bing.Offices.Exports.MiniExcelExcelExporter).Assembly.Location)] =
                HashAssembly(typeof(Bing.Offices.Exports.MiniExcelExcelExporter).Assembly.Location)
        };
        return assemblies;
    }

    private static string HashAssembly(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0)
            return 0;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 1 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2;
    }

    private static double Median(IEnumerable<long> values) => Median(values.Select(value => (double)value));

    private sealed class SampleDocument
    {
        public SampleDocument(string rawJson)
        {
            RawJson = rawJson;
            var parsed = JsonSerializer.Deserialize<SampleValues>(rawJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new InvalidOperationException("性能样本 JSON 无法解析。");
            ElapsedMilliseconds = parsed.ElapsedMilliseconds;
            AllocatedBytes = parsed.AllocatedBytes;
            RowsPerSecond = parsed.RowsPerSecond;
            PeakWorkingSetBytes = parsed.PeakWorkingSetBytes;
        }

        public string RawJson { get; }
        public double ElapsedMilliseconds { get; set; }
        public long AllocatedBytes { get; set; }
        public double RowsPerSecond { get; set; }
        public long PeakWorkingSetBytes { get; set; }
    }

    private sealed class SampleValues
    {
        public double ElapsedMilliseconds { get; set; }
        public long AllocatedBytes { get; set; }
        public double RowsPerSecond { get; set; }
        public long PeakWorkingSetBytes { get; set; }
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
                processId = Environment.ProcessId,
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
            GC.CollectionCount(2) - gen2Before, process.PeakWorkingSet64));
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
            GC.CollectionCount(2) - gen2Before, process.PeakWorkingSet64);
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
