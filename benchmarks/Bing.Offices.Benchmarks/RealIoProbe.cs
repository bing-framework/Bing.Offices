using System.Diagnostics;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 以独立子进程采集 Real IO 和受控异步流的重复样本。
/// </summary>
internal static class RealIoProbe
{
    private static readonly int[] RowCounts = { 1000, 10000, 100000 };
    private static readonly string[] Scenarios =
    {
        "csv-file-sync",
        "csv-file-async",
        "excel-file-sync",
        "excel-file-async",
        "csv-delayed-async",
        "csv-throttled-async",
        "excel-delayed-async",
        "excel-throttled-async"
    };

    public static void Run(string artifactPath, int repetitionCount)
    {
        if (repetitionCount < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitionCount), "正式对照至少需要三次重复。");

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var failed = false;
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "real-io-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            sourceRole = Environment.GetEnvironmentVariable("BING_OFFICES_SOURCE_ROLE") ?? "unknown",
            gitHead = Environment.GetEnvironmentVariable("BING_OFFICES_GIT_HEAD") ?? "not-provided",
            diffIdentity = Environment.GetEnvironmentVariable("BING_OFFICES_DIFF_ID") ?? "not-provided",
            budgetStatus = Environment.GetEnvironmentVariable("BING_OFFICES_BUDGET_STATUS") ?? "UNAPPROVED",
            dotnet = Environment.Version.ToString(),
            framework = RuntimeInformation.FrameworkDescription,
            os = RuntimeInformation.OSDescription,
            processorCount = Environment.ProcessorCount,
            serverGc = GCSettings.IsServerGC,
            repetitionCount,
            rowCounts = RowCounts,
            scenarios = Scenarios,
            percentileDefinition = "nearest-rank on per-operation elapsedMilliseconds samples",
            throughputDefinition = "rowCount divided by elapsed seconds, aggregated across samples"
        }));

        foreach (var rowCount in RowCounts)
        foreach (var scenario in Scenarios)
        {
            var child = RunChild(fullPath, scenario, rowCount, repetitionCount);
            writer.WriteLine(child);
            writer.Flush();
            using var document = JsonDocument.Parse(child);
            if (document.RootElement.GetProperty("exitCode").GetInt32() != 0)
                failed = true;
        }

        Console.WriteLine($"REAL_IO_PROBE artifact={fullPath} scenarios={RowCounts.Length * Scenarios.Length} "
            + $"repetitions={repetitionCount} status={(failed ? "failed" : "passed")}");
        if (failed)
            Environment.ExitCode = 1;
    }

    public static void RunScenario(string artifactPath, string scenario, int rowCount, int repetitionCount)
    {
        try
        {
            if (!Scenarios.Contains(scenario, StringComparer.Ordinal))
                throw new ArgumentException($"未知 Real IO 场景: {scenario}", nameof(scenario));
            if (rowCount < 1)
                throw new ArgumentOutOfRangeException(nameof(rowCount));
            if (repetitionCount < 1)
                throw new ArgumentOutOfRangeException(nameof(repetitionCount));

            using var serviceProvider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
            var csvExporter = serviceProvider.GetRequiredService<ICsvExporter>();
            var excelExporter = serviceProvider.GetRequiredService<IExcelExporter>();
            var rows = Enumerable.Range(0, rowCount).Select(index => new RealIoRow
            {
                Code = $"PROBE-{index:D6}",
                Quantity = index,
                Description = $"probe-io-{index}"
            }).ToArray();
            var excelRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows));
            var workDirectory = Path.Combine(Path.GetTempPath(), $"bing-offices-real-io-probe-{Guid.NewGuid():N}");
            Directory.CreateDirectory(workDirectory);
            try
            {
                ExecuteOnce(scenario, rowCount, csvExporter, excelExporter, rows, excelRequest, workDirectory)
                    .GetAwaiter().GetResult();
                var samples = new List<Sample>(repetitionCount);
                for (var repetition = 1; repetition <= repetitionCount; repetition++)
                {
                    GC.Collect(2, GCCollectionMode.Forced, true, true);
                    var allocatedBefore = GC.GetTotalAllocatedBytes(true);
                    var gen0Before = GC.CollectionCount(0);
                    var gen1Before = GC.CollectionCount(1);
                    var gen2Before = GC.CollectionCount(2);
                    var stopwatch = Stopwatch.StartNew();
                    var result = ExecuteOnce(scenario, rowCount, csvExporter, excelExporter, rows,
                            excelRequest, workDirectory).GetAwaiter().GetResult();
                    stopwatch.Stop();
                    samples.Add(new Sample
                    {
                        repetition = repetition,
                        elapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                        allocatedBytes = GC.GetTotalAllocatedBytes(true) - allocatedBefore,
                        gen0 = GC.CollectionCount(0) - gen0Before,
                        gen1 = GC.CollectionCount(1) - gen1Before,
                        gen2 = GC.CollectionCount(2) - gen2Before,
                        peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64,
                        outputBytes = result.OutputBytes,
                        asyncWriteCount = result.AsyncWriteCount,
                        asyncBytesWritten = result.AsyncBytesWritten
                    });
                }

                var elapsed = samples.Select(sample => sample.elapsedMilliseconds).OrderBy(value => value).ToArray();
                var totalSeconds = samples.Sum(sample => sample.elapsedMilliseconds) / 1000d;
                var resultDocument = new
                {
                    kind = "real-io-scenario",
                    artifact = Path.GetFullPath(artifactPath),
                    sourceRole = Environment.GetEnvironmentVariable("BING_OFFICES_SOURCE_ROLE") ?? "unknown",
                    scenario,
                    rowCount,
                    repetitionCount,
                    sampleCount = samples.Count,
                    meanMilliseconds = samples.Average(sample => sample.elapsedMilliseconds),
                    medianMilliseconds = Percentile(elapsed, 0.50),
                    p95Milliseconds = Percentile(elapsed, 0.95),
                    p99Milliseconds = Percentile(elapsed, 0.99),
                    throughputRowsPerSecond = rowCount * samples.Count / totalSeconds,
                    allocatedBytesMean = samples.Average(sample => sample.allocatedBytes),
                    gen0Mean = samples.Average(sample => sample.gen0),
                    gen1Mean = samples.Average(sample => sample.gen1),
                    gen2Mean = samples.Average(sample => sample.gen2),
                    peakWorkingSetBytesMax = samples.Max(sample => sample.peakWorkingSetBytes),
                    outputBytesMean = samples.Average(sample => sample.outputBytes),
                    outputBytesMin = samples.Min(sample => sample.outputBytes),
                    outputBytesMax = samples.Max(sample => sample.outputBytes),
                    asyncWriteCountMin = samples.Min(sample => sample.asyncWriteCount),
                    asyncBytesWrittenMin = samples.Min(sample => sample.asyncBytesWritten),
                    samples,
                    status = "passed",
                    exception = (string?)null
                };
                Console.WriteLine(JsonSerializer.Serialize(resultDocument));
            }
            finally
            {
                if (Directory.Exists(workDirectory))
                    Directory.Delete(workDirectory, true);
            }
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            Console.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "real-io-scenario",
                sourceRole = Environment.GetEnvironmentVariable("BING_OFFICES_SOURCE_ROLE") ?? "unknown",
                scenario,
                rowCount,
                repetitionCount,
                sampleCount = 0,
                status = "failed",
                exception = exception.ToString()
            }));
            Environment.ExitCode = 1;
        }
    }

    private static string RunChild(string artifactPath, string scenario, int rowCount, int repetitionCount)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("无法解析当前 benchmark 进程路径。");
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = processPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };
        if (Path.GetFileNameWithoutExtension(processPath).Equals("dotnet", StringComparison.OrdinalIgnoreCase))
            process.StartInfo.ArgumentList.Add(Environment.GetCommandLineArgs()[0]);
        process.StartInfo.ArgumentList.Add("--real-io-scenario");
        process.StartInfo.ArgumentList.Add(artifactPath);
        process.StartInfo.ArgumentList.Add(scenario);
        process.StartInfo.ArgumentList.Add(rowCount.ToString());
        process.StartInfo.ArgumentList.Add(repetitionCount.ToString());
        if (!process.Start())
            throw new InvalidOperationException("无法启动 Real IO 探针子进程。");
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        var line = stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault();
        if (string.IsNullOrWhiteSpace(line))
        {
            return JsonSerializer.Serialize(new
            {
                kind = "real-io-child",
                scenario,
                rowCount,
                repetitionCount,
                exitCode = process.ExitCode == 0 ? 1 : process.ExitCode,
                result = (object?)null,
                stderr
            });
        }
        using var parsed = JsonDocument.Parse(line);
        return JsonSerializer.Serialize(new
        {
            kind = "real-io-child",
            scenario,
            rowCount,
            repetitionCount,
            exitCode = process.ExitCode,
            result = parsed.RootElement.Clone(),
            stderr
        });
    }

    private static async Task<OperationResult> ExecuteOnce(string scenario, int rowCount,
        ICsvExporter csvExporter, IExcelExporter excelExporter, IReadOnlyList<RealIoRow> rows,
        ExcelWorkbookExportRequest excelRequest, string workDirectory)
    {
        switch (scenario)
        {
            case "csv-file-sync":
            {
                var path = Path.Combine(workDirectory, "sync.csv");
                DeleteIfExists(path);
                csvExporter.ExportToFile(rows, path);
                return new OperationResult(new FileInfo(path).Length, 0, 0);
            }
            case "csv-file-async":
            {
                var path = Path.Combine(workDirectory, "async.csv");
                DeleteIfExists(path);
                await csvExporter.ExportToFileAsync(rows, path).ConfigureAwait(false);
                return new OperationResult(new FileInfo(path).Length, 0, 0);
            }
            case "excel-file-sync":
            {
                var path = Path.Combine(workDirectory, "sync.xlsx");
                DeleteIfExists(path);
                excelExporter.ExportToFile(excelRequest, path);
                return new OperationResult(new FileInfo(path).Length, 0, 0);
            }
            case "excel-file-async":
            {
                var path = Path.Combine(workDirectory, "async.xlsx");
                DeleteIfExists(path);
                await excelExporter.ExportToFileAsync(excelRequest, path).ConfigureAwait(false);
                return new OperationResult(new FileInfo(path).Length, 0, 0);
            }
            case "csv-delayed-async":
            {
                using var destination = new DelayedAsyncWriteStream(TimeSpan.FromMilliseconds(1));
                await csvExporter.ExportAsync(rows, destination).ConfigureAwait(false);
                return Complete(destination);
            }
            case "csv-throttled-async":
            {
                using var destination = new ThrottledAsyncWriteStream(64 * 1024,
                    TimeSpan.FromMilliseconds(1));
                await csvExporter.ExportAsync(rows, destination).ConfigureAwait(false);
                return Complete(destination);
            }
            case "excel-delayed-async":
            {
                using var destination = new DelayedAsyncWriteStream(TimeSpan.FromMilliseconds(1));
                await excelExporter.ExportAsync(excelRequest, destination).ConfigureAwait(false);
                return Complete(destination);
            }
            case "excel-throttled-async":
            {
                using var destination = new ThrottledAsyncWriteStream(64 * 1024,
                    TimeSpan.FromMilliseconds(1));
                await excelExporter.ExportAsync(excelRequest, destination).ConfigureAwait(false);
                return Complete(destination);
            }
            default:
                throw new ArgumentOutOfRangeException(nameof(scenario), scenario, null);
        }
    }

    private static OperationResult Complete(AsyncWriteProbeStream destination)
    {
        if (destination.AsyncWriteCount == 0 || destination.AsyncBytesWritten != destination.Length)
            throw new InvalidOperationException("受控流未完成异步写入。");
        return new OperationResult(destination.Length, destination.AsyncWriteCount,
            destination.AsyncBytesWritten);
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
            File.Delete(path);
    }

    private static double Percentile(double[] sortedValues, double percentile)
    {
        var index = (int)Math.Ceiling(sortedValues.Length * percentile) - 1;
        return sortedValues[Math.Clamp(index, 0, sortedValues.Length - 1)];
    }

    private sealed class RealIoRow
    {
        public string Code { get; set; } = string.Empty;
        public int Quantity { get; set; }
        public string Description { get; set; } = string.Empty;
    }

    private sealed class Sample
    {
        public int repetition { get; set; }
        public double elapsedMilliseconds { get; set; }
        public long allocatedBytes { get; set; }
        public int gen0 { get; set; }
        public int gen1 { get; set; }
        public int gen2 { get; set; }
        public long peakWorkingSetBytes { get; set; }
        public long outputBytes { get; set; }
        public long asyncWriteCount { get; set; }
        public long asyncBytesWritten { get; set; }
    }

    private readonly record struct OperationResult(long OutputBytes, long AsyncWriteCount,
        long AsyncBytesWritten);
}
