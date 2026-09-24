using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 以相同 workload 采集 NPOI、MiniExcel 与 ClosedXML 的可重复 XLSX 往返样本。
/// </summary>
internal static class ProviderComparisonProbe
{
    /// <summary>
    /// 运行指定 Provider 的隔离同步/异步对照探针。
    /// </summary>
    /// <param name="artifactPath">JSONL 输出路径。</param>
    /// <param name="rowCount">每次操作的数据行数。</param>
    /// <param name="repetitions">每种 Provider 和模式的正式测量次数，至少为 3。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="providerFilter">可选 Provider 过滤器，支持 <c>npoi</c>、<c>miniexcel</c> 和 <c>closedxml</c>。</param>
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
            !string.Equals(providerFilter, "miniexcel", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(providerFilter, "closedxml", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("provider 必须是 npoi、miniexcel 或 closedxml。", nameof(providerFilter));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var providers = string.IsNullOrWhiteSpace(providerFilter)
            ? new[] { "npoi", "miniexcel", "closedxml" }
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

    /// <summary>
    /// 执行单个 Provider 和模式的子进程测量。
    /// </summary>
    /// <param name="artifactPath">JSONL 输出路径。</param>
    /// <param name="rowCount">每次操作的数据行数。</param>
    /// <param name="repetitions">正式测量次数。</param>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式，取值为 <c>sync</c> 或 <c>async</c>。</param>
    internal static async Task RunWorkerAsync(string artifactPath, int rowCount, int repetitions,
        string provider, string mode)
    {
        if ((provider != "npoi" && provider != "miniexcel" && provider != "closedxml")
            || (mode != "sync" && mode != "async"))
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
        using var services = provider switch
        {
            "npoi" => new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider(),
            "miniexcel" => new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider(),
            _ => new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider()
        };
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

    /// <summary>
    /// 在隔离子进程中运行一个 Provider 和模式的测量。
    /// </summary>
    /// <param name="artifactPath">子进程样本输出路径。</param>
    /// <param name="rowCount">每次操作的数据行数。</param>
    /// <param name="repetitions">要求的样本数。</param>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式。</param>
    /// <returns>已解析的样本文档集合。</returns>
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

    /// <summary>
    /// 获取 Provider 对照探针依赖程序集的身份哈希。
    /// </summary>
    /// <returns>按程序集文件名索引的 SHA-256 哈希。</returns>
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
                HashAssembly(typeof(Bing.Offices.Exports.MiniExcelExcelExporter).Assembly.Location),
            [Path.GetFileName(typeof(Bing.Offices.ClosedXml.Exports.ClosedXmlExcelExporter).Assembly.Location)] =
                HashAssembly(typeof(Bing.Offices.ClosedXml.Exports.ClosedXmlExcelExporter).Assembly.Location)
        };
        return assemblies;
    }

    /// <summary>
    /// 计算程序集文件的 SHA-256 哈希。
    /// </summary>
    /// <param name="path">程序集文件路径。</param>
    /// <returns>大写十六进制 SHA-256 哈希。</returns>
    private static string HashAssembly(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    /// <summary>
    /// 计算双精度样本的中位数。
    /// </summary>
    /// <param name="values">待计算的样本序列。</param>
    /// <returns>样本中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0)
            return 0;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 1 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2;
    }

    /// <summary>
    /// 计算长整数样本的中位数。
    /// </summary>
    /// <param name="values">待计算的样本序列。</param>
    /// <returns>样本中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<long> values) => Median(values.Select(value => (double)value));

    /// <summary>
    /// 保存原始 JSON 并缓存可比较的样本指标。
    /// </summary>
    private sealed class SampleDocument
    {
        /// <summary>
        /// 初始化样本文档。
        /// </summary>
        /// <param name="rawJson">子进程输出的原始 JSON。</param>
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

        /// <summary>
        /// 获取原始 JSON 文本。
        /// </summary>
        public string RawJson { get; }

        /// <summary>
        /// 获取或设置耗时，单位为毫秒。
        /// </summary>
        public double ElapsedMilliseconds { get; set; }

        /// <summary>
        /// 获取或设置托管堆分配字节数。
        /// </summary>
        public long AllocatedBytes { get; set; }

        /// <summary>
        /// 获取或设置处理吞吐量，单位为行每秒。
        /// </summary>
        public double RowsPerSecond { get; set; }

        /// <summary>
        /// 获取或设置进程峰值工作集字节数。
        /// </summary>
        public long PeakWorkingSetBytes { get; set; }
    }

    /// <summary>
    /// 从样本 JSON 中读取的可比较指标。
    /// </summary>
    private sealed class SampleValues
    {
        /// <summary>
        /// 获取或设置耗时，单位为毫秒。
        /// </summary>
        public double ElapsedMilliseconds { get; set; }

        /// <summary>
        /// 获取或设置托管堆分配字节数。
        /// </summary>
        public long AllocatedBytes { get; set; }

        /// <summary>
        /// 获取或设置处理吞吐量，单位为行每秒。
        /// </summary>
        public double RowsPerSecond { get; set; }

        /// <summary>
        /// 获取或设置进程峰值工作集字节数。
        /// </summary>
        public long PeakWorkingSetBytes { get; set; }
    }

    /// <summary>
    /// 预热并写入正式样本。
    /// </summary>
    /// <param name="writer">样本 JSONL 写入器。</param>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式。</param>
    /// <param name="repetitions">正式测量次数。</param>
    /// <param name="rowCount">每次操作的数据行数。</param>
    /// <param name="measure">单次测量工厂。</param>
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

    /// <summary>
    /// 同步执行一次导出和导入往返测量。
    /// </summary>
    /// <param name="exporter">Provider 导出器。</param>
    /// <param name="importer">Provider 导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="expectedRowCount">预期导入行数。</param>
    /// <returns>最终完成的测量结果。</returns>
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

    /// <summary>
    /// 异步执行一次导出和导入往返测量。
    /// </summary>
    /// <param name="exporter">Provider 导出器。</param>
    /// <param name="importer">Provider 导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="expectedRowCount">预期导入行数。</param>
    /// <returns>最终完成的异步测量结果。</returns>
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

    /// <summary>
    /// 在对照探针完整性条件不满足时抛出异常。
    /// </summary>
    /// <param name="condition">需要满足的条件。</param>
    /// <param name="message">条件不满足时使用的异常消息。</param>
    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
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

    /// <summary>
    /// 记录一次 Provider 对照样本的运行时指标。
    /// </summary>
    /// <param name="elapsedMilliseconds">耗时，单位为毫秒。</param>
    /// <param name="allocatedBytes">托管堆分配字节数。</param>
    /// <param name="outputBytes">输出工作簿字节数。</param>
    /// <param name="rowsPerSecond">处理吞吐量，单位为行每秒。</param>
    /// <param name="gen0Collections">第 0 代 GC 次数。</param>
    /// <param name="gen1Collections">第 1 代 GC 次数。</param>
    /// <param name="gen2Collections">第 2 代 GC 次数。</param>
    /// <param name="peakWorkingSetBytes">进程峰值工作集字节数。</param>
    private sealed record Sample(double elapsedMilliseconds, long allocatedBytes, int outputBytes,
        double rowsPerSecond, int gen0Collections, int gen1Collections, int gen2Collections,
        long peakWorkingSetBytes);
}
