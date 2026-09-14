using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using NPOI.SS.UserModel;

internal static class StagingEntrypointMatrix
{
    private const string TaskId = "BO-RC-20260908-002";
    private const string Strategy = "TempFile";
    private const int RepetitionCount = 3;
    private const long HybridThresholdBytes = 8L * 1024 * 1024;
    private static readonly int[] RowCounts = { 1000, 10000, 100000 };
    private static readonly string[] Entrypoints =
    {
        "Export",
        "ExportAsync",
        "ExportToFile",
        "ExportToFileAsync"
    };

    private enum FailureMode
    {
        None,
        Setup,
        Sampler,
        Parser
    }

    public static int Run(string artifactPath)
    {
        var fullPath = Path.GetFullPath(artifactPath);
        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("入口矩阵产物路径必须包含目录。");
        Directory.CreateDirectory(directory);

        var passed = true;
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "excel-staging-entrypoint-matrix",
            schema = 1,
            taskId = TaskId,
            generatedUtc = DateTimeOffset.UtcNow,
            rowCounts = RowCounts,
            entrypoints = Entrypoints,
            strategy = Strategy,
            repetitions = RepetitionCount,
            format = "Xlsx",
            contract = "四个 IExcelExporter 公开入口；每次操作后验证输出 hash、可重开、临时文件峰值和 cleanup。"
        }));
        writer.Flush();

        foreach (var rowCount in RowCounts)
        foreach (var entrypoint in Entrypoints)
        for (var repetition = 1; repetition <= RepetitionCount; repetition++)
        {
            var result = RunOne(entrypoint, rowCount, repetition, FailureMode.None);
            writer.WriteLine(JsonSerializer.Serialize(result));
            writer.Flush();
            passed &= result.Status == "passed";
        }

        Console.WriteLine($"STAGING_ENTRYPOINTS artifact={fullPath} scenarios={RowCounts.Length * Entrypoints.Length * RepetitionCount} status={(passed ? "passed" : "failed")}");
        return passed ? 0 : 1;
    }

    public static int RunFailureProbe(string artifactPath)
    {
        var fullPath = Path.GetFullPath(artifactPath);
        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("失败路径产物路径必须包含目录。");
        Directory.CreateDirectory(directory);
        var passed = true;
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "excel-staging-entrypoint-failure-probe",
            schema = 1,
            taskId = TaskId,
            modes = new[] { "Setup", "Sampler", "Parser" },
            contract = "每种受控失败都必须写出 failed 记录、异常信息和 cleanup=true。"
        }));

        var repetition = 1;
        foreach (var failureMode in new[] { FailureMode.Setup, FailureMode.Sampler, FailureMode.Parser })
        {
            var result = RunOne("ExportToFile", 4, repetition++, failureMode);
            writer.WriteLine(JsonSerializer.Serialize(result));
            writer.Flush();
            passed &= result.Status == "failed"
                && result.CleanupSucceeded
                && !string.IsNullOrWhiteSpace(result.Exception);
        }

        Console.WriteLine($"STAGING_ENTRYPOINT_FAILURE_PROBE artifact={fullPath} scenarios=3 status={(passed ? "passed" : "failed")}");
        return passed ? 0 : 1;
    }

    private static EntrypointResult RunOne(string entrypoint, int rowCount, int repetition,
        FailureMode failureMode)
    {
        var workDirectory = Path.Combine(Path.GetTempPath(), "bing-offices-entrypoint-matrix", TaskId,
            $"{entrypoint}-r{rowCount}-n{repetition}-{Guid.NewGuid():N}");
        var outputPath = entrypoint.Contains("File", StringComparison.Ordinal)
            ? Path.Combine(workDirectory, "result.xlsx")
            : null;
        var process = Process.GetCurrentProcess();
        var stopwatch = Stopwatch.StartNew();
        var allocationBefore = 0L;
        var gen0Before = 0;
        var gen1Before = 0;
        var gen2Before = 0;
        var sampler = default(DirectorySampler);
        byte[] outputBytes = null;
        var errors = new List<string>();
        var samplingSucceeded = false;
        var fileHandleReopenApplicable = outputPath != null;
        bool? fileHandleReopenSucceeded = null;
        var parserReopenSucceeded = false;
        var cleanupSucceeded = false;
        var transientFileCount = -1;
        var transientBytesAfter = -1L;

        try
        {
            try
            {
                Directory.CreateDirectory(workDirectory);
                if (failureMode == FailureMode.Setup)
                    throw new InvalidOperationException("injected setup failure.");
                var rows = Enumerable.Range(0, rowCount).Select(index => new EntrypointRow
                {
                    Code = $"ENTRY-{index:D6}",
                    Quantity = index,
                    Description = $"staging-entrypoint-{index}"
                }).ToArray();
                var request = ExcelExport.Workbook(workbook => workbook
                    .Format(ExcelFormat.Xlsx)
                    .AddSheet("Data", rows));
                var exporter = CreateExporter(workDirectory);

                GC.Collect(2, GCCollectionMode.Forced, true, true);
                allocationBefore = GC.GetTotalAllocatedBytes(true);
                gen0Before = GC.CollectionCount(0);
                gen1Before = GC.CollectionCount(1);
                gen2Before = GC.CollectionCount(2);
                sampler = new DirectorySampler(workDirectory, outputPath,
                    failureMode == FailureMode.Sampler);
                sampler.Start();

                switch (entrypoint)
                {
                    case "Export":
                        using (var destination = new MemoryStream())
                        {
                            exporter.Export(request, destination);
                            outputBytes = destination.ToArray();
                        }
                        break;
                    case "ExportAsync":
                        using (var destination = new MemoryStream())
                        {
                            exporter.ExportAsync(request, destination).GetAwaiter().GetResult();
                            outputBytes = destination.ToArray();
                        }
                        break;
                    case "ExportToFile":
                        exporter.ExportToFile(request, outputPath);
                        outputBytes = File.ReadAllBytes(outputPath);
                        break;
                    case "ExportToFileAsync":
                        exporter.ExportToFileAsync(request, outputPath).GetAwaiter().GetResult();
                        outputBytes = File.ReadAllBytes(outputPath);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(entrypoint), entrypoint, "未知入口。");
                }

                if (failureMode == FailureMode.Parser && outputBytes != null)
                    outputBytes = outputBytes.Take(Math.Max(1, outputBytes.Length / 2)).ToArray();
            }
            catch (Exception caught)
            {
                errors.Add($"operation: {caught.GetBaseException()}");
            }

            stopwatch.Stop();
            if (sampler == null)
            {
                errors.Add("sampling: sampler was not initialized.");
            }
            else
            {
                try
                {
                    sampler.Stop();
                    samplingSucceeded = true;
                    transientFileCount = sampler.TransientFileCount;
                    transientBytesAfter = sampler.TransientBytesAfter;
                }
                catch (Exception caught)
                {
                    errors.Add($"sampling: {caught.GetBaseException()}");
                }
            }

            if (outputBytes == null || outputBytes.Length == 0)
            {
                errors.Add("output: no non-empty output was produced.");
            }
            else
            {
                if (outputPath == null)
                {
                    fileHandleReopenSucceeded = null;
                }
                else
                {
                    try
                    {
                        using var exclusive = new FileStream(outputPath, FileMode.Open, FileAccess.Read, FileShare.None);
                        fileHandleReopenSucceeded = exclusive.Length == outputBytes.LongLength && exclusive.Length > 0;
                        if (fileHandleReopenSucceeded != true)
                            errors.Add("file-reopen: exclusive file length did not match output bytes.");
                    }
                    catch (Exception caught)
                    {
                        errors.Add($"file-reopen: {caught.GetBaseException()}");
                    }
                }

                if (!TryReopenWorkbook(outputBytes, rowCount, out var parserError))
                    errors.Add($"parser-reopen: {parserError}");
                else
                    parserReopenSucceeded = true;
            }
        }
        catch (Exception caught)
        {
            errors.Add($"scenario: {caught.GetBaseException()}");
        }
        finally
        {
            stopwatch.Stop();
            cleanupSucceeded = TryDeleteDirectory(workDirectory, out var cleanupException);
            if (!cleanupSucceeded)
                errors.Add($"cleanup: {cleanupException}");
        }

        var outputHash = outputBytes == null ? null : Convert.ToHexString(SHA256.HashData(outputBytes));
        var reopenSucceeded = (!fileHandleReopenApplicable || fileHandleReopenSucceeded == true)
            && parserReopenSucceeded;
        return new EntrypointResult
        {
            TaskId = TaskId,
            Strategy = Strategy,
            Entrypoint = entrypoint,
            FailureMode = failureMode.ToString(),
            RowCount = rowCount,
            Repetition = repetition,
            ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
            AllocatedBytes = Math.Max(0, GC.GetTotalAllocatedBytes(true) - allocationBefore),
            Gen0Collections = GC.CollectionCount(0) - gen0Before,
            Gen1Collections = GC.CollectionCount(1) - gen1Before,
            Gen2Collections = GC.CollectionCount(2) - gen2Before,
            PeakWorkingSetBytes = process.PeakWorkingSet64,
            TemporaryDiskPeakBytes = sampler?.PeakBytes ?? 0,
            TemporaryFilePeakCount = sampler?.PeakFiles ?? 0,
            TemporaryDiskBytesAfter = transientBytesAfter,
            LeftoverFiles = transientFileCount,
            OutputBytes = outputBytes?.LongLength ?? 0,
            OutputSha256 = outputHash,
            FileHandleReopenApplicable = fileHandleReopenApplicable,
            FileHandleReopenSucceeded = fileHandleReopenSucceeded,
            ParserReopenSucceeded = parserReopenSucceeded,
            ReopenSucceeded = reopenSucceeded,
            SamplingSucceeded = samplingSucceeded,
            CleanupSucceeded = cleanupSucceeded,
            Status = errors.Count == 0 && outputBytes != null && outputBytes.Length > 0
                && samplingSucceeded && reopenSucceeded && transientFileCount == 0 && cleanupSucceeded
                ? "passed" : "failed",
            Exception = errors.Count == 0 ? null : string.Join(Environment.NewLine, errors)
        };
    }

    private static bool TryReopenWorkbook(byte[] outputBytes, int rowCount, out string error)
    {
        error = null;
        try
        {
            using var source = new MemoryStream(outputBytes, writable: false);
            using var workbook = WorkbookFactory.Create(source);
            var sheet = workbook.GetSheet("Data")
                ?? throw new InvalidDataException("重开工作簿缺少 Data 工作表。");
            var firstRow = sheet.GetRow(1)
                ?? throw new InvalidDataException("重开工作簿缺少首条数据行。");
            var lastRow = sheet.GetRow(rowCount)
                ?? throw new InvalidDataException("重开工作簿缺少末条数据行。");
            if (!string.Equals(firstRow.GetCell(0)?.StringCellValue, "ENTRY-000000",
                    StringComparison.Ordinal)
                || !string.Equals(lastRow.GetCell(0)?.StringCellValue, $"ENTRY-{rowCount - 1:D6}",
                    StringComparison.Ordinal))
                throw new InvalidDataException("重开工作簿的数据首尾行内容不匹配。");
            return true;
        }
        catch (Exception caught)
        {
            error = caught.GetBaseException().ToString();
            return false;
        }
    }

    private static IExcelExporter CreateExporter(string directory)
    {
        var factory = CreateStagingFactory(directory);
        var constructor = typeof(NpoiExcelExporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 2
                && item.GetParameters()[0].ParameterType == typeof(IFileExportCommitter)
                && item.GetParameters()[1].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelExporter)constructor.Invoke(new object[] { new DefaultFileExportCommitter(), factory });
    }

    private static object CreateStagingFactory(string directory)
    {
        var assembly = typeof(NpoiExcelExporter).Assembly;
        var factoryType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingFactory", true);
        var strategyType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingStrategy", true);
        var strategyValue = Enum.Parse(strategyType, Strategy);
        return Activator.CreateInstance(factoryType, BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { strategyValue, HybridThresholdBytes, directory }, CultureInfo.InvariantCulture);
    }

    private static bool TryDeleteDirectory(string directory, out Exception exception)
    {
        exception = null;
        try
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
            if (Directory.Exists(directory))
            {
                exception = new IOException("工作目录删除后仍然存在。");
                return false;
            }
            return true;
        }
        catch (Exception caught)
        {
            exception = caught;
            return false;
        }
    }

    private sealed class EntrypointRow
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
        public string Description { get; set; }
    }

    private sealed class EntrypointResult
    {
        public string TaskId { get; set; }
        public string Strategy { get; set; }
        public string Entrypoint { get; set; }
        public string FailureMode { get; set; }
        public int RowCount { get; set; }
        public int Repetition { get; set; }
        public double ElapsedMilliseconds { get; set; }
        public long AllocatedBytes { get; set; }
        public int Gen0Collections { get; set; }
        public int Gen1Collections { get; set; }
        public int Gen2Collections { get; set; }
        public long PeakWorkingSetBytes { get; set; }
        public long TemporaryDiskPeakBytes { get; set; }
        public int TemporaryFilePeakCount { get; set; }
        public long TemporaryDiskBytesAfter { get; set; }
        public int LeftoverFiles { get; set; }
        public long OutputBytes { get; set; }
        public string OutputSha256 { get; set; }
        public bool FileHandleReopenApplicable { get; set; }
        public bool? FileHandleReopenSucceeded { get; set; }
        public bool ParserReopenSucceeded { get; set; }
        public bool ReopenSucceeded { get; set; }
        public bool SamplingSucceeded { get; set; }
        public bool CleanupSucceeded { get; set; }
        public string Status { get; set; }
        public string Exception { get; set; }
    }

    private sealed class DirectorySampler
    {
        private readonly string _directory;
        private readonly string _excludedPath;
        private readonly bool _injectFailure;
        private readonly CancellationTokenSource _cancellation = new();
        private Task _task;
        private long _peakBytes;
        private int _peakFiles;

        public DirectorySampler(string directory, string excludedPath, bool injectFailure)
        {
            _directory = directory;
            _excludedPath = excludedPath == null ? null : Path.GetFullPath(excludedPath);
            _injectFailure = injectFailure;
        }

        public long PeakBytes => Interlocked.Read(ref _peakBytes);
        public int PeakFiles => Volatile.Read(ref _peakFiles);
        public int TransientFileCount { get; private set; }
        public long TransientBytesAfter { get; private set; }

        public void Start() => _task = Task.Run(async () =>
        {
            while (!_cancellation.IsCancellationRequested)
            {
                Sample();
                try
                {
                    await Task.Delay(5, _cancellation.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }
        });

        public void Stop()
        {
            Exception failure = null;
            try
            {
                _cancellation.Cancel();
                _task?.GetAwaiter().GetResult();
                Sample();
                var files = GetTransientFiles();
                TransientFileCount = files.Count;
                TransientBytesAfter = files.Sum(GetFileLength);
            }
            catch (Exception caught)
            {
                failure = caught;
            }
            finally
            {
                _cancellation.Dispose();
            }

            if (failure != null)
                throw failure;
        }

        private void Sample()
        {
            if (_injectFailure)
                throw new IOException("injected sampler failure.");
            var files = GetTransientFiles();
            var bytes = files.Sum(GetFileLength);
            UpdateMaximum(ref _peakBytes, bytes);
            UpdateMaximum(ref _peakFiles, files.Count);
        }

        private List<string> GetTransientFiles() => Directory.Exists(_directory)
            ? Directory.GetFiles(_directory, "*", SearchOption.AllDirectories)
                .Where(path => _excludedPath == null
                    || !string.Equals(Path.GetFullPath(path), _excludedPath, StringComparison.OrdinalIgnoreCase))
                .ToList()
            : new List<string>();

        private static long GetFileLength(string path)
        {
            try
            {
                return new FileInfo(path).Length;
            }
            catch
            {
                return 0;
            }
        }

        private static void UpdateMaximum(ref long target, long candidate)
        {
            while (true)
            {
                var current = Interlocked.Read(ref target);
                if (candidate <= current || Interlocked.CompareExchange(ref target, candidate, current) == current)
                    return;
            }
        }

        private static void UpdateMaximum(ref int target, int candidate)
        {
            while (true)
            {
                var current = Volatile.Read(ref target);
                if (candidate <= current || Interlocked.CompareExchange(ref target, candidate, current) == current)
                    return;
            }
        }
    }
}
