using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Styles;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

internal static class StagingResourceMatrix
{
    private const string TaskId = "BO-RC-20260907-001";
    private const long HybridThresholdBytes = 8L * 1024 * 1024;
    private const long ChildWorkingSetGuardBytes = 2L * 1024 * 1024 * 1024;
    private const int MaxActualParallelism = 1;
    private const int WarmupIterations = 1;
    private const int MeasurementIterations = 2;
    private const int DescriptionLength = 512;
    private const string DescriptionAlphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
    private static readonly string[] Strategies = { "Memory", "TempFile", "Hybrid" };
    private static readonly string[] Scenarios = { "excel-100k", "failure-double-dom", "template-image-style" };
    private static readonly int[] ConcurrencyLevels = { 1, 4, 16, 64 };

    public static int Run(string artifactPath, int rowCount, string approvedBy, string approvedAt)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));

        var fullPath = Path.GetFullPath(artifactPath);
        var directory = Path.GetDirectoryName(fullPath)
            ?? throw new InvalidOperationException("矩阵产物路径必须包含目录。");
        Directory.CreateDirectory(directory);
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "excel-staging-resource-matrix",
            schema = 1,
            taskId = TaskId,
            generatedUtc = DateTimeOffset.UtcNow,
            command = $"dotnet run --project tests/Bing.Offices.ResourceProbe/Bing.Offices.ResourceProbe.csproj -c Release -- --staging-matrix {fullPath} {rowCount} {approvedBy} {approvedAt}",
            runner = new
            {
                machine = Environment.MachineName,
                os = RuntimeInformation.OSDescription,
                architecture = RuntimeInformation.ProcessArchitecture.ToString(),
                framework = RuntimeInformation.FrameworkDescription,
                dotnet = Environment.Version.ToString(),
                processorCount = Environment.ProcessorCount,
                serverGc = GCSettings.IsServerGC,
                gcLatencyMode = GCSettings.LatencyMode.ToString()
            },
            rowCount,
            strategies = Strategies,
            scenarios = Scenarios,
            concurrency = ConcurrencyLevels,
            maxActualParallelism = MaxActualParallelism,
            serverCpuCores = 2,
            serverMemoryBytes = 4L * 1024 * 1024 * 1024,
            hybridThresholdBytes = HybridThresholdBytes,
            childWorkingSetGuardBytes = ChildWorkingSetGuardBytes,
            defaultStrategy = "TempFile",
            fallbackSemantics = "Hybrid migrates to TempFile when the threshold is exceeded; disk failures are surfaced and do not fall back to Memory.",
            approvedBy,
            approvedAt,
            approvalStatus = "RUNNING"
        }));
        writer.Flush();

        var passed = true;
        foreach (var strategy in Strategies)
        foreach (var scenario in Scenarios)
        foreach (var concurrency in ConcurrencyLevels)
        {
            var result = RunChild(fullPath, strategy, scenario, concurrency, rowCount);
            writer.WriteLine(result);
            writer.Flush();
            using var parsed = JsonDocument.Parse(result);
            if (parsed.RootElement.GetProperty("exitCode").GetInt32() != 0)
                passed = false;
        }

        var approvalGranted = passed
            && !string.IsNullOrWhiteSpace(approvedBy)
            && !string.IsNullOrWhiteSpace(approvedAt);
        writer.Flush();
        writer.Dispose();
        RewriteApprovalStatus(fullPath, approvalGranted ? "APPROVED" : "BLOCKED");
        var reportPath = Path.ChangeExtension(fullPath, ".md");
        WriteReport(reportPath, fullPath, rowCount, passed, approvedBy, approvedAt);
        Console.WriteLine($"STAGING_MATRIX artifact={fullPath} scenarios={Strategies.Length * Scenarios.Length * ConcurrencyLevels.Length} status={(passed ? "passed" : "failed")} approval={(approvalGranted ? "APPROVED" : "BLOCKED")}");
        return passed ? 0 : 1;
    }

    private static void RewriteApprovalStatus(string artifactPath, string approvalStatus)
    {
        var lines = File.ReadAllLines(artifactPath, new UTF8Encoding(false));
        if (lines.Length == 0)
            throw new InvalidOperationException("资源矩阵产物缺少头部记录。");

        using var document = JsonDocument.Parse(lines[0]);
        var header = document.RootElement.EnumerateObject()
            .ToDictionary(property => property.Name, property => property.Value.Clone(), StringComparer.Ordinal);
        header["approvalStatus"] = JsonSerializer.SerializeToElement(approvalStatus);
        lines[0] = JsonSerializer.Serialize(header);
        File.WriteAllLines(artifactPath, lines, new UTF8Encoding(false));
    }

    public static int RunScenario(string artifactPath, string strategy, string scenario, int concurrency,
        int rowCount, string stagingDirectoryArgument = null)
    {
        if (!Strategies.Contains(strategy, StringComparer.Ordinal)
            || !Scenarios.Contains(scenario, StringComparer.Ordinal)
            || !ConcurrencyLevels.Contains(concurrency)
            || rowCount < 1)
            return 2;

        var artifactDirectory = Path.GetDirectoryName(Path.GetFullPath(artifactPath))
            ?? throw new InvalidOperationException("矩阵产物路径必须包含目录。");
        var stagingDirectory = stagingDirectoryArgument ?? Path.Combine(Path.GetTempPath(),
            "bing-offices-resource-matrix", TaskId, $"{strategy}-{scenario}-c{concurrency}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(stagingDirectory);

        var stopwatch = Stopwatch.StartNew();
        var process = Process.GetCurrentProcess();
        var sampler = new ResourceSampler(stagingDirectory);
        var startedUtc = DateTimeOffset.UtcNow;
        long allocatedBefore = 0;
        var collectionBefore = new[] { 0, 0, 0 };
        var operationResults = new List<OperationResult>();
        string warmupError = null;
        Exception failure = null;
        try
        {
            var payload = CreatePayload(scenario, rowCount);
            GC.Collect(2, GCCollectionMode.Forced, true, true);
            for (var warmup = 0; warmup < WarmupIterations; warmup++)
            {
                var warmupResult = RunOperationAsync(strategy, scenario, payload, stagingDirectory)
                    .GetAwaiter().GetResult();
                warmupError ??= warmupResult.Error;
                if (warmupError != null)
                    break;
            }
            GC.Collect(2, GCCollectionMode.Forced, true, true);
            allocatedBefore = GC.GetTotalAllocatedBytes(true);
            collectionBefore[0] = GC.CollectionCount(0);
            collectionBefore[1] = GC.CollectionCount(1);
            collectionBefore[2] = GC.CollectionCount(2);
            sampler.Start();
            using var concurrencyGate = new SemaphoreSlim(MaxActualParallelism, MaxActualParallelism);
            for (var sample = 0; sample < MeasurementIterations; sample++)
            {
                var operations = Enumerable.Range(0, Math.Min(concurrency, MaxActualParallelism))
                    .Select(_ => Task.Run(() => RunOperationWithGateAsync(
                        concurrencyGate, strategy, scenario, payload, stagingDirectory)))
                    .ToArray();
                operationResults.AddRange(Task.WhenAll(operations).GetAwaiter().GetResult());
            }
        }
        catch (Exception exception)
        {
            failure = exception.GetBaseException();
        }
        finally
        {
            sampler.Stop();
            stopwatch.Stop();
        }

        var leftoverFiles = Directory.Exists(stagingDirectory)
            ? Directory.GetFiles(stagingDirectory, "*", SearchOption.AllDirectories).Length
            : -1;
        var temporaryBytesAfter = GetDirectoryBytes(stagingDirectory);
        var elapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds;
        var durations = operationResults.Select(result => result.ElapsedMilliseconds).OrderBy(value => value).ToArray();
        var result = new
        {
            kind = "staging-scenario",
            taskId = TaskId,
            artifact = artifactPath,
            startedUtc,
            strategy,
            scenario,
            rowCount,
            concurrency,
            maxActualParallelism = MaxActualParallelism,
            measuredActiveOperations = Math.Min(concurrency, MaxActualParallelism),
            queuedRequestCount = Math.Max(0, concurrency - MaxActualParallelism),
            hybridThresholdBytes = HybridThresholdBytes,
            warmupIterations = WarmupIterations,
            measurementIterations = MeasurementIterations,
            elapsedMilliseconds,
            throughputOperationsPerSecond = elapsedMilliseconds <= 0
                ? 0
                : operationResults.Count / (elapsedMilliseconds / 1000d),
            p95Milliseconds = Percentile(durations, 0.95),
            p99Milliseconds = Percentile(durations, 0.99),
            operationCount = operationResults.Count,
            operationBytes = operationResults.Sum(operation => operation.Bytes),
            errorCount = operationResults.Count(operation => operation.Error != null),
            allocatedBytes = Math.Max(0, GC.GetTotalAllocatedBytes(true) - allocatedBefore),
            gen0Collections = GC.CollectionCount(0) - collectionBefore[0],
            gen1Collections = GC.CollectionCount(1) - collectionBefore[1],
            gen2Collections = GC.CollectionCount(2) - collectionBefore[2],
            lohPeakBytes = sampler.LohPeakBytes,
            lohRetainedBytes = GetLohBytes(),
            peakWorkingSetBytes = process.PeakWorkingSet64,
            workingSetBytes = process.WorkingSet64,
            temporaryDiskPeakBytes = sampler.TemporaryDiskPeakBytes,
            temporaryDiskBytesAfter = temporaryBytesAfter,
            leftoverFiles,
            status = failure == null && warmupError == null && operationResults.All(operation => operation.Error == null)
                && leftoverFiles == 0 ? "passed" : "failed",
            exitCode = failure == null && warmupError == null
                && operationResults.All(operation => operation.Error == null)
                && leftoverFiles == 0 ? 0 : 1,
            exception = failure?.ToString() ?? warmupError,
            operationErrors = operationResults.Where(operation => operation.Error != null)
                .Select(operation => operation.Error).ToArray()
        };
        Console.WriteLine(JsonSerializer.Serialize(result));
        try
        {
            Directory.Delete(stagingDirectory, true);
        }
        catch
        {
            // The result already records the cleanup state; keep the child output parseable.
        }
        return result.exitCode;
    }

    private static string RunChild(string artifactPath, string strategy, string scenario, int concurrency,
        int rowCount)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("无法解析资源矩阵进程路径。");
        var stagingDirectory = Path.Combine(Path.GetTempPath(), "bing-offices-resource-matrix", TaskId,
            $"{strategy}-{scenario}-c{concurrency}-{Guid.NewGuid():N}");
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
        process.StartInfo.ArgumentList.Add("--staging-scenario");
        process.StartInfo.ArgumentList.Add(artifactPath);
        process.StartInfo.ArgumentList.Add(strategy);
        process.StartInfo.ArgumentList.Add(scenario);
        process.StartInfo.ArgumentList.Add(concurrency.ToString(CultureInfo.InvariantCulture));
        process.StartInfo.ArgumentList.Add(rowCount.ToString(CultureInfo.InvariantCulture));
        process.StartInfo.ArgumentList.Add(stagingDirectory);
        if (!process.Start())
            throw new InvalidOperationException("无法启动 staging 资源矩阵子进程。");
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        var resourceGuardTriggered = false;
        long observedPeakWorkingSetBytes = 0;
        while (!process.HasExited)
        {
            try
            {
                process.Refresh();
                observedPeakWorkingSetBytes = Math.Max(observedPeakWorkingSetBytes, process.WorkingSet64);
                if (process.WorkingSet64 > ChildWorkingSetGuardBytes)
                {
                    resourceGuardTriggered = true;
                    try
                    {
                        process.Kill(true);
                    }
                    catch (InvalidOperationException)
                    {
                        // The child can exit between the sample and Kill; its exit code remains authoritative.
                    }
                    break;
                }
            }
            catch (InvalidOperationException)
            {
                break;
            }
            Thread.Sleep(100);
        }
        process.WaitForExit();
        var stdout = stdoutTask.GetAwaiter().GetResult();
        var stderr = stderrTask.GetAwaiter().GetResult();
        var line = stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault()
            ?? string.Empty;
        object scenarioResult = null;
        if (resourceGuardTriggered)
        {
            var temporaryDiskBytes = GetDirectoryBytes(stagingDirectory);
            var leftoverFiles = GetDirectoryFileCount(stagingDirectory);
            var cleanupStatus = TryDeleteDirectory(stagingDirectory);
            scenarioResult = new
            {
                kind = "staging-budget-failure",
                status = "budget-failed",
                exitCode = process.ExitCode,
                resourceGuardTriggered = true,
                resourceGuardLimitBytes = ChildWorkingSetGuardBytes,
                observedPeakWorkingSetBytes,
                temporaryDiskBytesAtTermination = temporaryDiskBytes,
                leftoverFilesAtTermination = leftoverFiles,
                cleanupStatus,
                metricsAvailable = false
            };
        }
        else if (!string.IsNullOrWhiteSpace(line))
        {
            using var parsed = JsonDocument.Parse(line);
            scenarioResult = parsed.RootElement.Clone();
        }
        return JsonSerializer.Serialize(new
        {
            kind = "staging-child",
            taskId = TaskId,
            strategy,
            scenario,
            concurrency,
            rowCount,
            exitCode = process.ExitCode,
            resourceGuardTriggered,
            resourceGuardLimitBytes = ChildWorkingSetGuardBytes,
            observedPeakWorkingSetBytes,
            result = scenarioResult,
            stdout,
            stderr
        });
    }

    private static async Task<OperationResult> RunOperationAsync(string strategy, string scenario,
        Payload payload, string stagingDirectory)
    {
        var stopwatch = Stopwatch.StartNew();
        try
        {
            long bytes;
            if (scenario == "failure-double-dom")
            {
                var importer = CreateImporter(strategy, stagingDirectory);
                using var source = new MemoryStream(payload.SourceBytes, writable: false);
                using var destination = new MemoryStream();
                var request = ExcelImport.Workbook<FailureWorkbook>(builder => builder
                    .FailureWorkbook(new ExcelImportFailureOptions
                    {
                        Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                        Destination = destination
                    })
                    .Sheet("Sheet1", root => root.Rows));
                var result = await importer.ImportAsync(source, request).ConfigureAwait(false);
                bytes = destination.Length + result.Errors.Count;
            }
            else
            {
                var exporter = CreateExporter(strategy, stagingDirectory);
                using var destination = new MemoryStream();
                var request = CreateExportRequest(scenario, payload.Rows, payload.TemplateBytes);
                await exporter.ExportAsync(request, destination).ConfigureAwait(false);
                bytes = destination.Length;
            }
            stopwatch.Stop();
            return new OperationResult(stopwatch.Elapsed.TotalMilliseconds, bytes, null);
        }
        catch (Exception exception)
        {
            stopwatch.Stop();
            return new OperationResult(stopwatch.Elapsed.TotalMilliseconds, 0, exception.ToString());
        }
    }

    private static async Task<OperationResult> RunOperationWithGateAsync(SemaphoreSlim concurrencyGate,
        string strategy, string scenario, Payload payload, string stagingDirectory)
    {
        await concurrencyGate.WaitAsync().ConfigureAwait(false);
        try
        {
            return await RunOperationAsync(strategy, scenario, payload, stagingDirectory)
                .ConfigureAwait(false);
        }
        finally
        {
            concurrencyGate.Release();
        }
    }

    private static IExcelExporter CreateExporter(string strategy, string directory)
    {
        var factory = CreateStagingFactory(strategy, directory);
        var constructor = typeof(NpoiExcelExporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 2
                && item.GetParameters()[0].ParameterType == typeof(IFileExportCommitter)
                && item.GetParameters()[1].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelExporter)constructor.Invoke(new object[] { new DefaultFileExportCommitter(), factory });
    }

    private static IExcelImporter CreateImporter(string strategy, string directory)
    {
        var factory = CreateStagingFactory(strategy, directory);
        var constructor = typeof(NpoiExcelImporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 1
                && item.GetParameters()[0].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelImporter)constructor.Invoke(new[] { factory });
    }

    private static object CreateStagingFactory(string strategy, string directory)
    {
        var assembly = typeof(NpoiExcelExporter).Assembly;
        var factoryType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingFactory", true);
        var strategyType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingStrategy", true);
        var strategyValue = Enum.Parse(strategyType, strategy);
        return Activator.CreateInstance(factoryType, BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { strategyValue, HybridThresholdBytes, directory }, CultureInfo.InvariantCulture);
    }

    private static Payload CreatePayload(string scenario, int rowCount)
    {
        var rows = Enumerable.Range(0, rowCount).Select(index => new ResourceRow
        {
            Code = $"CODE-{index:D6}",
            Quantity = index,
            Amount = index + 0.25m,
            OccurredAt = new DateTime(2026, 9, 8).AddMinutes(index),
            Description = scenario == "failure-double-dom" ? null : CreateDescription(index)
        }).ToArray();
        if (scenario == "failure-double-dom")
        {
            var failureRows = rows.Select(row => new FailureRow { Code = $"BAD-{row.Quantity:D6}" }).ToArray();
            using var output = new MemoryStream();
            new NpoiExcelExporter().Export(ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", failureRows)),
                output);
            return new Payload(rows, output.ToArray(), null);
        }
        return new Payload(rows, null, scenario == "template-image-style" ? CreateTemplateBytes() : null);
    }

    private static ExcelWorkbookExportRequest CreateExportRequest(string scenario, ResourceRow[] rows,
        byte[] templateBytes)
    {
        var template = templateBytes == null ? null : new MemoryStream(templateBytes, writable: false);
        return ExcelExport.Workbook(workbook =>
        {
            if (templateBytes != null)
                workbook.UseTemplate(template, leaveOpen: false);
            workbook.AddSheet("Sheet1", rows, sheet => sheet
                .HeaderStyle(new ExcelCellStyle
                {
                    FontName = "Arial",
                    FontSize = 11,
                    Bold = true,
                    ForegroundColor = new ExcelColor("FF1F4E78"),
                    FillPattern = ExcelFillPattern.Solid
                })
                .BodyStyle(new ExcelCellStyle
                {
                    FontName = "Calibri",
                    FontSize = 10,
                    NumberFormat = "#,##0.00"
                }));
        });
    }

    private static byte[] CreateTemplateBytes()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Template");
        var pictureIndex = workbook.AddPicture(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="),
            PictureType.PNG);
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Col1 = 4;
        anchor.Row1 = 0;
        anchor.Col2 = 6;
        anchor.Row2 = 3;
        sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
        using var output = new MemoryStream();
        workbook.Write(output, false);
        return output.ToArray();
    }

    private static long GetLohBytes()
    {
        var generations = GC.GetGCMemoryInfo().GenerationInfo;
        return generations.Length > 3 ? generations[3].SizeAfterBytes : 0;
    }

    private static double Percentile(double[] values, double percentile)
    {
        if (values.Length == 0)
            return 0;
        var index = (int)Math.Ceiling(values.Length * percentile) - 1;
        return values[Math.Clamp(index, 0, values.Length - 1)];
    }

    private static long GetDirectoryBytes(string path)
    {
        if (!Directory.Exists(path))
            return 0;
        try
        {
            return Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories)
                .Select(file => new FileInfo(file).Length)
                .Sum();
        }
        catch (IOException)
        {
            return 0;
        }
        catch (UnauthorizedAccessException)
        {
            return 0;
        }
    }

    private static int GetDirectoryFileCount(string path)
    {
        if (!Directory.Exists(path))
            return 0;
        try
        {
            return Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories).Count();
        }
        catch (IOException)
        {
            return -1;
        }
        catch (UnauthorizedAccessException)
        {
            return -1;
        }
    }

    private static string TryDeleteDirectory(string path)
    {
        if (!Directory.Exists(path))
            return "not-found";
        try
        {
            Directory.Delete(path, true);
            return "deleted";
        }
        catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException)
        {
            return $"failed:{exception.GetType().Name}";
        }
    }

    private static string CreateDescription(int index)
    {
        var chars = new char[DescriptionLength];
        var state = unchecked((uint)index + 0x9E3779B9u);
        for (var position = 0; position < chars.Length; position++)
        {
            state = unchecked(state * 1664525u + 1013904223u);
            chars[position] = DescriptionAlphabet[(int)(state % DescriptionAlphabet.Length)];
        }
        return new string(chars);
    }

    private static void WriteReport(string reportPath, string artifactPath, int rowCount, bool passed,
        string approvedBy, string approvedAt)
    {
        using var writer = new StreamWriter(reportPath, false, new UTF8Encoding(false));
        writer.WriteLine($"# Excel Staging Resource Matrix");
        writer.WriteLine();
        writer.WriteLine($"- Task-ID: `{TaskId}`");
        writer.WriteLine($"- Raw artifact: `{artifactPath}`");
        writer.WriteLine($"- Runner: Windows x64 local process; each matrix cell uses an isolated child process");
        writer.WriteLine($"- Rows per workload: `{rowCount}`");
        writer.WriteLine($"- Strategies: `{string.Join("`, `", Strategies)}`");
        writer.WriteLine($"- Scenarios: `{string.Join("`, `", Scenarios)}`");
        writer.WriteLine($"- Concurrency: `{string.Join(", ", ConcurrencyLevels)}`");
        writer.WriteLine($"- Maximum actual parallelism: `{MaxActualParallelism}`; higher request concurrency is queued");
        writer.WriteLine("- Server budget: `2 CPU cores / 4 GiB RAM`");
        writer.WriteLine($"- Hybrid threshold: `{HybridThresholdBytes}` bytes");
        writer.WriteLine($"- Warmup iterations per child: `{WarmupIterations}`");
        writer.WriteLine($"- Measurement iterations per child: `{MeasurementIterations}`");
        writer.WriteLine($"- Child WorkingSet guard: `{ChildWorkingSetGuardBytes}` bytes; guarded cells are recorded as failed rather than allowed to exhaust the runner");
        writer.WriteLine($"- Child process execution: `{(passed ? "PASS" : "FAIL")}`");
        writer.WriteLine();
        writer.WriteLine("## Metrics");
        writer.WriteLine();
        writer.WriteLine("Each child warms up once, then measures the approved active slice (maximum one NPOI DOM operation) twice. PeakWorkingSet, LOH peak/retained bytes, GC collections and allocated bytes, temporary disk peak/after bytes, throughput, P95 and P99 operation latency, operation errors and cleanup residue are recorded. Logical request concurrency above the active slice is represented as queuedRequestCount; queued requests do not create additional DOM workbooks in the 2C4G profile.");
        writer.WriteLine();
        writer.WriteLine("## Strategy Contract");
        writer.WriteLine();
        writer.WriteLine("- Production default is `TempFile`, as selected by the public NPOI importer/exporter constructors.");
        writer.WriteLine("- `Hybrid` keeps writes in memory until the configured threshold, then migrates to a unique temp file while preserving the current stream position.");
        writer.WriteLine("- There is no automatic disk-to-memory fallback: temp-file creation, flush and cleanup failures remain observable failures.");
        writer.WriteLine("- Failure Workbook Async uses the same injected staging strategy and serializes directly to that staging stream before asynchronous copy to the caller destination.");
        writer.WriteLine();
        writer.WriteLine("## Approval");
        writer.WriteLine();
        writer.WriteLine("The matrix records logical request concurrency separately from the approved actual parallelism. Requests above the actual parallelism limit are queued before entering the NPOI DOM pipeline.");
        writer.WriteLine();
        writer.WriteLine("| Field | Value |");
        writer.WriteLine("| --- | --- |");
        writer.WriteLine($"| approvedBy | {approvedBy ?? ""} |");
        writer.WriteLine($"| approvedAt | {approvedAt ?? ""} |");
        writer.WriteLine($"| approvalStatus | {(passed && !string.IsNullOrWhiteSpace(approvedBy) && !string.IsNullOrWhiteSpace(approvedAt) ? "APPROVED" : "BLOCKED")} |");
    }

    private sealed class ResourceSampler
    {
        private readonly string _directory;
        private CancellationTokenSource _cancellation;
        private Task _task;

        public ResourceSampler(string directory) => _directory = directory;

        public long TemporaryDiskPeakBytes { get; private set; }
        public long LohPeakBytes { get; private set; }

        public void Start()
        {
            _cancellation = new CancellationTokenSource();
            _task = Task.Run(async () =>
            {
                while (!_cancellation.IsCancellationRequested)
                {
                    TemporaryDiskPeakBytes = Math.Max(TemporaryDiskPeakBytes, GetDirectoryBytes(_directory));
                    LohPeakBytes = Math.Max(LohPeakBytes, GetLohBytes());
                    await Task.Delay(5, _cancellation.Token).ConfigureAwait(false);
                }
            });
        }

        public void Stop()
        {
            if (_cancellation == null)
                return;
            _cancellation.Cancel();
            try
            {
                _task.GetAwaiter().GetResult();
            }
            catch (OperationCanceledException)
            {
            }
            TemporaryDiskPeakBytes = Math.Max(TemporaryDiskPeakBytes, GetDirectoryBytes(_directory));
            LohPeakBytes = Math.Max(LohPeakBytes, GetLohBytes());
            _cancellation.Dispose();
        }
    }

    private sealed class Payload
    {
        public Payload(ResourceRow[] rows, byte[] sourceBytes, byte[] templateBytes)
        {
            Rows = rows;
            SourceBytes = sourceBytes;
            TemplateBytes = templateBytes;
        }

        public ResourceRow[] Rows { get; }
        public byte[] SourceBytes { get; }
        public byte[] TemplateBytes { get; }
    }

    private sealed record OperationResult(double ElapsedMilliseconds, long Bytes, string Error);

    private sealed class ResourceRow
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
        public decimal Amount { get; set; }
        public DateTime OccurredAt { get; set; }
        public string Description { get; set; }
    }

    private sealed class FailureRow
    {
        [ExcelRegex("^OK-")]
        public string Code { get; set; }
    }

    private sealed class FailureWorkbook
    {
        public List<FailureRow> Rows { get; } = new();
    }
}
