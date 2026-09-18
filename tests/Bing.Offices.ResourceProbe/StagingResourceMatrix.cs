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

/// <summary>
/// 提供测试场景使用的资源矩阵。
/// </summary>
internal static class StagingResourceMatrix
{
    /// <summary>
    /// 资源矩阵任务标识。
    /// </summary>
    private const string TaskId = "BO-RC-20260908-002";
    /// <summary>
    /// 混合 staging 迁移到临时文件的字节阈值。
    /// </summary>
    private const long HybridThresholdBytes = 8L * 1024 * 1024;
    /// <summary>
    /// 子进程工作集允许的字节上限。
    /// </summary>
    private const long ChildWorkingSetGuardBytes = 2L * 1024 * 1024 * 1024;
    /// <summary>
    /// 资源矩阵实际并发操作上限。
    /// </summary>
    private const int MaxActualParallelism = 1;
    /// <summary>
    /// 每个场景的预热迭代次数。
    /// </summary>
    private const int WarmupIterations = 1;
    /// <summary>
    /// 每个场景的测量迭代次数。
    /// </summary>
    private const int MeasurementIterations = 2;
    /// <summary>
    /// 生成测试描述文本的目标长度（字符数）。
    /// </summary>
    private const int DescriptionLength = 512;
    /// <summary>
    /// 生成测试描述文本时使用的字符集合。
    /// </summary>
    private const string DescriptionAlphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
    /// <summary>
    /// 资源矩阵覆盖的 staging 策略名称集合。
    /// </summary>
    private static readonly string[] Strategies = { "Memory", "TempFile", "Hybrid" };
    /// <summary>
    /// 资源矩阵覆盖的测试场景名称集合。
    /// </summary>
    private static readonly string[] Scenarios = { "excel-100k", "failure-double-dom", "template-image-style" };
    /// <summary>
    /// 资源矩阵覆盖的请求并发级别集合。
    /// </summary>
    private static readonly int[] ConcurrencyLevels = { 1, 4, 16, 64 };

    /// <summary>
    /// 运行。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <param name="approvedBy">审批人标识。</param>
    /// <param name="approvedAt">审批时间文本。</param>
    /// <returns>计算得到的数值。</returns>
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

    /// <summary>
    /// 重写资源矩阵审批状态。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <param name="approvalStatus">要写入的审批状态。</param>
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

    /// <summary>
    /// 运行指定资源矩阵场景。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="concurrency">并发请求数量。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <param name="stagingDirectoryArgument">目标目录路径。</param>
    /// <returns>计算得到的数值。</returns>
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
        var requestMetrics = new RequestExecutionMetrics();
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
            using var concurrencyGate = new SemaphoreSlim(1, 1);
            for (var sample = 0; sample < MeasurementIterations; sample++)
            {
                var batch = new RequestBatch(concurrency, requestMetrics);
                var operations = Enumerable.Range(0, concurrency)
                    .Select(_ => RunOperationWithGateAsync(
                        concurrencyGate, batch, strategy, scenario, payload, stagingDirectory))
                    .ToArray();
                batch.WaitUntilReady().GetAwaiter().GetResult();
                batch.Start();
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
            requestedConcurrency = concurrency,
            maxActualParallelism = MaxActualParallelism,
            actualParallelism = requestMetrics.MaxActiveOperations,
            measuredActiveOperations = requestMetrics.MaxActiveOperations,
            submittedRequests = requestMetrics.SubmittedRequests,
            completedRequests = requestMetrics.CompletedRequests,
            maxActiveOperations = requestMetrics.MaxActiveOperations,
            maxQueuedRequests = requestMetrics.MaxQueuedRequests,
            queuedRequestCount = requestMetrics.MaxQueuedRequests,
            queueEntryEvents = requestMetrics.QueueEntryEvents,
            queueExitEvents = requestMetrics.QueueExitEvents,
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

    /// <summary>
    /// 运行单个资源矩阵子进程。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="concurrency">并发请求数量。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <returns>包含子进程结果、退出码、资源保护状态和标准输出的 JSON 记录。</returns>
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
            requestedConcurrency = concurrency,
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

    /// <summary>
    /// 异步运行单次资源操作。
    /// </summary>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="payload">负载数据。</param>
    /// <param name="stagingDirectory">目标目录路径。</param>
    /// <returns>包含操作耗时、输出量和异常文本的异步结果。</returns>
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

    /// <summary>
    /// 在并发门控下异步运行资源操作。
    /// </summary>
    /// <param name="concurrencyGate">并发控制信号。</param>
    /// <param name="batch">批次信息。</param>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="payload">负载数据。</param>
    /// <param name="stagingDirectory">目标目录路径。</param>
    /// <returns>获得并发许可后执行操作产生的异步结果。</returns>
    private static async Task<OperationResult> RunOperationWithGateAsync(SemaphoreSlim concurrencyGate,
        RequestBatch batch, string strategy, string scenario, Payload payload, string stagingDirectory)
    {
        batch.Metrics.RecordSubmitted();
        batch.SignalReady();
        await batch.StartSignal.ConfigureAwait(false);

        var waitTask = concurrencyGate.WaitAsync();
        var enteredQueue = !waitTask.IsCompletedSuccessfully;
        if (enteredQueue)
            batch.Metrics.RecordQueueEntry();
        batch.SignalGateWaitStarted();

        if (enteredQueue)
        {
            try
            {
                await waitTask.ConfigureAwait(false);
            }
            finally
            {
                batch.Metrics.RecordQueueExit();
            }
        }

        await batch.AllGateWaitsStarted.ConfigureAwait(false);
        batch.Metrics.RecordActiveStarted();
        try
        {
            return await RunOperationAsync(strategy, scenario, payload, stagingDirectory)
                .ConfigureAwait(false);
        }
        finally
        {
            batch.Metrics.RecordActiveCompleted();
            batch.Metrics.RecordCompleted();
            concurrencyGate.Release();
        }
    }

    /// <summary>
    /// 表示资源探测场景的一批请求。
    /// </summary>
    private sealed class RequestBatch
    {
        /// <summary>
        /// 请求批次中所有请求已准备就绪的信号。
        /// </summary>
        private readonly TaskCompletionSource<bool> _ready = new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 启动请求批次执行的信号。
        /// </summary>
        private readonly TaskCompletionSource<bool> _start = new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 请求批次中所有并发门控等待已开始的信号。
        /// </summary>
        private readonly TaskCompletionSource<bool> _allGateWaitsStarted =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        /// <summary>
        /// 请求批次要求的并发请求数。
        /// </summary>
        private readonly int _requestedConcurrency;
        /// <summary>
        /// 已准备就绪的请求数。
        /// </summary>
        private int _readyCount;
        /// <summary>
        /// 已开始等待并发门控的请求数。
        /// </summary>
        private int _gateWaitStartedCount;

        /// <summary>
        /// 初始化一个 <see cref="RequestBatch" /> 类型的实例。
        /// </summary>
        /// <param name="requestedConcurrency">本批请求并发数量。</param>
        /// <param name="metrics">本批共享的请求计数与并发指标。</param>
        public RequestBatch(int requestedConcurrency, RequestExecutionMetrics metrics)
        {
            _requestedConcurrency = requestedConcurrency;
            Metrics = metrics;
        }

        /// <summary>
        /// 获取指标集合。
        /// </summary>
        public RequestExecutionMetrics Metrics { get; }
        /// <summary>
        /// 获取开始信号。
        /// </summary>
        public Task StartSignal => _start.Task;
        /// <summary>
        /// 获取所有门控等待是否已开始。
        /// </summary>
        public Task AllGateWaitsStarted => _allGateWaitsStarted.Task;

        /// <summary>
        /// 等待请求批次准备就绪。
        /// </summary>
        public Task WaitUntilReady() => _ready.Task;

        /// <summary>
        /// 开始资源采样或请求批次。
        /// </summary>
        public void Start() => _start.TrySetResult(true);

        /// <summary>
        /// 标记请求批次已准备就绪。
        /// </summary>
        public void SignalReady()
        {
            if (Interlocked.Increment(ref _readyCount) == _requestedConcurrency)
                _ready.TrySetResult(true);
        }

        /// <summary>
        /// 标记并发门控已开始等待。
        /// </summary>
        public void SignalGateWaitStarted()
        {
            if (Interlocked.Increment(ref _gateWaitStartedCount) == _requestedConcurrency)
                _allGateWaitsStarted.TrySetResult(true);
        }
    }

    /// <summary>
    /// 记录测试场景的资源或性能指标。
    /// </summary>
    private sealed class RequestExecutionMetrics
    {
        /// <summary>
        /// 已提交请求的计数。
        /// </summary>
        private int _submittedRequests;
        /// <summary>
        /// 已完成请求的计数。
        /// </summary>
        private int _completedRequests;
        /// <summary>
        /// 当前活动操作数。
        /// </summary>
        private int _activeOperations;
        /// <summary>
        /// 采样到的活动操作数峰值。
        /// </summary>
        private int _maxActiveOperations;
        /// <summary>
        /// 当前等待执行的请求数。
        /// </summary>
        private int _queuedRequests;
        /// <summary>
        /// 采样到的排队请求数峰值。
        /// </summary>
        private int _maxQueuedRequests;
        /// <summary>
        /// 请求进入队列的事件计数。
        /// </summary>
        private int _queueEntryEvents;
        /// <summary>
        /// 请求离开队列的事件计数。
        /// </summary>
        private int _queueExitEvents;

        /// <summary>
        /// 获取已提交请求数。
        /// </summary>
        public int SubmittedRequests => Volatile.Read(ref _submittedRequests);
        /// <summary>
        /// 获取已完成请求数。
        /// </summary>
        public int CompletedRequests => Volatile.Read(ref _completedRequests);
        /// <summary>
        /// 获取最大活动操作数。
        /// </summary>
        public int MaxActiveOperations => Volatile.Read(ref _maxActiveOperations);
        /// <summary>
        /// 获取最大排队请求数。
        /// </summary>
        public int MaxQueuedRequests => Volatile.Read(ref _maxQueuedRequests);
        /// <summary>
        /// 获取队列进入事件数。
        /// </summary>
        public int QueueEntryEvents => Volatile.Read(ref _queueEntryEvents);
        /// <summary>
        /// 获取队列退出事件数。
        /// </summary>
        public int QueueExitEvents => Volatile.Read(ref _queueExitEvents);

        /// <summary>
        /// 记录请求已提交。
        /// </summary>
        public void RecordSubmitted() => Interlocked.Increment(ref _submittedRequests);

        /// <summary>
        /// 记录请求已完成。
        /// </summary>
        public void RecordCompleted() => Interlocked.Increment(ref _completedRequests);

        /// <summary>
        /// 记录请求进入队列。
        /// </summary>
        public void RecordQueueEntry()
        {
            Interlocked.Increment(ref _queueEntryEvents);
            var queued = Interlocked.Increment(ref _queuedRequests);
            UpdateMaximum(ref _maxQueuedRequests, queued);
        }

        /// <summary>
        /// 记录请求离开队列。
        /// </summary>
        public void RecordQueueExit()
        {
            Interlocked.Decrement(ref _queuedRequests);
            Interlocked.Increment(ref _queueExitEvents);
        }

        /// <summary>
        /// 记录活动请求开始执行。
        /// </summary>
        public void RecordActiveStarted()
        {
            var active = Interlocked.Increment(ref _activeOperations);
            UpdateMaximum(ref _maxActiveOperations, active);
        }

        /// <summary>
        /// 记录活动请求执行完成。
        /// </summary>
        public void RecordActiveCompleted() => Interlocked.Decrement(ref _activeOperations);

        /// <summary>
        /// 更新采样值的最大值。
        /// </summary>
        /// <param name="target">目标对象。</param>
        /// <param name="candidate">候选项。</param>
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

    /// <summary>
    /// 创建导出器。
    /// </summary>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="directory">目标目录路径。</param>
    /// <returns>配置了测试暂存工厂的 NPOI 导出器。</returns>
    private static IExcelExporter CreateExporter(string strategy, string directory)
    {
        var factory = CreateStagingFactory(strategy, directory);
        var constructor = typeof(NpoiExcelExporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 2
                && item.GetParameters()[0].ParameterType == typeof(IFileExportCommitter)
                && item.GetParameters()[1].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelExporter)constructor.Invoke(new object[] { new DefaultFileExportCommitter(), factory });
    }

    /// <summary>
    /// 创建导入器。
    /// </summary>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="directory">目标目录路径。</param>
    /// <returns>配置了测试暂存工厂的 NPOI 导入器。</returns>
    private static IExcelImporter CreateImporter(string strategy, string directory)
    {
        var factory = CreateStagingFactory(strategy, directory);
        var constructor = typeof(NpoiExcelImporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 1
                && item.GetParameters()[0].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelImporter)constructor.Invoke(new[] { factory });
    }

    /// <summary>
    /// 创建暂存工厂实例。
    /// </summary>
    /// <param name="strategy">资源暂存策略。</param>
    /// <param name="directory">目标目录路径。</param>
    /// <returns>使用指定暂存策略、阈值和目录的 NPOI 暂存工厂实例。</returns>
    private static object CreateStagingFactory(string strategy, string directory)
    {
        var assembly = typeof(NpoiExcelExporter).Assembly;
        var factoryType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingFactory", true);
        var strategyType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingStrategy", true);
        var strategyValue = Enum.Parse(strategyType, strategy);
        return Activator.CreateInstance(factoryType, BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { strategyValue, HybridThresholdBytes, directory }, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 创建资源矩阵操作负载。
    /// </summary>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <returns>与场景对应的数据行、输入文件字节和可选模板。</returns>
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

    /// <summary>
    /// 创建 Excel 导出请求。
    /// </summary>
    /// <param name="scenario">资源矩阵场景。</param>
    /// <param name="rows">数据行或行集合。</param>
    /// <param name="templateBytes">待处理的字节内容。</param>
    /// <returns>包含样式设置和可选模板的 Sheet1 导出请求。</returns>
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

    /// <summary>
    /// 创建模板文件字节内容。
    /// </summary>
    /// <returns>包含 Sheet1 工作表和 PNG 图片的 XLSX 模板字节。</returns>
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

    /// <summary>
    /// 获取大对象堆占用字节数。
    /// </summary>
    /// <returns>最近一次 GC 后大对象堆的字节数；没有对应代信息时返回 0。</returns>
    private static long GetLohBytes()
    {
        var generations = GC.GetGCMemoryInfo().GenerationInfo;
        return generations.Length > 3 ? generations[3].SizeAfterBytes : 0;
    }

    /// <summary>
    /// 计算样本的百分位数。
    /// </summary>
    /// <param name="values">值集合。</param>
    /// <param name="percentile">百分位值。</param>
    /// <returns>排序样本中对应分位的值；空数组返回 0，索引超界时取边界值。</returns>
    private static double Percentile(double[] values, double percentile)
    {
        if (values.Length == 0)
            return 0;
        var index = (int)Math.Ceiling(values.Length * percentile) - 1;
        return values[Math.Clamp(index, 0, values.Length - 1)];
    }

    /// <summary>
    /// 计算目录占用的字节数。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>目录内文件的总字节数；目录不存在或发生 IO、权限异常时返回 0。</returns>
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

    /// <summary>
    /// 统计目录中的文件数量。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>目录内文件数量；目录不存在时返回 0，发生 IO 或权限异常时返回 -1。</returns>
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

    /// <summary>
    /// 尝试删除暂存目录。
    /// </summary>
    /// <param name="path">目标文件或目录路径。</param>
    /// <returns>目录不存在时返回 not-found，删除成功时返回 deleted，失败时返回含异常类型的状态文本。</returns>
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

    /// <summary>
    /// 创建资源矩阵场景描述。
    /// </summary>
    /// <param name="index">代码围栏或集合中的索引。</param>
    /// <returns>由行索引确定的固定长度伪随机描述文本。</returns>
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

    /// <summary>
    /// 写入资源矩阵报告。
    /// </summary>
    /// <param name="reportPath">目标文件或目录路径。</param>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <param name="passed">执行是否通过。</param>
    /// <param name="approvedBy">审批人标识。</param>
    /// <param name="approvedAt">审批时间文本。</param>
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
        writer.WriteLine("Each child warms up once, then runs two measurement batches. Every batch creates exactly the requested number of lightweight request tasks; submittedRequests, completedRequests, actualParallelism, maxQueuedRequests, queueEntryEvents and queueExitEvents are recorded from the shared SemaphoreSlim(1,1) events. Workbook construction starts only after a request owns the single active slot. PeakWorkingSet, LOH peak/retained bytes, GC collections and allocated bytes, temporary disk peak/after bytes, throughput, P95 and P99 operation latency, operation errors and cleanup residue are also recorded. queuedRequestCount remains as a compatibility alias for the measured maximum queue depth.");
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

    /// <summary>
    /// 提供资源探测场景使用的采样器。
    /// </summary>
    private sealed class ResourceSampler
    {
        /// <summary>
        /// 资源采样器监控的目录路径。
        /// </summary>
        private readonly string _directory;
        /// <summary>
        /// 停止资源采样的取消源。
        /// </summary>
        private CancellationTokenSource _cancellation;
        /// <summary>
        /// 资源采样后台任务。
        /// </summary>
        private Task _task;

        /// <summary>
        /// 初始化一个 <see cref="ResourceSampler" /> 类型的实例。
        /// </summary>
        /// <param name="directory">待采样的暂存目录。</param>
        public ResourceSampler(string directory) => _directory = directory;

        /// <summary>
        /// 获取或设置临时磁盘峰值字节数。
        /// </summary>
        public long TemporaryDiskPeakBytes { get; private set; }
        /// <summary>
        /// 获取或设置LOH 峰值字节数。
        /// </summary>
        public long LohPeakBytes { get; private set; }

        /// <summary>
        /// 开始资源采样或请求批次。
        /// </summary>
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

        /// <summary>
        /// 停止资源采样。
        /// </summary>
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

    /// <summary>
    /// 表示资源探测场景使用的请求负载。
    /// </summary>
    private sealed class Payload
    {
        /// <summary>
        /// 初始化一个 <see cref="Payload" /> 类型的实例。
        /// </summary>
        /// <param name="rows">导出场景使用的数据行。</param>
        /// <param name="sourceBytes">失败工作簿场景的输入字节；其他场景为 null。</param>
        /// <param name="templateBytes">模板场景使用的工作簿字节；其他场景为 null。</param>
        public Payload(ResourceRow[] rows, byte[] sourceBytes, byte[] templateBytes)
        {
            Rows = rows;
            SourceBytes = sourceBytes;
            TemplateBytes = templateBytes;
        }

        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public ResourceRow[] Rows { get; }
        /// <summary>
        /// 获取源数据字节数。
        /// </summary>
        public byte[] SourceBytes { get; }
        /// <summary>
        /// 获取模板数据字节数。
        /// </summary>
        public byte[] TemplateBytes { get; }
    }

    /// <summary>
    /// 表示测试场景的结果数据模型。
    /// </summary>
    /// <param name="ElapsedMilliseconds">操作耗时，单位为毫秒。</param>
    /// <param name="Bytes">输出量指标；失败工作簿场景为输出字节数与错误数量之和。</param>
    /// <param name="Error">操作异常文本；成功时为 null。</param>
    private sealed record OperationResult(double ElapsedMilliseconds, long Bytes, string Error);

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ResourceRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
        /// <summary>
        /// 获取或设置发生时间。
        /// </summary>
        public DateTime OccurredAt { get; set; }
        /// <summary>
        /// 获取或设置描述。
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class FailureRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        [ExcelRegex("^OK-")]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class FailureWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<FailureRow> Rows { get; } = new();
    }
}
