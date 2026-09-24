using System.Diagnostics;
using System.Collections;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Xml;
using Bing.Offices.Imports;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 采集 RawDateReader 与关系绑定的 before/after 原始样本。
/// </summary>
internal static class HotspotProbe
{
    /// <summary>
    /// worker 超时时间的默认值，单位为毫秒。
    /// </summary>
    private const int DefaultWorkerTimeoutMilliseconds = 60_000;
    /// <summary>
    /// 诊断输出的最大保留长度。
    /// </summary>
    private const int MaxDiagnosticOutputLength = 4_096;
    /// <summary>
    /// 超时后等待 worker 退出的时间，单位为毫秒。
    /// </summary>
    private const int WorkerExitWaitTimeoutMilliseconds = 5_000;
    /// <summary>
    /// 读取 worker 诊断输出的超时时间，单位为毫秒。
    /// </summary>
    private const int DiagnosticReadTimeoutMilliseconds = 1_000;
    /// <summary>
    /// RawDate 工作负载使用的列数。
    /// </summary>
    private static readonly string[] RawColumnWorkloads = { "3", "10", "30" };

    /// <summary>
    /// 关系绑定工作负载使用的行数。
    /// </summary>
    private static readonly string[] RelationWorkloads = { "1000", "10000", "100000" };

    /// <summary>
    /// 获取环境变量配置的 worker 超时时间，未配置或无效时使用默认值。
    /// </summary>
    private static int WorkerTimeoutMilliseconds =>
        int.TryParse(Environment.GetEnvironmentVariable("BING_OFFICES_HOTSPOT_TIMEOUT_MS"), out var value)
            && value > 0
            ? value
            : DefaultWorkerTimeoutMilliseconds;

    /// <summary>
    /// 运行 hotspot 场景并写入原始测量样本。
    /// </summary>
    /// <param name="artifactPath">JSONL 输出路径。</param>
    /// <param name="phase">候选阶段标识，取值为 <c>before</c> 或 <c>after</c>。</param>
    /// <param name="repetitions">每个场景要求的重复次数，至少为 3。</param>
    public static async Task RunAsync(string artifactPath, string phase, int repetitions)
    {
        if (!string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(phase, "after", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("phase 必须是 before 或 after。", nameof(phase));
        if (repetitions < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitions), "性能探针至少需要三次重复。");

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var workerTimeoutMilliseconds = WorkerTimeoutMilliseconds;
        var timedOutScenarioCount = 0;
        var failedScenarioCount = 0;
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        writer.NewLine = "\n";
        writer.WriteLine(JsonSerializer.Serialize(new
        {
            kind = "hotspot-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            phase,
            repetitions,
            rawDateRows = 100000,
            rawDateColumns = RawColumnWorkloads,
            relationRows = RelationWorkloads,
            workerTimeoutMilliseconds,
            timeoutPolicy = "每个 worker 超时后终止子进程；只保留已刷盘的完整样本，不补写缺失样本",
            beforeSource = "investigative numeric-cell replay; not a HEAD Reader baseline",
            candidateIdentity = GetCandidateIdentity()
        }));

        var stopAfterFailure = false;
        foreach (var columns in RawColumnWorkloads)
        {
            var result = await RunScenarioAsync(writer, phase, repetitions, "raw-date", int.Parse(columns))
                .ConfigureAwait(false);
            timedOutScenarioCount += result.TimedOut ? 1 : 0;
            failedScenarioCount += result.Failed ? 1 : 0;
            if (result.Failed)
            {
                stopAfterFailure = true;
                break;
            }
        }
        foreach (var rows in RelationWorkloads)
        {
            if (stopAfterFailure)
                break;
            var result = await RunScenarioAsync(writer, phase, repetitions, "relation", int.Parse(rows))
                .ConfigureAwait(false);
            timedOutScenarioCount += result.TimedOut ? 1 : 0;
            failedScenarioCount += result.Failed ? 1 : 0;
            if (result.Failed)
                break;
        }

        Console.WriteLine($"HOTSPOT_PROBE artifact={fullPath} phase={phase} "
            + $"rawDateScenarios={RawColumnWorkloads.Length} relationScenarios={RelationWorkloads.Length} "
            + $"timedOutScenarios={timedOutScenarioCount} "
            + $"failedScenarios={failedScenarioCount} "
            + $"status={(failedScenarioCount > 0
                ? "failed"
                : timedOutScenarioCount == 0 ? "passed" : "not-verified")}");
        Environment.ExitCode = failedScenarioCount > 0 ? 3 : timedOutScenarioCount > 0 ? 2 : 0;
    }

    /// <summary>
    /// 在子进程模式下执行单个 hotspot 工作负载。
    /// </summary>
    /// <param name="artifactPath">worker 样本输出路径。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="workload">工作负载名称。</param>
    /// <param name="size">工作负载规模。</param>
    /// <param name="repetitions">重复次数。</param>
    public static void RunWorker(string artifactPath, string phase, string workload, int size,
        int repetitions)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(artifactPath))!);
        using var writer = new StreamWriter(artifactPath, false, new UTF8Encoding(false));
        writer.NewLine = "\n";
        if (workload == "raw-date")
        {
            var workbook = CreateWorkbook(100000, size);
            MeasureRawDate(workbook, phase, size);
            for (var repetition = 1; repetition <= repetitions; repetition++)
            {
                var sample = MeasureRawDate(workbook, phase, size);
                writer.WriteLine(JsonSerializer.Serialize(sample with { repetition = repetition }));
                writer.Flush();
            }
            return;
        }
        if (workload == "relation")
        {
            var request = CreateRelationRequest();
            BindRelations(request, phase, size);
            for (var repetition = 1; repetition <= repetitions; repetition++)
            {
                var sample = BindRelations(request, phase, size);
                writer.WriteLine(JsonSerializer.Serialize(sample with { repetition = repetition }));
                writer.Flush();
            }
            return;
        }
        throw new ArgumentException($"未知 hotspot workload: {workload}", nameof(workload));
    }

    /// <summary>
    /// 启动并汇总一个 hotspot 场景的 worker 结果。
    /// </summary>
    /// <param name="writer">场景结果写入器。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="repetitions">要求的样本数。</param>
    /// <param name="workload">工作负载名称。</param>
    /// <param name="size">工作负载规模。</param>
    /// <returns>场景是否超时或失败的状态。</returns>
    private static async Task<ScenarioResult> RunScenarioAsync(StreamWriter writer, string phase, int repetitions,
        string workload, int size)
    {
        var path = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(writer.BaseStream is FileStream file
            ? file.Name : throw new InvalidOperationException("探针输出必须是文件。")))!,
            $".{Guid.NewGuid():N}.{workload}.{size}.jsonl");
        try
        {
            var worker = await RunWorkerProcessAsync(path, phase, workload, size, repetitions)
                .ConfigureAwait(false);
            var samples = ReadSamples(path, out var discardedLineCount);
            var workerFailed = !worker.TimedOut && worker.ExitCode is not null && worker.ExitCode.Value != 0;
            if (!worker.TimedOut && !workerFailed && discardedLineCount != 0)
                throw new InvalidOperationException(
                    $"{workload}/{size} 存在无法解析的完整样本行: {discardedLineCount}");
            if (!worker.TimedOut && !workerFailed && samples.Length != repetitions)
                throw new InvalidOperationException($"{workload}/{size} 样本数不正确: {samples.Length}");
            var complete = !worker.TimedOut
                && !workerFailed
                && worker.ExitCode == 0
                && discardedLineCount == 0
                && samples.Length == repetitions;
            foreach (var sample in samples)
                writer.WriteLine(JsonSerializer.Serialize(sample));
            if (worker.TimedOut)
            {
                writer.WriteLine(JsonSerializer.Serialize(new
                {
                    kind = "hotspot-timeout",
                    workload,
                    size,
                    phase,
                    requiredRepetitions = repetitions,
                    completedRepetitions = samples.Length,
                    timeoutMilliseconds = WorkerTimeoutMilliseconds,
                    elapsedMilliseconds = worker.ElapsedMilliseconds,
                    workerExitCode = worker.ExitCode,
                    exitObserved = worker.ExitObserved,
                    killAttempted = worker.KillAttempted,
                    killSucceeded = worker.KillSucceeded,
                    killError = worker.KillError,
                    failureCategory = worker.KillSucceeded
                        ? worker.ExitObserved ? "timeout-killed" : "timeout-kill-no-exit"
                        : !worker.ExitObserved
                            ? "timeout-no-exit"
                            : worker.ExitCode != 0
                                ? "timeout-with-nonzero-exit"
                                : "timeout",
                    standardError = worker.StandardError,
                    standardOutput = worker.StandardOutput,
                    discardedPartialLineCount = discardedLineCount,
                    status = "not-verified",
                    reason = "worker-timeout; incomplete before/after comparison is not a complete result"
                }));
            }
            else if (workerFailed)
            {
                writer.WriteLine(JsonSerializer.Serialize(new
                {
                    kind = "hotspot-failure",
                    workload,
                    size,
                    phase,
                    requiredRepetitions = repetitions,
                    completedRepetitions = samples.Length,
                    workerExitCode = worker.ExitCode,
                    exitObserved = worker.ExitObserved,
                    failureCategory = "worker-failure",
                    standardError = worker.StandardError,
                    standardOutput = worker.StandardOutput,
                    discardedPartialLineCount = discardedLineCount,
                    status = "failed",
                    reason = "worker-exited-nonzero; benchmark scenario was not accepted"
                }));
            }
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "hotspot-summary",
                workload,
                size,
                phase,
                repetitions,
                requiredRepetitions = repetitions,
                completedRepetitions = samples.Length,
                status = workerFailed ? "failed" : complete ? "passed" : "not-verified",
                medianElapsedMilliseconds = !complete
                    ? (double?)null
                    : Median(samples.Select(sample => sample.elapsedMilliseconds)),
                medianAllocatedBytes = !complete
                    ? (double?)null
                    : Median(samples.Select(sample => sample.allocatedBytes)),
                medianPeakWorkingSetBytes = !complete
                    ? (double?)null
                    : Median(samples.Select(sample => sample.peakWorkingSetBytes)),
                medianIndexCount = !complete
                    ? (double?)null
                    : Median(samples.Select(sample => (long)sample.indexCount)),
                medianLohSnapshotBytes = !complete
                    ? (double?)null
                    : Median(samples.Select(sample => sample.lohSnapshotBytes))
            }));
            writer.Flush();
            return new ScenarioResult(worker.TimedOut, workerFailed);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 读取 worker 已刷盘且结构完整的样本行。
    /// </summary>
    /// <param name="path">worker 样本文件路径。</param>
    /// <param name="discardedLineCount">返回被丢弃的空或无效样本行数量。</param>
    /// <returns>可解析的 hotspot 样本。</returns>
    private static HotspotSample[] ReadSamples(string path, out int discardedLineCount)
    {
        discardedLineCount = 0;
        if (!File.Exists(path))
            return Array.Empty<HotspotSample>();

        var samples = new List<HotspotSample>();
        foreach (var line in File.ReadLines(path, Encoding.UTF8))
        {
            if (string.IsNullOrWhiteSpace(line))
                continue;
            try
            {
                using var document = JsonDocument.Parse(line);
                if (!document.RootElement.TryGetProperty("kind", out var kind)
                    || !string.Equals(kind.GetString(), "hotspot-sample", StringComparison.Ordinal))
                    continue;
                var sample = JsonSerializer.Deserialize<HotspotSample>(line);
                if (sample == null)
                    discardedLineCount++;
                else
                    samples.Add(sample);
            }
            catch (JsonException)
            {
                // 子进程被终止时，最后一行可能只写入了半条 JSON；该行不能算作样本。
                discardedLineCount++;
            }
        }
        return samples.ToArray();
    }

    /// <summary>
    /// 启动 worker 并在超时或退出时收集进程结果。
    /// </summary>
    /// <param name="artifactPath">worker 样本输出路径。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="workload">工作负载名称。</param>
    /// <param name="size">工作负载规模。</param>
    /// <param name="repetitions">要求的样本数。</param>
    /// <returns>worker 退出、超时和诊断信息。</returns>
    private static async Task<WorkerProcessResult> RunWorkerProcessAsync(string artifactPath, string phase,
        string workload,
        int size, int repetitions)
    {
        var processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("无法定位性能探针进程。");
        var process = new Process
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
            process.StartInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        process.StartInfo.ArgumentList.Add("--hotspot-worker");
        process.StartInfo.ArgumentList.Add(artifactPath);
        process.StartInfo.ArgumentList.Add(phase);
        process.StartInfo.ArgumentList.Add(workload);
        process.StartInfo.ArgumentList.Add(size.ToString());
        process.StartInfo.ArgumentList.Add(repetitions.ToString());
        using (process)
        {
            if (!process.Start())
                throw new InvalidOperationException("无法启动 hotspot worker。");
            var stopwatch = Stopwatch.StartNew();
            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var exitTask = process.WaitForExitAsync();
            var timeoutTask = Task.Delay(WorkerTimeoutMilliseconds);
            await Task.WhenAny(exitTask, timeoutTask).ConfigureAwait(false);
            var timedOut = !exitTask.IsCompleted;
            var exitObserved = false;
            var killAttempted = false;
            var killSucceeded = false;
            string? killError = null;
            if (timedOut)
            {
                killAttempted = true;
                try
                {
                    process.Kill(entireProcessTree: true);
                    killSucceeded = true;
                }
                catch (InvalidOperationException exception)
                {
                    // 进程可能刚好在超时检查与 Kill 之间退出；仍需等待退出任务完成。
                    killError = exception.Message;
                }
                catch (System.ComponentModel.Win32Exception exception)
                {
                    killError = exception.Message;
                }
            }
            if (timedOut)
            {
                var exitWaitTask = Task.Delay(WorkerExitWaitTimeoutMilliseconds);
                await Task.WhenAny(exitTask, exitWaitTask).ConfigureAwait(false);
                exitObserved = exitTask.IsCompleted;
            }
            else
            {
                await exitTask.ConfigureAwait(false);
                exitObserved = true;
            }
            stopwatch.Stop();
            var stderr = await ReadDiagnosticAsync(stderrTask).ConfigureAwait(false);
            var stdout = await ReadDiagnosticAsync(stdoutTask).ConfigureAwait(false);
            int? exitCode = exitObserved ? process.ExitCode : null;
            return new WorkerProcessResult(timedOut, exitCode, exitObserved, stopwatch.Elapsed.TotalMilliseconds,
                killAttempted, killSucceeded, killError, TruncateDiagnostic(stderr),
                TruncateDiagnostic(stdout));
        }
    }

    /// <summary>
    /// 在有限等待时间内读取进程诊断输出。
    /// </summary>
    /// <param name="readTask">异步诊断读取任务。</param>
    /// <returns>读取到的诊断文本或超时标记。</returns>
    private static async Task<string> ReadDiagnosticAsync(Task<string> readTask)
    {
        var timeoutTask = Task.Delay(DiagnosticReadTimeoutMilliseconds);
        var completedTask = await Task.WhenAny(readTask, timeoutTask).ConfigureAwait(false);
        if (completedTask != readTask)
            return "<diagnostic-read-timeout>";
        try
        {
            return await readTask.ConfigureAwait(false);
        }
        catch (Exception exception)
        {
            return $"<diagnostic-read-failed: {exception.GetType().Name}: {exception.Message}>";
        }
    }

    /// <summary>
    /// 测量指定 RawDate 工作簿的读取和索引处理。
    /// </summary>
    /// <param name="workbook">待处理的 XLSX 字节。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="columns">工作簿中的数据列数。</param>
    /// <returns>RawDate 测量样本。</returns>
    private static HotspotSample MeasureRawDate(byte[] workbook, string phase, int columns)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var lohBefore = GetLohSnapshotBytes();
        var stopwatch = Stopwatch.StartNew();
        int indexCount;
        using (var source = new MemoryStream(workbook, writable: false))
        {
            indexCount = string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase)
                ? ReadAllNumeric(source)
                : ReadCurrentDateColumns(source);
        }
        stopwatch.Stop();
        process.Refresh();
        var lohAfter = GetLohSnapshotBytes();
        return new HotspotSample("hotspot-sample", "raw-date", columns, phase, 0,
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, indexCount,
            workbook.Length, GC.CollectionCount(0) - gen0Before,
            GC.CollectionCount(1) - gen1Before, GC.CollectionCount(2) - gen2Before,
            Math.Max(lohBefore, lohAfter), process.PeakWorkingSet64, Environment.ProcessId);
    }

    /// <summary>
    /// 通过当前 RawDate reader 读取日期列索引数量。
    /// </summary>
    /// <param name="source">位于起始位置的 XLSX 输入流。</param>
    /// <returns>读取到的日期列索引数量。</returns>
    private static int ReadCurrentDateColumns(Stream source)
    {
        var type = typeof(Bing.Offices.Imports.MiniExcelExcelImporter).Assembly
            .GetType("Bing.Offices.MiniExcel.Internals.MiniExcelRawDateSerialReader")!;
        var method = type.GetMethod("Read", BindingFlags.Static | BindingFlags.NonPublic,
            binder: null, new[] { typeof(Stream), typeof(string), typeof(CancellationToken) },
            modifiers: null)
            ?? throw new MissingMethodException(type.FullName, "Read(Stream, string, CancellationToken)");
        var result = (IReadOnlyDictionary<long, double>)method.Invoke(null,
            new object[] { source, "Data", CancellationToken.None })!;
        return result.Count;
    }

    /// <summary>
    /// 扫描工作表 XML 中的全部数值单元格。
    /// </summary>
    /// <param name="source">XLSX 输入流。</param>
    /// <returns>唯一数值单元格数量。</returns>
    private static int ReadAllNumeric(Stream source)
    {
        source.Position = 0;
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        using var stream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true
        });
        var values = new Dictionary<long, double>();
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "c")
                continue;
            var reference = reader.GetAttribute("r");
            var value = ReadCellValue(reader);
            if (TryParseReference(reference, out var row, out var column)
                && double.TryParse(value, System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var serial))
                values[((long)row << 32) | (uint)column] = serial;
        }
        return values.Count;
    }

    /// <summary>
    /// 从当前单元格元素中读取数值文本。
    /// </summary>
    /// <param name="reader">定位在单元格元素上的 XML 读取器。</param>
    /// <returns>单元格的数值文本；没有数值元素时返回空字符串。</returns>
    private static string ReadCellValue(XmlReader reader)
    {
        using var subtree = reader.ReadSubtree();
        while (subtree.Read())
            if (subtree.NodeType == XmlNodeType.Element && subtree.LocalName == "v")
                return subtree.ReadElementContentAsString();
        return string.Empty;
    }

    /// <summary>
    /// 测量关系绑定工作负载。
    /// </summary>
    /// <param name="request">关系绑定导入请求。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="size">父子数据规模。</param>
    /// <returns>关系绑定测量样本。</returns>
    private static HotspotSample BindRelations(ExcelWorkbookImportRequest<RelationProbeWorkbook> request,
        string phase, int size)
    {
        var root = CreateRelationWorkbook(size);
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var lohBefore = GetLohSnapshotBytes();
        var stopwatch = Stopwatch.StartNew();
        if (string.Equals(phase, "before", StringComparison.OrdinalIgnoreCase))
        {
            var relation = request.Relations.Single();
            var parents = ((IEnumerable)relation.Parents(root)).Cast<object>().ToArray();
            var children = ((IEnumerable)relation.Children(root)).Cast<object>().ToArray();
            foreach (var child in children)
            {
                var key = relation.ChildKey.DynamicInvoke(child);
                var parent = parents.FirstOrDefault(candidate => RelationKeysEqual(
                    relation.ParentKey.DynamicInvoke(candidate), key, relation.Comparer));
                (relation.Navigation(parent) as IList)?.Add(child);
            }
        }
        else
        {
            var method = typeof(Bing.Offices.Imports.MiniExcelExcelImporter)
                .GetMethod("BindRelations", BindingFlags.Static | BindingFlags.NonPublic)!
                .MakeGenericMethod(typeof(RelationProbeWorkbook));
            var errors = new List<ExcelImportError>();
            method.Invoke(null, new object[] { root, request.Relations, errors, request, CancellationToken.None });
        }
        stopwatch.Stop();
        process.Refresh();
        var lohAfter = GetLohSnapshotBytes();
        var linked = root.Parents.Sum(parent => parent.Items.Count);
        if (linked != size)
            throw new InvalidOperationException($"关系探针绑定数量错误: {linked}/{size}");
        return new HotspotSample("hotspot-sample", "relation", size, phase, 0,
            stopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(true) - allocatedBefore, linked, 0,
            GC.CollectionCount(0) - gen0Before, GC.CollectionCount(1) - gen1Before,
            GC.CollectionCount(2) - gen2Before, Math.Max(lohBefore, lohAfter),
            process.PeakWorkingSet64, Environment.ProcessId);
    }

    /// <summary>
    /// 使用关系配置中的比较器比较父子键。
    /// </summary>
    /// <param name="left">左侧键值。</param>
    /// <param name="right">右侧键值。</param>
    /// <param name="comparer">关系配置提供的比较器。</param>
    /// <returns>键值相等时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool RelationKeysEqual(object? left, object? right, object? comparer)
    {
        if (comparer is System.Collections.IEqualityComparer nonGeneric)
            return nonGeneric.Equals(left, right);
        if (comparer is IEqualityComparer<object> objectComparer)
            return objectComparer.Equals(left, right);
        if (comparer != null)
        {
            var leftType = left?.GetType() ?? right?.GetType();
            if (leftType != null)
            {
                var equals = comparer.GetType().GetMethod(nameof(object.Equals),
                    BindingFlags.Instance | BindingFlags.Public, binder: null,
                    types: new[] { leftType, leftType }, modifiers: null);
                if (equals != null)
                    return equals.Invoke(comparer, new[] { left, right }) is true;
            }
        }
        return Equals(left, right);
    }

    /// <summary>
    /// 获取当前进程的大对象堆大小快照。
    /// </summary>
    /// <returns>大对象堆大小，单位为字节。</returns>
    private static long GetLohSnapshotBytes() =>
        GC.GetGCMemoryInfo().GenerationInfo.Length > 3
            ? GC.GetGCMemoryInfo().GenerationInfo[3].SizeAfterBytes : 0;

    /// <summary>
    /// 创建关系绑定测量使用的导入请求。
    /// </summary>
    /// <returns>包含父子关系配置的工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<RelationProbeWorkbook> CreateRelationRequest() =>
        ExcelImport.Workbook<RelationProbeWorkbook>(workbook =>
        {
            workbook.Sheet("Probe", root => root.Parents);
            workbook.HasMany(root => root.Parents, root => root.Children,
                parent => parent.Key, child => child.Key,
                parent => parent.Items, StringComparer.OrdinalIgnoreCase);
        });

    /// <summary>
    /// 创建具有大小写差异键值的关系测试工作簿。
    /// </summary>
    /// <param name="size">要生成的父子项数量。</param>
    /// <returns>生成的关系测试工作簿。</returns>
    private static RelationProbeWorkbook CreateRelationWorkbook(int size)
    {
        var root = new RelationProbeWorkbook();
        for (var index = 0; index < size; index++)
        {
            root.Parents.Add(new RelationProbeParent { Key = $"K-{index}" });
            root.Children.Add(new RelationProbeChild { Key = $"k-{index}" });
        }
        return root;
    }

    /// <summary>
    /// 创建包含数值单元格的最小 XLSX 工作簿。
    /// </summary>
    /// <param name="rows">工作表行数。</param>
    /// <param name="columns">工作表列数。</param>
    /// <returns>生成的 XLSX 字节。</returns>
    private static byte[] CreateWorkbook(int rows, int columns)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "xl/workbook.xml",
                "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>"
                + "<sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            WriteEntry(archive, "xl/_rels/workbook.xml.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>");
            var entry = archive.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false), 65536);
            writer.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>");
            for (var row = 1; row <= rows; row++)
            {
                writer.Write($"<row r=\"{row}\">");
                for (var column = 1; column <= columns; column++)
                    writer.Write($"<c r=\"{ColumnName(column)}{row}\" t=\"n\"><v>{row + column}</v></c>");
                writer.Write("</row>");
            }
            writer.Write("</sheetData></worksheet>");
        }
        return stream.ToArray();
    }

    /// <summary>
    /// 向 ZIP 工作簿写入 UTF-8 文本条目。
    /// </summary>
    /// <param name="archive">目标 ZIP 存档。</param>
    /// <param name="name">条目名称。</param>
    /// <param name="value">条目文本内容。</param>
    private static void WriteEntry(ZipArchive archive, string name, string value)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(value);
    }

    /// <summary>
    /// 将一位起始的列序号转换为 Excel 列名。
    /// </summary>
    /// <param name="column">列序号，从 1 开始。</param>
    /// <returns>对应的 Excel 列名。</returns>
    private static string ColumnName(int column)
    {
        var result = string.Empty;
        while (column > 0)
        {
            column--;
            result = (char)('A' + column % 26) + result;
            column /= 26;
        }
        return result;
    }

    /// <summary>
    /// 解析 A1 单元格引用。
    /// </summary>
    /// <param name="reference">待解析的单元格引用。</param>
    /// <param name="row">返回从 1 开始的行号。</param>
    /// <param name="column">返回从 1 开始的列号。</param>
    /// <returns>引用有效时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool TryParseReference(string? reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
            return false;
        var index = 0;
        long parsedColumn = 0;
        while (index < reference.Length && char.IsLetter(reference[index]))
        {
            parsedColumn = parsedColumn * 26 + char.ToUpperInvariant(reference[index]) - 'A' + 1;
            index++;
        }
        if (index == 0 || index >= reference.Length
            || !int.TryParse(reference.Substring(index), out row) || row <= 0)
            return false;
        column = (int)parsedColumn;
        return column > 0;
    }

    /// <summary>
    /// 获取 hotspot 探针依赖程序集的身份哈希。
    /// </summary>
    /// <returns>按文件名索引的程序集 SHA-256 哈希。</returns>
    private static IReadOnlyDictionary<string, string> GetCandidateIdentity()
    {
        var types = new[]
        {
            typeof(HotspotProbe),
            typeof(Bing.Offices.Exports.IExcelExporter),
            typeof(Bing.Offices.Providers.IExcelMappingPlanFactory),
            typeof(Bing.Offices.IO.DefaultFileExportCommitter),
            typeof(Bing.Offices.Imports.MiniExcelExcelImporter)
        };
        return types.Select(type => type.Assembly.Location)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(path => Path.GetFileName(path),
                path => Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
                    File.ReadAllBytes(path))), StringComparer.Ordinal);
    }

    /// <summary>
    /// 计算双精度样本的中位数。
    /// </summary>
    /// <param name="values">待计算的样本序列。</param>
    /// <returns>样本中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        return ordered.Length == 0 ? 0 : ordered[ordered.Length / 2];
    }

    /// <summary>
    /// 计算长整数样本的中位数。
    /// </summary>
    /// <param name="values">待计算的样本序列。</param>
    /// <returns>样本中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<long> values) => Median(values.Select(value => (double)value));

    /// <summary>
    /// 将诊断文本截断到记录允许的最大长度。
    /// </summary>
    /// <param name="value">原始诊断文本。</param>
    /// <returns>截断后的诊断文本。</returns>
    private static string TruncateDiagnostic(string value) => value.Length <= MaxDiagnosticOutputLength
        ? value
        : value.Substring(0, MaxDiagnosticOutputLength) + "...<truncated>";

    /// <summary>
    /// 记录 hotspot 场景的超时和失败状态。
    /// </summary>
    /// <param name="TimedOut">场景是否在 worker 超时时结束。</param>
    /// <param name="Failed">场景是否因 worker 非零退出而失败。</param>
    private sealed record ScenarioResult(bool TimedOut, bool Failed);

    /// <summary>
    /// 记录 hotspot worker 的进程执行结果。
    /// </summary>
    /// <param name="TimedOut">worker 是否超时。</param>
    /// <param name="ExitCode">worker 退出码；未观察到退出时为 <see langword="null"/>。</param>
    /// <param name="ExitObserved">是否观察到 worker 退出。</param>
    /// <param name="ElapsedMilliseconds">worker 运行耗时，单位为毫秒。</param>
    /// <param name="KillAttempted">是否尝试终止超时 worker。</param>
    /// <param name="KillSucceeded">终止操作是否成功。</param>
    /// <param name="KillError">终止失败消息；没有失败时为 <see langword="null"/>。</param>
    /// <param name="StandardError">worker 标准错误输出。</param>
    /// <param name="StandardOutput">worker 标准输出。</param>
    private sealed record WorkerProcessResult(bool TimedOut, int? ExitCode, bool ExitObserved,
        double ElapsedMilliseconds, bool KillAttempted, bool KillSucceeded, string? KillError, string StandardError,
        string StandardOutput);

    /// <summary>
    /// 记录一次 hotspot 工作负载样本。
    /// </summary>
    /// <param name="kind">样本记录类型。</param>
    /// <param name="workload">工作负载名称。</param>
    /// <param name="size">工作负载规模。</param>
    /// <param name="phase">候选阶段标识。</param>
    /// <param name="repetition">测量重复序号。</param>
    /// <param name="elapsedMilliseconds">耗时，单位为毫秒。</param>
    /// <param name="allocatedBytes">托管堆分配字节数。</param>
    /// <param name="indexCount">生成或读取的索引数量。</param>
    /// <param name="workbookBytes">输入工作簿字节数。</param>
    /// <param name="gen0Collections">第 0 代 GC 次数。</param>
    /// <param name="gen1Collections">第 1 代 GC 次数。</param>
    /// <param name="gen2Collections">第 2 代 GC 次数。</param>
    /// <param name="lohSnapshotBytes">大对象堆快照字节数。</param>
    /// <param name="peakWorkingSetBytes">进程峰值工作集字节数。</param>
    /// <param name="processId">采样进程标识。</param>
    private sealed record HotspotSample(string kind, string workload, int size, string phase, int repetition,
        double elapsedMilliseconds, long allocatedBytes, int indexCount, int workbookBytes,
        int gen0Collections, int gen1Collections, int gen2Collections, long lohSnapshotBytes,
        long peakWorkingSetBytes, int processId);

    /// <summary>
    /// 关系绑定探针的工作簿模型。
    /// </summary>
    private sealed class RelationProbeWorkbook
    {
        /// <summary>
        /// 获取或设置父项集合。
        /// </summary>
        public List<RelationProbeParent> Parents { get; } = new();

        /// <summary>
        /// 获取或设置子项集合。
        /// </summary>
        public List<RelationProbeChild> Children { get; } = new();
    }

    /// <summary>
    /// 关系绑定探针的父项模型。
    /// </summary>
    private sealed class RelationProbeParent
    {
        /// <summary>
        /// 获取或设置父项键。
        /// </summary>
        public string Key { get; set; } = string.Empty;

        /// <summary>
        /// 获取已绑定的子项集合。
        /// </summary>
        public List<RelationProbeChild> Items { get; } = new();
    }

    /// <summary>
    /// 关系绑定探针的子项模型。
    /// </summary>
    private sealed class RelationProbeChild
    {
        /// <summary>
        /// 获取或设置子项键。
        /// </summary>
        public string Key { get; set; } = string.Empty;
    }
}
