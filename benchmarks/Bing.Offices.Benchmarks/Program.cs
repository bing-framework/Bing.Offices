using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 性能基准程序入口。
/// </summary>
public static class Program
{
    /// <summary>
    /// 运行基准测试或命令行探针。
    /// </summary>
    /// <remarks>
    /// 未指定探针参数时启动 BenchmarkDotNet；指定探针参数时执行对应的隔离采集流程。
    /// </remarks>
    /// <param name="args">BenchmarkDotNet 命令行参数。</param>
    public static void Main(string[] args)
    {
        if (args.Length >= 2 && string.Equals(args[0], "--resource-probe", StringComparison.OrdinalIgnoreCase))
        {
            ResourceProbe.Run(args[1]);
            return;
        }
        if (args.Length >= 6 && string.Equals(args[0], "--resource-scenario", StringComparison.OrdinalIgnoreCase))
        {
            ResourceProbe.RunScenario(args[1], int.Parse(args[2]), int.Parse(args[3]), int.Parse(args[4]),
                int.Parse(args[5]));
            return;
        }
        if (args.Length >= 3 && string.Equals(args[0], "--tail-latency", StringComparison.OrdinalIgnoreCase))
        {
            TailLatency.Run(args[1], int.Parse(args[2]));
            return;
        }
        if (args.Length >= 3 && string.Equals(args[0], "--real-io-probe", StringComparison.OrdinalIgnoreCase))
        {
            RealIoProbe.Run(args[1], int.Parse(args[2]));
            return;
        }
        if (args.Length >= 5 && string.Equals(args[0], "--real-io-scenario", StringComparison.OrdinalIgnoreCase))
        {
            RealIoProbe.RunScenario(args[1], args[2], int.Parse(args[3]), int.Parse(args[4]));
            return;
        }
        if (args.Length >= 3 && string.Equals(args[0], "--mini-excel-probe", StringComparison.OrdinalIgnoreCase))
        {
            MiniExcelProbe.Run(args[1], int.Parse(args[2]));
            return;
        }
        if (args.Length >= 4 && string.Equals(args[0], "--entity-probe", StringComparison.OrdinalIgnoreCase))
        {
            var phase = args.Length >= 5
                ? args[4]
                : Environment.GetEnvironmentVariable("BING_OFFICES_BENCHMARK_PHASE") ?? "after";
            EntityProbe.RunAsync(args[1], int.Parse(args[2]), int.Parse(args[3]), phase)
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 4 && string.Equals(args[0], "--provider-comparison-probe",
                StringComparison.OrdinalIgnoreCase))
        {
            var phase = args.Length >= 5
                ? args[4]
                : Environment.GetEnvironmentVariable("BING_OFFICES_BENCHMARK_PHASE") ?? "after";
            string? provider = args.Length >= 6 ? args[5] : null;
            ProviderComparisonProbe.RunAsync(args[1], int.Parse(args[2]), int.Parse(args[3]), phase, provider)
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 5 && string.Equals(args[0], "--closedxml-scenario-probe",
                StringComparison.OrdinalIgnoreCase))
        {
            ClosedXmlScenarioProbe.RunAsync(args[1], int.Parse(args[2]), int.Parse(args[3]), args[4])
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 6 && string.Equals(args[0], "--provider-comparison-worker",
                StringComparison.OrdinalIgnoreCase))
        {
            ProviderComparisonProbe.RunWorkerAsync(args[1], int.Parse(args[2]), int.Parse(args[3]),
                    args[4], args[5])
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 4 && string.Equals(args[0], "--materialization-binding-probe",
                StringComparison.OrdinalIgnoreCase))
        {
            MaterializationBindingProbe.RunAsync(args[1], int.Parse(args[2]), int.Parse(args[3]))
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 4 && string.Equals(args[0], "--hotspot-probe",
                StringComparison.OrdinalIgnoreCase))
        {
            HotspotProbe.RunAsync(args[1], args[2], int.Parse(args[3]))
                .GetAwaiter().GetResult();
            return;
        }
        if (args.Length >= 6 && string.Equals(args[0], "--hotspot-worker",
                StringComparison.OrdinalIgnoreCase))
        {
            HotspotProbe.RunWorker(args[1], args[2], args[3], int.Parse(args[4]), int.Parse(args[5]));
            return;
        }
        var benchmarkArtifacts = Path.Combine(FindRepositoryRoot(), "artifacts", "benchmarks");
        var benchmarkConfig = ManualConfig.Create(DefaultConfig.Instance)
            .WithArtifactsPath(benchmarkArtifacts);
        BenchmarkSwitcher.FromTypes(
            new[]
            {
                typeof(StreamPipelineBenchmarks),
                typeof(CsvPipelineBenchmarks),
                typeof(RealIoBenchmarks),
                typeof(RealIoPipelineBenchmarks),
                typeof(MiniExcelRealIoBenchmarks),
                typeof(ProviderComparisonBenchmarks),
                typeof(GenericSheetDispatchBenchmarks),
                typeof(FailureWorkbookBenchmarks),
                typeof(HeaderStyleBenchmarks),
                typeof(ValidationRangeBenchmarks),
                typeof(MappingValidationBenchmarks),
                typeof(PropertyAccessorBenchmarks),
                typeof(DynamicPlanBenchmarks),
                typeof(TenantPlanCacheBenchmarks),
                typeof(RegexCacheBenchmarks),
                typeof(UniqueJournalBenchmarks)
            }).Run(args, benchmarkConfig);
    }

    /// <summary>
    /// 从当前目录向上查找仓库根目录。
    /// </summary>
    /// <returns>包含解决方案文件的仓库根目录。</returns>
    private static string FindRepositoryRoot()
    {
        for (var directory = new DirectoryInfo(Directory.GetCurrentDirectory()); directory != null;
             directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Bing.Offices.sln")))
                return directory.FullName;
        }
        throw new InvalidOperationException("无法定位 Bing.Offices 仓库根目录。");
    }

    /// <summary>
    /// 运行资源限制和峰值内存探针。
    /// </summary>
    private static class ResourceProbe
    {
        /// <summary>
        /// 资源探针允许的大对象堆上限，单位为字节。
        /// </summary>
        private const long LohCeilingBytes = 512L * 1024 * 1024;

        /// <summary>
        /// 资源探针允许的进程峰值工作集上限，单位为字节。
        /// </summary>
        private const long PeakWorkingSetCeilingBytes = 1024L * 1024 * 1024;

        /// <summary>
        /// 执行全部资源矩阵场景并写入结果。
        /// </summary>
        /// <param name="artifactPath">JSONL 输出路径。</param>
        public static void Run(string artifactPath)
        {
            var fullPath = Path.GetFullPath(artifactPath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "resource-probe",
                schema = 1,
                startedUtc = DateTimeOffset.UtcNow,
                process = Environment.ProcessPath,
                dotnet = Environment.Version.ToString(),
                lohCeilingBytes = LohCeilingBytes,
                peakWorkingSetCeilingBytes = PeakWorkingSetCeilingBytes,
                lohSampling = "lohSampledPeakBytes is the maximum GenerationInfo[3].SizeBeforeBytes sampled after each workload phase; lohRetainedBytes includes the live payload after forced GC."
            }));
            foreach (var planBuildCount in new[] { 100, 500 })
                foreach (var tenantCount in new[] { 100, 1000 })
                    foreach (var uniqueColumnCount in new[] { 1, 5 })
                        foreach (var uniqueRowCount in new[] { 10000, 100000 })
                        {
                            var result = RunChild(fullPath, planBuildCount, tenantCount, uniqueColumnCount, uniqueRowCount);
                            writer.WriteLine(result);
                            writer.Flush();
                            using var parsed = JsonDocument.Parse(result);
                            if (parsed.RootElement.GetProperty("exitCode").GetInt32() != 0)
                                throw new InvalidOperationException($"资源场景执行失败: {result}");
                        }
            Console.WriteLine($"RESOURCE_PROBE artifact={fullPath} scenarios=16 status=passed");
        }

        /// <summary>
        /// 执行单个映射计划和唯一值跟踪资源场景。
        /// </summary>
        /// <param name="artifactPath">结果所属的资源探针路径。</param>
        /// <param name="planBuildCount">额外构建映射计划的次数。</param>
        /// <param name="tenantCount">租户配置数量。</param>
        /// <param name="uniqueColumnCount">唯一值跟踪的列数。</param>
        /// <param name="uniqueRowCount">唯一值跟踪的行数。</param>
        public static void RunScenario(string artifactPath, int planBuildCount, int tenantCount,
            int uniqueColumnCount, int uniqueRowCount)
        {
            var stopwatch = Stopwatch.StartNew();
            var factory = Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider.CreateDefault();
            var plans = new List<Bing.Offices.Providers.IExcelMappingPlan>();
            var sampledLohPeakBytes = GetLohSizeBeforeBytes();
            for (var tenant = 0; tenant < tenantCount; tenant++)
            {
                var document = CreateDocument(tenant);
                plans.Add(factory.Create<ProbeRow>(document, Bing.Offices.Configurations.MappingDirection.Import));
            }
            sampledLohPeakBytes = Math.Max(sampledLohPeakBytes, GetLohSizeBeforeBytes());
            for (var index = 0; index < planBuildCount; index++)
            {
                var document = CreateDocument(index % Math.Max(1, tenantCount));
                plans.Add(factory.Create<ProbeRow>(document, Bing.Offices.Configurations.MappingDirection.Import));
            }

            var values = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var tracker = new Bing.Offices.Providers.UniqueTracker(values,
                uniqueRowCount * uniqueColumnCount);
            for (var row = 0; row < uniqueRowCount; row++)
            {
                tracker.BeginRow();
                for (var column = 0; column < uniqueColumnCount; column++)
                    tracker.TryReserve($"unique-{column}", $"value-{column}-{row}", false, false, row + 1);
                tracker.CommitRow();
            }
            sampledLohPeakBytes = Math.Max(sampledLohPeakBytes, GetLohSizeBeforeBytes());

            var gcBefore = GetLohSizeBeforeBytes();
            sampledLohPeakBytes = Math.Max(sampledLohPeakBytes, gcBefore);
            GC.Collect(2, GCCollectionMode.Forced, true, false);
            GC.KeepAlive(plans);
            GC.KeepAlive(tracker);
            GC.KeepAlive(values);
            var gcAfter = GC.GetGCMemoryInfo().GenerationInfo[3].SizeAfterBytes;
            var lohRetainedBytes = gcAfter;
            sampledLohPeakBytes = Math.Max(sampledLohPeakBytes, lohRetainedBytes);
            var peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64;
            var passed = sampledLohPeakBytes <= LohCeilingBytes
                         && lohRetainedBytes <= LohCeilingBytes
                         && peakWorkingSetBytes <= PeakWorkingSetCeilingBytes;
            stopwatch.Stop();
            var result = new
            {
                kind = "scenario",
                artifact = artifactPath,
                planBuildCount,
                tenantCount,
                uniqueColumnCount,
                uniqueRowCount,
                tenantPlanCount = tenantCount,
                planBuildCountAfterTenantWarmup = planBuildCount,
                workload = "mapping-plan-and-unique-tracker",
                gcLohSizeBeforeBytes = gcBefore,
                gcLohSizeAfterBytes = gcAfter,
                lohSampledPeakBytes = sampledLohPeakBytes,
                lohRetainedBytes,
                peakWorkingSetBytes,
                lohCeilingBytes = LohCeilingBytes,
                peakWorkingSetCeilingBytes = PeakWorkingSetCeilingBytes,
                status = passed ? "passed" : "failed",
                elapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
                exitCode = passed ? 0 : 1
            };
            Console.WriteLine(JsonSerializer.Serialize(result));
            Environment.ExitCode = passed ? 0 : 1;
        }

        /// <summary>
        /// 获取当前大对象堆回收前的大小。
        /// </summary>
        /// <returns>大对象堆大小，单位为字节。</returns>
        private static long GetLohSizeBeforeBytes() =>
            GC.GetGCMemoryInfo().GenerationInfo[3].SizeBeforeBytes;

        /// <summary>
        /// 启动子进程执行单个资源场景并包装结果。
        /// </summary>
        /// <param name="artifactPath">结果所属的资源探针路径。</param>
        /// <param name="planBuildCount">额外构建映射计划的次数。</param>
        /// <param name="tenantCount">租户配置数量。</param>
        /// <param name="uniqueColumnCount">唯一值跟踪的列数。</param>
        /// <param name="uniqueRowCount">唯一值跟踪的行数。</param>
        /// <returns>子进程结果的 JSON 文本。</returns>
        private static string RunChild(string artifactPath, int planBuildCount, int tenantCount,
            int uniqueColumnCount, int uniqueRowCount)
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
            process.StartInfo.ArgumentList.Add("--resource-scenario");
            process.StartInfo.ArgumentList.Add(artifactPath);
            process.StartInfo.ArgumentList.Add(planBuildCount.ToString());
            process.StartInfo.ArgumentList.Add(tenantCount.ToString());
            process.StartInfo.ArgumentList.Add(uniqueColumnCount.ToString());
            process.StartInfo.ArgumentList.Add(uniqueRowCount.ToString());
            if (!process.Start())
                throw new InvalidOperationException("无法启动资源探针子进程。");
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            var line = stdout.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).LastOrDefault()
                ?? string.Empty;
            using var scenario = JsonDocument.Parse(line);
            return JsonSerializer.Serialize(new
            {
                kind = "child",
                planBuildCount,
                tenantCount,
                uniqueColumnCount,
                uniqueRowCount,
                exitCode = process.ExitCode,
                result = scenario.RootElement.Clone(),
                stderr
            });
        }

        /// <summary>
        /// 创建资源探针使用的租户映射文档。
        /// </summary>
        /// <param name="tenant">租户序号。</param>
        /// <returns>租户映射文档。</returns>
        private static Bing.Offices.Configurations.ExcelMappingDocument CreateDocument(int tenant)
            => new()
            {
                TenantId = $"tenant-{tenant}",
                ConfigurationVersion = "resource-probe",
                Import = new Bing.Offices.Configurations.ExcelMappingConfiguration
                {
                    Columns =
                    {
                        new Bing.Offices.Configurations.ExcelColumnConfiguration
                        {
                            PropertyName = nameof(ProbeRow.Code), Title = "编码"
                        }
                    }
                }
            };

        /// <summary>
        /// 资源探针使用的映射行模型。
        /// </summary>
        private sealed class ProbeRow
        {
            /// <summary>
            /// 获取或设置行编码。
            /// </summary>
            public string Code { get; set; } = string.Empty;
        }
    }

    /// <summary>
    /// 运行异步资源准入的尾延迟探针。
    /// </summary>
    private static class TailLatency
    {
        /// <summary>
        /// 尾延迟探针的预热操作数上限。
        /// </summary>
        private const int WarmupOperationCount = 64;

        /// <summary>
        /// 每个并发度执行的测量重复次数。
        /// </summary>
        private const int RepetitionCount = 5;

        /// <summary>
        /// 执行映射计划冷启动尾延迟探针。
        /// </summary>
        /// <param name="artifactPath">JSONL 输出路径。</param>
        /// <param name="operationCount">每个并发度的测量操作数。</param>
        public static void Run(string artifactPath, int operationCount)
        {
            if (operationCount < 1)
                throw new ArgumentOutOfRangeException(nameof(operationCount));

            var fullPath = Path.GetFullPath(artifactPath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
            writer.WriteLine(JsonSerializer.Serialize(new
            {
                kind = "tail-latency",
                schema = 1,
                generatedUtc = DateTimeOffset.UtcNow,
                dotnet = Environment.Version.ToString(),
                operationCount,
                warmupOperationCount = Math.Min(WarmupOperationCount, operationCount),
                repetitionCount = RepetitionCount,
                processorCount = Environment.ProcessorCount,
                os = RuntimeInformation.OSDescription,
                processArchitecture = RuntimeInformation.ProcessArchitecture.ToString(),
                framework = RuntimeInformation.FrameworkDescription,
                serverGc = GCSettings.IsServerGC,
                gcLatencyMode = GCSettings.LatencyMode.ToString(),
                stopwatchFrequency = Stopwatch.Frequency,
                processId = Environment.ProcessId,
                processPath = Environment.ProcessPath,
                gitHead = Environment.GetEnvironmentVariable("BING_OFFICES_GIT_HEAD") ?? "not-provided",
                diffIdentity = Environment.GetEnvironmentVariable("BING_OFFICES_DIFF_ID") ?? "not-provided",
                workload = "cold-plan-build",
                latencyDefinition = "从队列提交前时间戳到 mapping plan 完成的端到端样本",
                budgetStatus = "UNAPPROVED"
            }));

            foreach (var concurrency in new[] { 1, 4, 16, 64 })
            {
                RunBatch(concurrency, Math.Min(WarmupOperationCount, operationCount), false);
                var samples = new List<long>(operationCount * RepetitionCount);
                var elapsedSeconds = 0d;
                var workerStartupMilliseconds = 0d;
                for (var repetition = 1; repetition <= RepetitionCount; repetition++)
                {
                    var batch = RunBatch(concurrency, operationCount, true);
                    samples.AddRange(batch.Samples);
                    elapsedSeconds += batch.ElapsedSeconds;
                    workerStartupMilliseconds += batch.WorkerStartupMilliseconds;
                    var repetitionSamples = batch.Samples.OrderBy(sample => sample).ToArray();
                    writer.WriteLine(JsonSerializer.Serialize(new
                    {
                        kind = "tail-latency-repetition",
                        concurrency,
                        repetition,
                        operationCount,
                        p50Microseconds = Percentile(repetitionSamples, 0.50),
                        p95Microseconds = Percentile(repetitionSamples, 0.95),
                        p99Microseconds = Percentile(repetitionSamples, 0.99),
                        throughputOperationsPerSecond = operationCount / batch.ElapsedSeconds,
                        workerStartupMilliseconds = batch.WorkerStartupMilliseconds,
                        budgetStatus = "UNAPPROVED"
                    }));
                }

                var sortedSamples = samples.OrderBy(sample => sample).ToArray();
                var result = new
                {
                    kind = "tail-latency-scenario",
                    concurrency,
                    operationCount,
                    warmupOperationCount = Math.Min(WarmupOperationCount, operationCount),
                    repetitionCount = RepetitionCount,
                    sampleCount = sortedSamples.Length,
                    p50Microseconds = Percentile(sortedSamples, 0.50),
                    p95Microseconds = Percentile(sortedSamples, 0.95),
                    p99Microseconds = Percentile(sortedSamples, 0.99),
                    throughputOperationsPerSecond = operationCount * RepetitionCount / elapsedSeconds,
                    averageWorkerStartupMilliseconds = workerStartupMilliseconds / RepetitionCount,
                    budgetStatus = "UNAPPROVED"
                };
                writer.WriteLine(JsonSerializer.Serialize(result));
                writer.Flush();
            }
            Console.WriteLine($"TAIL_LATENCY artifact={fullPath} scenarios=4 budget=UNAPPROVED status=measured");
        }

        /// <summary>
        /// 以指定并发度执行一批映射计划构建操作。
        /// </summary>
        /// <param name="concurrency">worker 并发数。</param>
        /// <param name="operationCount">本批操作数。</param>
        /// <param name="captureSamples">是否记录单操作样本。</param>
        /// <returns>批次耗时、启动耗时和样本。</returns>
        private static BatchResult RunBatch(int concurrency, int operationCount, bool captureSamples)
        {
            using var queue = new BlockingCollection<int>();
            using var ready = new CountdownEvent(concurrency);
            using var startGate = new ManualResetEventSlim(false);
            var documents = Enumerable.Range(0, operationCount)
                .Select(CreateDocument)
                .ToArray();
            var submittedAt = new long[operationCount];
            var samples = captureSamples ? new long[operationCount] : Array.Empty<long>();
            var factory = Bing.Offices.Mappings.ExcelMappingPlanFactoryProvider.CreateDefault();
            var workers = Enumerable.Range(0, concurrency)
                .Select(_ => Task.Factory.StartNew(
                    () => RunWorker(queue, documents, submittedAt, samples, factory, captureSamples,
                        ready, startGate),
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default))
                .ToArray();
            var startupStopwatch = Stopwatch.StartNew();
            ready.Wait();
            startupStopwatch.Stop();
            var stopwatch = Stopwatch.StartNew();
            startGate.Set();
            for (var index = 0; index < operationCount; index++)
            {
                submittedAt[index] = Stopwatch.GetTimestamp();
                queue.Add(index);
            }
            queue.CompleteAdding();
            Task.WaitAll(workers);
            stopwatch.Stop();
            return new BatchResult(samples, stopwatch.Elapsed.TotalSeconds,
                startupStopwatch.Elapsed.TotalMilliseconds);
        }

        /// <summary>
        /// 消费队列并构建分配到当前 worker 的映射计划。
        /// </summary>
        /// <param name="queue">待处理操作索引队列。</param>
        /// <param name="documents">按操作索引排列的映射文档。</param>
        /// <param name="submittedAt">各操作提交时间戳。</param>
        /// <param name="samples">各操作的延迟样本数组。</param>
        /// <param name="factory">映射计划工厂。</param>
        /// <param name="captureSamples">是否写入延迟样本。</param>
        /// <param name="ready">worker 就绪计数器。</param>
        /// <param name="startGate">统一开始信号。</param>
        private static void RunWorker(BlockingCollection<int> queue,
            Bing.Offices.Configurations.ExcelMappingDocument[] documents,
            long[] submittedAt,
            long[] samples,
            Bing.Offices.Providers.IExcelMappingPlanFactory factory,
            bool captureSamples,
            CountdownEvent ready,
            ManualResetEventSlim startGate)
        {
            ready.Signal();
            startGate.Wait();
            foreach (var index in queue.GetConsumingEnumerable())
            {
                _ = factory.Create<TailLatencyRow>(documents[index],
                    Bing.Offices.Configurations.MappingDirection.Import);
                if (captureSamples)
                {
                    var elapsedTicks = Stopwatch.GetTimestamp() - submittedAt[index];
                    samples[index] = Math.Max(1L,
                        (long)(elapsedTicks * (1_000_000d / Stopwatch.Frequency)));
                }
            }
        }

        /// <summary>
        /// 创建尾延迟探针使用的映射文档。
        /// </summary>
        /// <param name="index">文档序号。</param>
        /// <returns>尾延迟映射文档。</returns>
        private static Bing.Offices.Configurations.ExcelMappingDocument CreateDocument(int index) => new()
        {
            TenantId = $"tail-{index}",
            ConfigurationVersion = "tail-latency",
            Import = new Bing.Offices.Configurations.ExcelMappingConfiguration
            {
                Columns =
                {
                    new Bing.Offices.Configurations.ExcelColumnConfiguration
                    {
                        PropertyName = nameof(TailLatencyRow.Code), Title = "编码"
                    }
                }
            }
        };

        /// <summary>
        /// 从已排序的延迟样本中读取指定分位点。
        /// </summary>
        /// <param name="sortedSamples">升序排列的延迟样本。</param>
        /// <param name="percentile">分位点，通常位于 0 到 1 之间。</param>
        /// <returns>对应分位点的延迟，单位为微秒。</returns>
        private static long Percentile(long[] sortedSamples, double percentile)
        {
            var index = (int)Math.Ceiling(sortedSamples.Length * percentile) - 1;
            return sortedSamples[Math.Clamp(index, 0, sortedSamples.Length - 1)];
        }

        /// <summary>
        /// 尾延迟批次测量结果。
        /// </summary>
        private sealed class BatchResult
        {
            /// <summary>
            /// 初始化批次测量结果。
            /// </summary>
            /// <param name="samples">批次延迟样本。</param>
            /// <param name="elapsedSeconds">批次总耗时，单位为秒。</param>
            /// <param name="workerStartupMilliseconds">worker 启动耗时，单位为毫秒。</param>
            public BatchResult(long[] samples, double elapsedSeconds, double workerStartupMilliseconds)
            {
                Samples = samples;
                ElapsedSeconds = elapsedSeconds;
                WorkerStartupMilliseconds = workerStartupMilliseconds;
            }

            /// <summary>
            /// 获取批次延迟样本。
            /// </summary>
            public long[] Samples { get; }

            /// <summary>
            /// 获取批次总耗时，单位为秒。
            /// </summary>
            public double ElapsedSeconds { get; }

            /// <summary>
            /// 获取 worker 启动耗时，单位为毫秒。
            /// </summary>
            public double WorkerStartupMilliseconds { get; }
        }

        /// <summary>
        /// 尾延迟探针使用的映射行模型。
        /// </summary>
        private sealed class TailLatencyRow
        {
            /// <summary>
            /// 获取或设置行编码。
            /// </summary>
            public string Code { get; set; } = string.Empty;
        }
    }
}
