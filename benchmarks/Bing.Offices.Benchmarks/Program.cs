using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using BenchmarkDotNet.Running;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 性能基准程序入口。
/// </summary>
public static class Program
{
    /// <summary>
    /// 运行流式 Excel 管线基准。
    /// </summary>
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
        BenchmarkSwitcher.FromTypes(
            new[]
            {
                typeof(StreamPipelineBenchmarks),
                typeof(FailureWorkbookBenchmarks),
                typeof(HeaderStyleBenchmarks),
                typeof(ValidationRangeBenchmarks),
                typeof(MappingValidationBenchmarks),
                typeof(DynamicPlanBenchmarks),
                typeof(TenantPlanCacheBenchmarks),
                typeof(RegexCacheBenchmarks),
                typeof(UniqueJournalBenchmarks)
            }).Run(args);
    }

    private static class ResourceProbe
    {
        private const long LohCeilingBytes = 512L * 1024 * 1024;
        private const long PeakWorkingSetCeilingBytes = 1024L * 1024 * 1024;
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

        public static void RunScenario(string artifactPath, int planBuildCount, int tenantCount,
            int uniqueColumnCount, int uniqueRowCount)
        {
            var stopwatch = Stopwatch.StartNew();
            var factory = new Bing.Offices.Mappings.ExcelMappingPlanFactory(
                cacheCapacity: Math.Max(1, tenantCount));
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

        private static long GetLohSizeBeforeBytes() =>
            GC.GetGCMemoryInfo().GenerationInfo[3].SizeBeforeBytes;

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

        private sealed class ProbeRow
        {
            public string Code { get; set; } = string.Empty;
        }
    }

    private static class TailLatency
    {
        private const int WarmupOperationCount = 64;
        private const int RepetitionCount = 5;

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
            var factory = new Bing.Offices.Mappings.ExcelMappingPlanFactory(
                cacheCapacity: Math.Max(1, operationCount));
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

        private static void RunWorker(BlockingCollection<int> queue,
            Bing.Offices.Configurations.ExcelMappingDocument[] documents,
            long[] submittedAt,
            long[] samples,
            Bing.Offices.Mappings.ExcelMappingPlanFactory factory,
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

        private static long Percentile(long[] sortedSamples, double percentile)
        {
            var index = (int)Math.Ceiling(sortedSamples.Length * percentile) - 1;
            return sortedSamples[Math.Clamp(index, 0, sortedSamples.Length - 1)];
        }

        private sealed class BatchResult
        {
            public BatchResult(long[] samples, double elapsedSeconds, double workerStartupMilliseconds)
            {
                Samples = samples;
                ElapsedSeconds = elapsedSeconds;
                WorkerStartupMilliseconds = workerStartupMilliseconds;
            }

            public long[] Samples { get; }

            public double ElapsedSeconds { get; }

            public double WorkerStartupMilliseconds { get; }
        }

        private sealed class TailLatencyRow
        {
            public string Code { get; set; } = string.Empty;
        }
    }
}
