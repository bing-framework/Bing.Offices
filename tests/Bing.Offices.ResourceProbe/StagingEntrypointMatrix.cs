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

/// <summary>
/// 提供测试场景使用的资源矩阵。
/// </summary>
internal static class StagingEntrypointMatrix
{
    /// <summary>
    /// 资源入口矩阵任务标识。
    /// </summary>
    private const string TaskId = "BO-RC-20260908-002";
    /// <summary>
    /// 入口矩阵使用的 staging 策略名称。
    /// </summary>
    private const string Strategy = "TempFile";
    /// <summary>
    /// 每个入口与行数组合的重复次数。
    /// </summary>
    private const int RepetitionCount = 3;
    /// <summary>
    /// 混合 staging 迁移到临时文件的字节阈值。
    /// </summary>
    private const long HybridThresholdBytes = 8L * 1024 * 1024;
    /// <summary>
    /// 入口矩阵测试覆盖的数据行数集合。
    /// </summary>
    private static readonly int[] RowCounts = { 1000, 10000, 100000 };
    /// <summary>
    /// 入口矩阵测试覆盖的 Excel 导出入口名称集合。
    /// </summary>
    private static readonly string[] Entrypoints =
    {
        "Export",
        "ExportAsync",
        "ExportToFile",
        "ExportToFileAsync"
    };

    /// <summary>
    /// 表示失败场景使用的测试夹具。
    /// </summary>
    private enum FailureMode
    {
        /// <summary>
        /// 表示未发生失败。
        /// </summary>
        None,
        /// <summary>
        /// 表示资源探测的初始化阶段。
        /// </summary>
        Setup,
        /// <summary>
        /// 表示资源探测的采样阶段。
        /// </summary>
        Sampler,
        /// <summary>
        /// 表示资源探测的解析阶段。
        /// </summary>
        Parser
    }

    /// <summary>
    /// 运行。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <returns>计算得到的数值。</returns>
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

    /// <summary>
    /// 运行失败路径探针。
    /// </summary>
    /// <param name="artifactPath">目标文件或目录路径。</param>
    /// <returns>计算得到的数值。</returns>
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

    /// <summary>
    /// 运行单个入口探针场景。
    /// </summary>
    /// <param name="entrypoint">入口结果。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <param name="repetition">重复次数。</param>
    /// <param name="failureMode">失败处理模式。</param>
    /// <returns>包含输出校验、资源指标、清理状态和错误信息的入口探测结果。</returns>
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

    /// <summary>
    /// 尝试重新打开工作簿。
    /// </summary>
    /// <param name="outputBytes">待处理的字节内容。</param>
    /// <param name="rowCount">要处理的数据行数。</param>
    /// <param name="error">错误信息。</param>
    /// <returns>工作簿可重新打开且首尾行内容匹配时为 true；发生异常时为 false。</returns>
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

    /// <summary>
    /// 创建导出器。
    /// </summary>
    /// <param name="directory">目标目录路径。</param>
    /// <returns>配置了测试暂存工厂的 NPOI 导出器。</returns>
    private static IExcelExporter CreateExporter(string directory)
    {
        var factory = CreateStagingFactory(directory);
        var constructor = typeof(NpoiExcelExporter).GetConstructors(BindingFlags.Instance | BindingFlags.NonPublic)
            .Single(item => item.GetParameters().Length == 2
                && item.GetParameters()[0].ParameterType == typeof(IFileExportCommitter)
                && item.GetParameters()[1].ParameterType.Name == "INpoiAsyncStagingFactory");
        return (IExcelExporter)constructor.Invoke(new object[] { new DefaultFileExportCommitter(), factory });
    }

    /// <summary>
    /// 创建暂存工厂实例。
    /// </summary>
    /// <param name="directory">目标目录路径。</param>
    /// <returns>使用指定暂存策略、阈值和目录的 NPOI 暂存工厂实例。</returns>
    private static object CreateStagingFactory(string directory)
    {
        var assembly = typeof(NpoiExcelExporter).Assembly;
        var factoryType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingFactory", true);
        var strategyType = assembly.GetType("Bing.Offices.IO.NpoiAsyncStagingStrategy", true);
        var strategyValue = Enum.Parse(strategyType, Strategy);
        return Activator.CreateInstance(factoryType, BindingFlags.Instance | BindingFlags.NonPublic, null,
            new object[] { strategyValue, HybridThresholdBytes, directory }, CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// 尝试删除暂存目录。
    /// </summary>
    /// <param name="directory">目标目录路径。</param>
    /// <param name="exception">测试期间要传播的异常。</param>
    /// <returns>目录已删除或原本不存在时为 true；删除失败时为 false。</returns>
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

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class EntrypointRow
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
        /// 获取或设置描述。
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// 表示测试场景的结果数据模型。
    /// </summary>
    private sealed class EntrypointResult
    {
        /// <summary>
        /// 获取或设置任务标识。
        /// </summary>
        public string TaskId { get; set; }
        /// <summary>
        /// 获取或设置策略。
        /// </summary>
        public string Strategy { get; set; }
        /// <summary>
        /// 获取或设置入口名称。
        /// </summary>
        public string Entrypoint { get; set; }
        /// <summary>
        /// 获取或设置失败模式。
        /// </summary>
        public string FailureMode { get; set; }
        /// <summary>
        /// 获取或设置行数。
        /// </summary>
        public int RowCount { get; set; }
        /// <summary>
        /// 获取或设置重复次数。
        /// </summary>
        public int Repetition { get; set; }
        /// <summary>
        /// 获取或设置耗时（毫秒）。
        /// </summary>
        public double ElapsedMilliseconds { get; set; }
        /// <summary>
        /// 获取或设置已分配字节数。
        /// </summary>
        public long AllocatedBytes { get; set; }
        /// <summary>
        /// 获取或设置第 0 代 GC 次数。
        /// </summary>
        public int Gen0Collections { get; set; }
        /// <summary>
        /// 获取或设置第 1 代 GC 次数。
        /// </summary>
        public int Gen1Collections { get; set; }
        /// <summary>
        /// 获取或设置第 2 代 GC 次数。
        /// </summary>
        public int Gen2Collections { get; set; }
        /// <summary>
        /// 获取或设置峰值工作集字节数。
        /// </summary>
        public long PeakWorkingSetBytes { get; set; }
        /// <summary>
        /// 获取或设置临时磁盘峰值字节数。
        /// </summary>
        public long TemporaryDiskPeakBytes { get; set; }
        /// <summary>
        /// 获取或设置临时文件峰值数量。
        /// </summary>
        public int TemporaryFilePeakCount { get; set; }
        /// <summary>
        /// 获取或设置完成后的临时磁盘字节数。
        /// </summary>
        public long TemporaryDiskBytesAfter { get; set; }
        /// <summary>
        /// 获取或设置残留文件集合。
        /// </summary>
        public int LeftoverFiles { get; set; }
        /// <summary>
        /// 获取或设置输出数据字节数。
        /// </summary>
        public long OutputBytes { get; set; }
        /// <summary>
        /// 获取或设置输出数据 SHA-256 摘要。
        /// </summary>
        public string OutputSha256 { get; set; }
        /// <summary>
        /// 获取或设置文件句柄重开是否适用。
        /// </summary>
        public bool FileHandleReopenApplicable { get; set; }
        /// <summary>
        /// 获取或设置文件句柄重开是否成功。
        /// </summary>
        public bool? FileHandleReopenSucceeded { get; set; }
        /// <summary>
        /// 获取或设置解析器是否成功重新打开。
        /// </summary>
        public bool ParserReopenSucceeded { get; set; }
        /// <summary>
        /// 获取或设置是否成功重新打开。
        /// </summary>
        public bool ReopenSucceeded { get; set; }
        /// <summary>
        /// 获取或设置采样是否成功。
        /// </summary>
        public bool SamplingSucceeded { get; set; }
        /// <summary>
        /// 获取或设置清理是否成功。
        /// </summary>
        public bool CleanupSucceeded { get; set; }
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 获取或设置异常。
        /// </summary>
        public string Exception { get; set; }
    }

    /// <summary>
    /// 提供资源探测场景使用的采样器。
    /// </summary>
    private sealed class DirectorySampler
    {
        /// <summary>
        /// 采样器监控的工作目录路径。
        /// </summary>
        private readonly string _directory;
        /// <summary>
        /// 采样时排除的输出文件路径。
        /// </summary>
        private readonly string _excludedPath;
        /// <summary>
        /// 是否在采样过程中注入失败。
        /// </summary>
        private readonly bool _injectFailure;
        /// <summary>
        /// 停止目录采样的取消源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation = new();
        /// <summary>
        /// 目录采样后台任务。
        /// </summary>
        private Task _task;
        /// <summary>
        /// 采样期间观测到的临时文件总字节峰值。
        /// </summary>
        private long _peakBytes;
        /// <summary>
        /// 采样期间观测到的临时文件数量峰值。
        /// </summary>
        private int _peakFiles;

        /// <summary>
        /// 初始化一个 <see cref="DirectorySampler" /> 类型的实例。
        /// </summary>
        /// <param name="directory">待采样的工作目录。</param>
        /// <param name="excludedPath">不纳入采样的输出文件路径；null 表示不排除。</param>
        /// <param name="injectFailure">是否在采样时注入 IO 异常。</param>
        public DirectorySampler(string directory, string excludedPath, bool injectFailure)
        {
            _directory = directory;
            _excludedPath = excludedPath == null ? null : Path.GetFullPath(excludedPath);
            _injectFailure = injectFailure;
        }

        /// <summary>
        /// 获取峰值字节数。
        /// </summary>
        public long PeakBytes => Interlocked.Read(ref _peakBytes);
        /// <summary>
        /// 获取峰值文件数。
        /// </summary>
        public int PeakFiles => Volatile.Read(ref _peakFiles);
        /// <summary>
        /// 获取或设置瞬时文件数量。
        /// </summary>
        public int TransientFileCount { get; private set; }
        /// <summary>
        /// 获取或设置完成后的瞬时字节数。
        /// </summary>
        public long TransientBytesAfter { get; private set; }

        /// <summary>
        /// 开始资源采样或请求批次。
        /// </summary>
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

        /// <summary>
        /// 停止资源采样。
        /// </summary>
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

        /// <summary>
        /// 采样暂存目录的文件状态。
        /// </summary>
        private void Sample()
        {
            if (_injectFailure)
                throw new IOException("injected sampler failure.");
            var files = GetTransientFiles();
            var bytes = files.Sum(GetFileLength);
            UpdateMaximum(ref _peakBytes, bytes);
            UpdateMaximum(ref _peakFiles, files.Count);
        }

        /// <summary>
        /// 获取目录中的临时文件。
        /// </summary>
        /// <returns>目录中除输出文件外的临时文件路径列表；目录不存在时为空列表。</returns>
        private List<string> GetTransientFiles() => Directory.Exists(_directory)
            ? Directory.GetFiles(_directory, "*", SearchOption.AllDirectories)
                .Where(path => _excludedPath == null
                    || !string.Equals(Path.GetFullPath(path), _excludedPath, StringComparison.OrdinalIgnoreCase))
                .ToList()
            : new List<string>();

        /// <summary>
        /// 获取文件长度。
        /// </summary>
        /// <param name="path">目标文件或目录路径。</param>
        /// <returns>文件长度，单位为字节；读取失败时返回 0。</returns>
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

        /// <summary>
        /// 更新采样值的最大值。
        /// </summary>
        /// <param name="target">目标对象。</param>
        /// <param name="candidate">候选项。</param>
        private static void UpdateMaximum(ref long target, long candidate)
        {
            while (true)
            {
                var current = Interlocked.Read(ref target);
                if (candidate <= current || Interlocked.CompareExchange(ref target, candidate, current) == current)
                    return;
            }
        }

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
}
