using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 使用公开 Provider API 采集实体行 materialization 和 binding 的真实导入/导出证据。
/// </summary>
internal static class MaterializationBindingProbe
{
    /// <summary>
    /// 探针 JSON 文档的版本号。
    /// </summary>
    private const int SchemaVersion = 2;
    /// <summary>
    /// 固定输入数据集使用的种子值。
    /// </summary>
    private const int FixedSeed = 20260920;
    /// <summary>
    /// 固定映射数据集的列数。
    /// </summary>
    private const int ColumnCount = 6;
    /// <summary>
    /// 当前候选缺少可比较 before 数据时使用的基线状态。
    /// </summary>
    private const string BaselineStatus = "NOT_APPLICABLE_BEFORE";
    /// <summary>
    /// 说明当前探针未提供独立 before 候选的原因。
    /// </summary>
    private const string BaselineReason =
        "当前工作区没有可运行且身份独立的 before candidate；本探针只记录当前候选的隔离 E2E 证据。";
    /// <summary>
    /// 参与探针采集的 Provider 名称。
    /// </summary>
    private static readonly string[] Providers = { "npoi", "miniexcel" };
    /// <summary>
    /// 参与探针采集的执行模式。
    /// </summary>
    private static readonly string[] Modes = { "sync", "async" };
    /// <summary>
    /// 参与探针采集的操作类型。
    /// </summary>
    private static readonly string[] Operations = { "export-only", "import-only", "roundtrip" };

    /// <summary>
    /// 运行固定数据、显式映射以及分离 import-only/export-only 的 Provider 探针。
    /// </summary>
    /// <param name="artifactPath">JSON 输出路径。</param>
    /// <param name="rowCount">每种操作的数据行数。</param>
    /// <param name="repetitions">每种 Provider/mode/operation 的重复次数。</param>
    public static async Task RunAsync(string artifactPath, int rowCount, int repetitions)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (repetitions < 1)
            throw new ArgumentOutOfRangeException(nameof(repetitions));

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        var rows = CreateRows(rowCount);
        var mapping = CreateMapping();
        var samples = new List<Sample>(repetitions * Providers.Length * Modes.Length * Operations.Length);
        var fixtures = new Dictionary<string, Fixture>(StringComparer.OrdinalIgnoreCase);

        foreach (var provider in Providers)
        {
            try
            {
                fixtures[provider] = await CreateFixtureAsync(provider, rows, mapping).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                foreach (var mode in Modes)
                {
                    foreach (var operation in Operations)
                    {
                        for (var repetition = 1; repetition <= repetitions; repetition++)
                            samples.Add(Sample.Failed(provider, mode, operation, repetition, exception));
                    }
                }
                continue;
            }

            var fixture = fixtures[provider];
            foreach (var mode in Modes)
            {
                foreach (var operation in Operations)
                {
                    for (var repetition = 1; repetition <= repetitions; repetition++)
                    {
                        try
                        {
                            samples.Add(await MeasureAsync(provider, mode, operation, repetition, rows, mapping,
                                fixture).ConfigureAwait(false));
                        }
                        catch (Exception exception)
                        {
                            samples.Add(Sample.Failed(provider, mode, operation, repetition, exception,
                                fixture.Identity, fixture.Bytes.Length));
                        }
                    }
                }
            }
        }

        var document = new ProbeDocument
        {
            Kind = "materialization-binding-probe",
            Schema = SchemaVersion,
            GeneratedUtc = DateTimeOffset.UtcNow,
            Framework = RuntimeInformation.FrameworkDescription,
            Os = RuntimeInformation.OSDescription,
            ProcessId = Environment.ProcessId,
            ProcessorCount = Environment.ProcessorCount,
            Seed = FixedSeed,
            RowCount = rowCount,
            ColumnCount = ColumnCount,
            Repetitions = repetitions,
            Providers = Providers,
            Modes = Modes,
            Operations = Operations,
            Workload = "public-provider-export-only-import-only-roundtrip-explicit-binding",
            DatasetIdentity = HashRows(rows),
            CandidateIdentity = GetCandidateIdentity(),
            BaselineStatus = BaselineStatus,
            BaselineReason = BaselineReason,
            FixtureIdentities = fixtures.ToDictionary(pair => pair.Key, pair => pair.Value.Identity,
                StringComparer.OrdinalIgnoreCase),
            ThresholdStatus = "UNAPPROVED",
            EvidenceStatus = samples.All(sample => sample.Status == "measured")
                ? "measured"
                : "failed",
            Samples = samples,
            Summaries = CreateSummaries(samples)
        };

        var json = JsonSerializer.Serialize(document, new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });
        ValidateSerializedArtifact(json, document);
        await File.WriteAllTextAsync(fullPath, json, new UTF8Encoding(false)).ConfigureAwait(false);

        Console.WriteLine($"MATERIALIZATION_BINDING_PROBE artifact={fullPath} rows={rowCount} "
            + $"repetitions={repetitions} samples={samples.Count} operations={string.Join(',', Operations)} "
            + $"status={document.EvidenceStatus} baseline={BaselineStatus} threshold=UNAPPROVED");
        if (document.EvidenceStatus == "failed")
            Environment.ExitCode = 1;
    }

    /// <summary>
    /// 为指定 Provider 创建并校验固定输入工作簿。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="rows">用于生成工作簿的数据行。</param>
    /// <param name="mapping">导出使用的映射配置。</param>
    /// <returns>包含工作簿字节和身份信息的测试夹具。</returns>
    private static async Task<Fixture> CreateFixtureAsync(string provider,
        IReadOnlyList<MaterializationRow> rows, ExcelMappingConfiguration mapping)
    {
        using var services = provider == "npoi"
            ? new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider()
            : new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        var exporter = services.GetRequiredService<IExcelExporter>();
        using var output = new MemoryStream();
        await exporter.ExportAsync(CreateExportRequest(rows, mapping), output).ConfigureAwait(false);
        var bytes = output.ToArray();
        if (!HasWorkbookStructure(bytes))
            throw new InvalidDataException($"{provider} fixture 未生成完整 XLSX 结构。");
        return new Fixture(provider, bytes, HashBytes(bytes), rows.Count);
    }

    /// <summary>
    /// 按执行模式测量一次 Provider 操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">同步或异步执行模式。</param>
    /// <param name="operation">导出、导入或往返操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="mapping">操作使用的映射配置。</param>
    /// <param name="fixture">预生成的输入工作簿。</param>
    /// <returns>本次测量结果。</returns>
    private static async Task<Sample> MeasureAsync(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, ExcelMappingConfiguration mapping, Fixture fixture)
    {
        using var services = provider == "npoi"
            ? new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider()
            : new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        var exporter = services.GetRequiredService<IExcelExporter>();
        var importer = services.GetRequiredService<IExcelImporter>();
        var exportRequest = CreateExportRequest(expectedRows, mapping);
        var importRequest = CreateImportRequest(mapping);
        return mode == "sync"
            ? MeasureSync(provider, mode, operation, repetition, expectedRows, exporter, importer, exportRequest,
                importRequest, fixture)
            : await MeasureAsyncCore(provider, mode, operation, repetition, expectedRows, exporter, importer,
                exportRequest, importRequest, fixture).ConfigureAwait(false);
    }

    /// <summary>
    /// 同步执行一次指定类型的 Provider 操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <returns>本次同步测量结果。</returns>
    private static Sample MeasureSync(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<MaterializationWorkbook> importRequest, Fixture fixture)
    {
        PrepareMeasurement(out var allocatedBefore, out var gen0Before, out var gen1Before, out var gen2Before,
            out var process);
        return operation switch
        {
            "export-only" => MeasureExportSync(provider, mode, operation, repetition, expectedRows, exporter,
                exportRequest, fixture, allocatedBefore, gen0Before, gen1Before, gen2Before, process),
            "import-only" => MeasureImportSync(provider, mode, operation, repetition, expectedRows, importer,
                importRequest, fixture, allocatedBefore, gen0Before, gen1Before, gen2Before, process),
            "roundtrip" => MeasureRoundtripSync(provider, mode, operation, repetition, expectedRows, exporter,
                importer, exportRequest, importRequest, fixture, allocatedBefore, gen0Before, gen1Before,
                gen2Before, process),
            _ => throw new ArgumentOutOfRangeException(nameof(operation), operation, "未知 materialization 操作。")
        };
    }

    /// <summary>
    /// 异步执行一次指定类型的 Provider 操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <returns>最终完成的异步测量结果。</returns>
    private static async Task<Sample> MeasureAsyncCore(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<MaterializationWorkbook> importRequest, Fixture fixture)
    {
        PrepareMeasurement(out var allocatedBefore, out var gen0Before, out var gen1Before, out var gen2Before,
            out var process);
        switch (operation)
        {
            case "export-only":
                return await MeasureExportAsync(provider, mode, operation, repetition, expectedRows, exporter,
                    exportRequest, fixture, allocatedBefore, gen0Before, gen1Before, gen2Before, process)
                    .ConfigureAwait(false);
            case "import-only":
                return await MeasureImportAsync(provider, mode, operation, repetition, expectedRows, importer,
                    importRequest, fixture, allocatedBefore, gen0Before, gen1Before, gen2Before, process)
                    .ConfigureAwait(false);
            case "roundtrip":
                return await MeasureRoundtripAsync(provider, mode, operation, repetition, expectedRows, exporter,
                    importer, exportRequest, importRequest, fixture, allocatedBefore, gen0Before, gen1Before,
                    gen2Before, process).ConfigureAwait(false);
            default:
                throw new ArgumentOutOfRangeException(nameof(operation), operation, "未知 materialization 操作。");
        }
    }

    /// <summary>
    /// 同步测量一次仅导出操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="request">导出请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>本次导出测量结果。</returns>
    private static Sample MeasureExportSync(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter,
        ExcelWorkbookExportRequest request, Fixture fixture, long allocatedBefore, int gen0Before, int gen1Before,
        int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var output = new MemoryStream();
        var export = Stopwatch.StartNew();
        exporter.Export(request, output);
        export.Stop();
        total.Stop();
        var bytes = output.ToArray();
        return CreateExportSample(provider, mode, operation, repetition, expectedRows, fixture, bytes,
            total.Elapsed.TotalMilliseconds, export.Elapsed.TotalMilliseconds, allocatedBefore, gen0Before,
            gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 异步测量一次仅导出操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="request">导出请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>最终完成的异步导出测量结果。</returns>
    private static async Task<Sample> MeasureExportAsync(string provider, string mode, string operation,
        int repetition, IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter,
        ExcelWorkbookExportRequest request, Fixture fixture, long allocatedBefore, int gen0Before, int gen1Before,
        int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var output = new MemoryStream();
        var export = Stopwatch.StartNew();
        await exporter.ExportAsync(request, output).ConfigureAwait(false);
        export.Stop();
        total.Stop();
        var bytes = output.ToArray();
        return CreateExportSample(provider, mode, operation, repetition, expectedRows, fixture, bytes,
            total.Elapsed.TotalMilliseconds, export.Elapsed.TotalMilliseconds, allocatedBefore, gen0Before,
            gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 同步测量一次仅导入操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>本次导入测量结果。</returns>
    private static Sample MeasureImportSync(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, IExcelImporter importer,
        ExcelWorkbookImportRequest<MaterializationWorkbook> request, Fixture fixture, long allocatedBefore,
        int gen0Before, int gen1Before, int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var source = new MemoryStream(fixture.Bytes, writable: false);
        var import = Stopwatch.StartNew();
        var result = importer.Import(source, request);
        import.Stop();
        total.Stop();
        return CreateImportSample(provider, mode, operation, repetition, expectedRows, fixture,
            result.Workbook.Rows.ToArray(), result.IsSuccess, result.Errors.Count, total.Elapsed.TotalMilliseconds,
            import.Elapsed.TotalMilliseconds, allocatedBefore, gen0Before, gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 异步测量一次仅导入操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="request">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>最终完成的异步导入测量结果。</returns>
    private static async Task<Sample> MeasureImportAsync(string provider, string mode, string operation,
        int repetition, IReadOnlyList<MaterializationRow> expectedRows, IExcelImporter importer,
        ExcelWorkbookImportRequest<MaterializationWorkbook> request, Fixture fixture, long allocatedBefore,
        int gen0Before, int gen1Before, int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var source = new MemoryStream(fixture.Bytes, writable: false);
        var import = Stopwatch.StartNew();
        var result = await importer.ImportAsync(source, request).ConfigureAwait(false);
        import.Stop();
        total.Stop();
        return CreateImportSample(provider, mode, operation, repetition, expectedRows, fixture,
            result.Workbook.Rows.ToArray(), result.IsSuccess, result.Errors.Count, total.Elapsed.TotalMilliseconds,
            import.Elapsed.TotalMilliseconds, allocatedBefore, gen0Before, gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 同步测量一次导出后导入的往返操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>本次往返测量结果。</returns>
    private static Sample MeasureRoundtripSync(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter, IExcelImporter importer,
        ExcelWorkbookExportRequest exportRequest, ExcelWorkbookImportRequest<MaterializationWorkbook> importRequest,
        Fixture fixture, long allocatedBefore, int gen0Before, int gen1Before, int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var output = new MemoryStream();
        var export = Stopwatch.StartNew();
        exporter.Export(exportRequest, output);
        export.Stop();
        var bytes = output.ToArray();
        var import = Stopwatch.StartNew();
        using var source = new MemoryStream(bytes, writable: false);
        var result = importer.Import(source, importRequest);
        import.Stop();
        total.Stop();
        return CreateImportSample(provider, mode, operation, repetition, expectedRows, fixture,
            result.Workbook.Rows.ToArray(), result.IsSuccess, result.Errors.Count, total.Elapsed.TotalMilliseconds,
            export.Elapsed.TotalMilliseconds, import.Elapsed.TotalMilliseconds, bytes, allocatedBefore, gen0Before,
            gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 异步测量一次导出后导入的往返操作。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式标签。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">当前测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="exporter">导出器。</param>
    /// <param name="importer">导入器。</param>
    /// <param name="exportRequest">导出请求。</param>
    /// <param name="importRequest">导入请求。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>最终完成的异步往返测量结果。</returns>
    private static async Task<Sample> MeasureRoundtripAsync(string provider, string mode, string operation,
        int repetition, IReadOnlyList<MaterializationRow> expectedRows, IExcelExporter exporter,
        IExcelImporter importer, ExcelWorkbookExportRequest exportRequest,
        ExcelWorkbookImportRequest<MaterializationWorkbook> importRequest, Fixture fixture, long allocatedBefore,
        int gen0Before, int gen1Before, int gen2Before, Process process)
    {
        var total = Stopwatch.StartNew();
        using var output = new MemoryStream();
        var export = Stopwatch.StartNew();
        await exporter.ExportAsync(exportRequest, output).ConfigureAwait(false);
        export.Stop();
        var bytes = output.ToArray();
        var import = Stopwatch.StartNew();
        using var source = new MemoryStream(bytes, writable: false);
        var result = await importer.ImportAsync(source, importRequest).ConfigureAwait(false);
        import.Stop();
        total.Stop();
        return CreateImportSample(provider, mode, operation, repetition, expectedRows, fixture,
            result.Workbook.Rows.ToArray(), result.IsSuccess, result.Errors.Count, total.Elapsed.TotalMilliseconds,
            export.Elapsed.TotalMilliseconds, import.Elapsed.TotalMilliseconds, bytes, allocatedBefore, gen0Before,
            gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 根据导出测量数据创建样本记录。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="bytes">导出结果字节。</param>
    /// <param name="elapsedMilliseconds">总耗时，单位为毫秒。</param>
    /// <param name="exportElapsedMilliseconds">导出耗时，单位为毫秒。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>导出样本记录。</returns>
    private static Sample CreateExportSample(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, Fixture fixture, byte[] bytes,
        double elapsedMilliseconds, double exportElapsedMilliseconds, long allocatedBefore, int gen0Before,
        int gen1Before, int gen2Before, Process process)
    {
        var structureValid = HasWorkbookStructure(bytes);
        process.Refresh();
        return new Sample
        {
            Provider = provider,
            Mode = mode,
            Operation = operation,
            Repetition = repetition,
            Status = structureValid ? "measured" : "failed",
            ElapsedMilliseconds = elapsedMilliseconds,
            ExportElapsedMilliseconds = exportElapsedMilliseconds,
            AllocatedBytes = Math.Max(0, GC.GetTotalAllocatedBytes(true) - allocatedBefore),
            Gen0Collections = GC.CollectionCount(0) - gen0Before,
            Gen1Collections = GC.CollectionCount(1) - gen1Before,
            Gen2Collections = GC.CollectionCount(2) - gen2Before,
            PeakWorkingSetBytes = process.PeakWorkingSet64,
            OutputBytes = bytes.Length,
            OutputIdentity = HashBytes(bytes),
            FixtureIdentity = fixture.Identity,
            FixtureBytes = fixture.Bytes.Length,
            ExportedRows = expectedRows.Count,
            OutputStructureValid = structureValid,
            Integrity = CreateExportIntegrity(expectedRows.Count),
            Failure = structureValid ? null : new Failure
            {
                Type = "WorkbookStructureValidationFailure",
                Message = "导出输出不是包含 workbook 与 sheet XML 的完整 XLSX。"
            }
        };
    }

    /// <summary>
    /// 根据导入或往返测量数据创建样本记录。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="importedRows">实际导入的数据行。</param>
    /// <param name="importSucceeded">导入是否成功。</param>
    /// <param name="importErrorCount">导入错误数量。</param>
    /// <param name="elapsedMilliseconds">总耗时，单位为毫秒。</param>
    /// <param name="importElapsedMilliseconds">导入耗时，单位为毫秒。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>导入样本记录。</returns>
    private static Sample CreateImportSample(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, Fixture fixture,
        IReadOnlyList<MaterializationRow> importedRows, bool importSucceeded, int importErrorCount,
        double elapsedMilliseconds, double importElapsedMilliseconds, long allocatedBefore, int gen0Before,
        int gen1Before, int gen2Before, Process process)
    {
        return CreateImportSample(provider, mode, operation, repetition, expectedRows, fixture, importedRows,
            importSucceeded, importErrorCount, elapsedMilliseconds, null, importElapsedMilliseconds, fixture.Bytes,
            allocatedBefore, gen0Before, gen1Before, gen2Before, process);
    }

    /// <summary>
    /// 根据导入或往返测量数据创建样本记录。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="mode">执行模式。</param>
    /// <param name="operation">操作类型。</param>
    /// <param name="repetition">测量重复序号。</param>
    /// <param name="expectedRows">预期数据行。</param>
    /// <param name="fixture">输入工作簿夹具。</param>
    /// <param name="importedRows">实际导入的数据行。</param>
    /// <param name="importSucceeded">导入是否成功。</param>
    /// <param name="importErrorCount">导入错误数量。</param>
    /// <param name="elapsedMilliseconds">总耗时，单位为毫秒。</param>
    /// <param name="exportElapsedMilliseconds">导出耗时，单位为毫秒；无导出阶段时为 <see langword="null"/>。</param>
    /// <param name="importElapsedMilliseconds">导入耗时，单位为毫秒。</param>
    /// <param name="bytes">本次输出或输入工作簿字节。</param>
    /// <param name="allocatedBefore">测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">当前进程对象。</param>
    /// <returns>导入或往返样本记录。</returns>
    private static Sample CreateImportSample(string provider, string mode, string operation, int repetition,
        IReadOnlyList<MaterializationRow> expectedRows, Fixture fixture,
        IReadOnlyList<MaterializationRow> importedRows, bool importSucceeded, int importErrorCount,
        double elapsedMilliseconds, double? exportElapsedMilliseconds, double importElapsedMilliseconds,
        byte[] bytes, long allocatedBefore, int gen0Before, int gen1Before, int gen2Before, Process process)
    {
        var integrity = CompareRows(expectedRows, importedRows, importErrorCount);
        var success = importSucceeded && integrity.RoundTrip;
        process.Refresh();
        return new Sample
        {
            Provider = provider,
            Mode = mode,
            Operation = operation,
            Repetition = repetition,
            Status = success ? "measured" : "failed",
            ElapsedMilliseconds = elapsedMilliseconds,
            ExportElapsedMilliseconds = exportElapsedMilliseconds,
            ImportElapsedMilliseconds = importElapsedMilliseconds,
            AllocatedBytes = Math.Max(0, GC.GetTotalAllocatedBytes(true) - allocatedBefore),
            Gen0Collections = GC.CollectionCount(0) - gen0Before,
            Gen1Collections = GC.CollectionCount(1) - gen1Before,
            Gen2Collections = GC.CollectionCount(2) - gen2Before,
            PeakWorkingSetBytes = process.PeakWorkingSet64,
            OutputBytes = bytes.Length,
            OutputIdentity = HashBytes(bytes),
            FixtureIdentity = fixture.Identity,
            FixtureBytes = fixture.Bytes.Length,
            ExportedRows = fixture.RowCount,
            ImportedRows = importedRows.Count,
            ImportErrorCount = importErrorCount,
            RoundTrip = integrity.RoundTrip,
            OutputStructureValid = HasWorkbookStructure(bytes),
            Integrity = integrity,
            Failure = success ? null : new Failure
            {
                Type = importSucceeded ? "RoundTripValidationFailure" : "ImportContractFailure",
                Message = importSucceeded
                    ? "导入完成但固定字段完整性校验失败。"
                    : "导入返回结构化错误。"
            }
        };
    }

    /// <summary>
    /// 创建仅导出样本的初始完整性记录。
    /// </summary>
    /// <param name="expectedRows">预期导出行数。</param>
    /// <returns>尚未执行导入校验的完整性记录。</returns>
    private static Integrity CreateExportIntegrity(int expectedRows) => new()
    {
        RoundTrip = false,
        ExpectedRows = expectedRows,
        ActualRows = 0,
        ExpectedColumns = ColumnCount,
        ImportErrorCount = 0,
        Mismatch = null
    };

    /// <summary>
    /// 重置 GC 统计并采集测量起始状态。
    /// </summary>
    /// <param name="allocatedBefore">返回测量开始时的累计分配字节数。</param>
    /// <param name="gen0Before">返回测量开始时的第 0 代 GC 次数。</param>
    /// <param name="gen1Before">返回测量开始时的第 1 代 GC 次数。</param>
    /// <param name="gen2Before">返回测量开始时的第 2 代 GC 次数。</param>
    /// <param name="process">返回用于读取工作集的当前进程对象。</param>
    private static void PrepareMeasurement(out long allocatedBefore, out int gen0Before, out int gen1Before,
        out int gen2Before, out Process process)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        process = Process.GetCurrentProcess();
        process.Refresh();
        allocatedBefore = GC.GetTotalAllocatedBytes(true);
        gen0Before = GC.CollectionCount(0);
        gen1Before = GC.CollectionCount(1);
        gen2Before = GC.CollectionCount(2);
    }

    /// <summary>
    /// 比较预期行与实际行并生成完整性结果。
    /// </summary>
    /// <param name="expected">预期数据行。</param>
    /// <param name="actual">实际导入的数据行。</param>
    /// <param name="importErrorCount">导入阶段报告的错误数量。</param>
    /// <returns>包含首个差异和行数统计的完整性记录。</returns>
    private static Integrity CompareRows(IReadOnlyList<MaterializationRow> expected,
        IReadOnlyList<MaterializationRow> actual, int importErrorCount)
    {
        var mismatch = string.Empty;
        var count = Math.Min(expected.Count, actual.Count);
        for (var index = 0; index < count; index++)
        {
            if (!RowsEqual(expected[index], actual[index]))
            {
                mismatch = $"row={index};expectedCode={expected[index].Code};actualCode={actual[index].Code}";
                break;
            }
        }

        if (string.IsNullOrEmpty(mismatch) && expected.Count != actual.Count)
            mismatch = $"row-count;expected={expected.Count};actual={actual.Count}";
        return new Integrity
        {
            RoundTrip = string.IsNullOrEmpty(mismatch) && importErrorCount == 0,
            ExpectedRows = expected.Count,
            ActualRows = actual.Count,
            ExpectedColumns = ColumnCount,
            ImportErrorCount = importErrorCount,
            Mismatch = string.IsNullOrEmpty(mismatch) ? null : mismatch
        };
    }

    /// <summary>
    /// 比较两行的固定字段是否完全相等。
    /// </summary>
    /// <param name="expected">预期数据行。</param>
    /// <param name="actual">实际数据行。</param>
    /// <returns>全部固定字段相等时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool RowsEqual(MaterializationRow expected, MaterializationRow actual) =>
        string.Equals(expected.Code, actual.Code, StringComparison.Ordinal)
        && expected.Quantity == actual.Quantity
        && expected.Amount == actual.Amount
        && expected.OccurredAt == actual.OccurredAt
        && expected.Enabled == actual.Enabled
        && string.Equals(expected.Description, actual.Description, StringComparison.Ordinal);

    /// <summary>
    /// 创建带显式列映射的工作簿导出请求。
    /// </summary>
    /// <param name="rows">待导出的数据行。</param>
    /// <param name="mapping">列映射配置。</param>
    /// <returns>工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateExportRequest(IReadOnlyList<MaterializationRow> rows,
        ExcelMappingConfiguration mapping) => ExcelExport.Workbook(workbook =>
            workbook.AddSheet("Data", rows, sheet => sheet.Mapping(mapping)));

    /// <summary>
    /// 创建带显式列映射的工作簿导入请求。
    /// </summary>
    /// <param name="mapping">列映射配置。</param>
    /// <returns>工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<MaterializationWorkbook> CreateImportRequest(
        ExcelMappingConfiguration mapping) => ExcelImport.Workbook<MaterializationWorkbook>(workbook =>
            workbook.Sheet("Data", root => root.Rows, sheet => sheet.Mapping(mapping)));

    /// <summary>
    /// 创建固定数据集使用的列映射配置。
    /// </summary>
    /// <returns>包含六个固定列的映射配置。</returns>
    private static ExcelMappingConfiguration CreateMapping() => new()
    {
        Columns = new List<ExcelColumnConfiguration>
        {
            new() { PropertyName = nameof(MaterializationRow.Code), Title = "Code" },
            new() { PropertyName = nameof(MaterializationRow.Quantity), Title = "Quantity" },
            new() { PropertyName = nameof(MaterializationRow.Amount), Title = "Amount" },
            new() { PropertyName = nameof(MaterializationRow.OccurredAt), Title = "OccurredAt" },
            new() { PropertyName = nameof(MaterializationRow.Enabled), Title = "Enabled" },
            new() { PropertyName = nameof(MaterializationRow.Description), Title = "Description" }
        }
    };

    /// <summary>
    /// 按固定种子生成可重复的数据行。
    /// </summary>
    /// <param name="rowCount">要生成的行数。</param>
    /// <returns>生成的数据行集合。</returns>
    private static IReadOnlyList<MaterializationRow> CreateRows(int rowCount) =>
        Enumerable.Range(0, rowCount).Select(index => new MaterializationRow
        {
            Code = $"MB-{FixedSeed}-{index:D8}",
            Quantity = checked(index * 17 + 3),
            Amount = index * 0.125m + 10.25m,
            OccurredAt = new DateTime(2026, 9, 20, 8, 0, 0, DateTimeKind.Unspecified).AddMinutes(index),
            Enabled = index % 2 == 0,
            Description = $"materialization-binding-{index % 97:D2}"
        }).ToArray();

    /// <summary>
    /// 按 Provider、模式和操作聚合样本统计。
    /// </summary>
    /// <param name="samples">待聚合的测量样本。</param>
    /// <returns>聚合后的统计记录。</returns>
    private static IReadOnlyList<Summary> CreateSummaries(IEnumerable<Sample> samples) => samples
        .GroupBy(sample => new { sample.Provider, sample.Mode, sample.Operation })
        .Select(group => new Summary
        {
            Provider = group.Key.Provider,
            Mode = group.Key.Mode,
            Operation = group.Key.Operation,
            SampleCount = group.Count(),
            MeasuredSampleCount = group.Count(sample => sample.Status == "measured"),
            FailedSampleCount = group.Count(sample => sample.Status == "failed"),
            MedianElapsedMilliseconds = Median(group.Select(sample => sample.ElapsedMilliseconds)),
            MedianAllocatedBytes = Median(group.Select(sample => sample.AllocatedBytes)),
            MedianPeakWorkingSetBytes = Median(group.Select(sample => sample.PeakWorkingSetBytes)),
            MedianOutputBytes = Median(group.Select(sample => sample.OutputBytes.HasValue
                ? (long?)sample.OutputBytes.Value
                : null))
        }).ToArray();

    /// <summary>
    /// 检查字节是否包含最小的 XLSX 工作簿结构。
    /// </summary>
    /// <param name="bytes">待检查的工作簿字节。</param>
    /// <returns>包含工作簿和首个工作表 XML 时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool HasWorkbookStructure(byte[] bytes)
    {
        if (bytes.Length < 4 || bytes[0] != 0x50 || bytes[1] != 0x4B)
            return false;
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
            var names = archive.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return names.Contains("xl/workbook.xml") && names.Contains("xl/worksheets/sheet1.xml");
        }
        catch (InvalidDataException)
        {
            return false;
        }
    }

    /// <summary>
    /// 计算固定数据集的规范化身份哈希。
    /// </summary>
    /// <param name="rows">待哈希的数据行。</param>
    /// <returns>数据行的 SHA-256 十六进制哈希。</returns>
    private static string HashRows(IReadOnlyList<MaterializationRow> rows)
    {
        var canonical = new StringBuilder(rows.Count * 96);
        foreach (var row in rows)
        {
            canonical.Append(row.Code).Append('|')
                .Append(row.Quantity).Append('|')
                .Append(row.Amount.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append('|')
                .Append(row.OccurredAt.ToString("O", System.Globalization.CultureInfo.InvariantCulture)).Append('|')
                .Append(row.Enabled ? '1' : '0').Append('|')
                .Append(row.Description).Append('\n');
        }
        return HashBytes(Encoding.UTF8.GetBytes(canonical.ToString()));
    }

    /// <summary>
    /// 计算字节数组的 SHA-256 哈希。
    /// </summary>
    /// <param name="bytes">待哈希的字节数组。</param>
    /// <returns>大写十六进制 SHA-256 哈希。</returns>
    private static string HashBytes(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes));

    /// <summary>
    /// 计算可空双精度数值序列的中位数。
    /// </summary>
    /// <param name="values">待计算的数值序列。</param>
    /// <returns>非空值的中位数；没有非空值时返回 <see langword="null"/>。</returns>
    private static double? Median(IEnumerable<double?> values)
    {
        var ordered = values.Where(value => value.HasValue).Select(value => value!.Value).OrderBy(value => value)
            .ToArray();
        if (ordered.Length == 0)
            return null;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 1 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2d;
    }

    /// <summary>
    /// 计算可空长整数序列的中位数。
    /// </summary>
    /// <param name="values">待计算的数值序列。</param>
    /// <returns>非空值的中位数；没有非空值时返回 <see langword="null"/>。</returns>
    private static long? Median(IEnumerable<long?> values)
    {
        var ordered = values.Where(value => value.HasValue).Select(value => value!.Value).OrderBy(value => value)
            .ToArray();
        if (ordered.Length == 0)
            return null;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 1 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2;
    }

    /// <summary>
    /// 校验序列化探针文档的基本结构和样本数量。
    /// </summary>
    /// <param name="json">待校验的 JSON 文本。</param>
    /// <param name="document">用于推导预期结构的探针文档。</param>
    private static void ValidateSerializedArtifact(string json, ProbeDocument document)
    {
        using var parsed = JsonDocument.Parse(json);
        var root = parsed.RootElement;
        if (root.GetProperty("schema").GetInt32() != SchemaVersion
            || root.GetProperty("kind").GetString() != "materialization-binding-probe"
            || root.GetProperty("rowCount").GetInt32() != document.RowCount
            || root.GetProperty("columnCount").GetInt32() != ColumnCount)
            throw new InvalidOperationException("materialization-binding 探针 JSON schema 自检失败。 ");

        var serializedSamples = root.GetProperty("samples");
        var expectedSampleCount = document.Repetitions * document.Providers.Count * document.Modes.Count
            * document.Operations.Count;
        if (serializedSamples.GetArrayLength() != expectedSampleCount)
            throw new InvalidOperationException("materialization-binding 探针样本数量自检失败。 ");
        foreach (var sample in serializedSamples.EnumerateArray())
        {
            _ = sample.GetProperty("provider").GetString()
                ?? throw new InvalidOperationException("materialization-binding 样本缺少 provider。 ");
            _ = sample.GetProperty("mode").GetString()
                ?? throw new InvalidOperationException("materialization-binding 样本缺少 mode。 ");
            _ = sample.GetProperty("operation").GetString()
                ?? throw new InvalidOperationException("materialization-binding 样本缺少 operation。 ");
            _ = sample.GetProperty("status").GetString()
                ?? throw new InvalidOperationException("materialization-binding 样本缺少 status。 ");
            _ = sample.GetProperty("fixtureIdentity").GetString()
                ?? throw new InvalidOperationException("materialization-binding 样本缺少 fixture identity。 ");
            _ = sample.GetProperty("roundTrip");
            _ = sample.GetProperty("outputStructureValid");
            _ = sample.GetProperty("integrity");
        }
    }

    /// <summary>
    /// 获取参与探针运行的程序集身份哈希。
    /// </summary>
    /// <returns>按文件名索引的程序集哈希。</returns>
    private static IReadOnlyDictionary<string, string> GetCandidateIdentity()
    {
        var types = new[]
        {
            typeof(MaterializationBindingProbe),
            typeof(Program),
            typeof(ExcelMappingConfigurationLoader),
            typeof(IExcelExporter),
            typeof(Bing.Offices.Exports.NpoiExcelExporter),
            typeof(Bing.Offices.Exports.MiniExcelExcelExporter)
        };
        return types.Select(type => type.Assembly.Location).Distinct(StringComparer.OrdinalIgnoreCase)
            .ToDictionary(path => Path.GetFileName(path), path => HashAssembly(path), StringComparer.Ordinal);
    }

    /// <summary>
    /// 读取程序集并计算其身份哈希。
    /// </summary>
    /// <param name="path">程序集文件路径。</param>
    /// <returns>程序集文件的 SHA-256 十六进制哈希。</returns>
    private static string HashAssembly(string path) => HashBytes(File.ReadAllBytes(path));

    /// <summary>
    /// 探针结果文档模型。
    /// </summary>
    private sealed class ProbeDocument
    {
        /// <summary>
        /// 获取或设置文档类型。
        /// </summary>
        public string Kind { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置文档版本。
        /// </summary>
        public int Schema { get; set; }
        /// <summary>
        /// 获取或设置生成时间。
        /// </summary>
        public DateTimeOffset GeneratedUtc { get; set; }
        /// <summary>
        /// 获取或设置运行时框架描述。
        /// </summary>
        public string Framework { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置运行系统描述。
        /// </summary>
        public string Os { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置进程标识。
        /// </summary>
        public int ProcessId { get; set; }
        /// <summary>
        /// 获取或设置处理器数量。
        /// </summary>
        public int ProcessorCount { get; set; }
        /// <summary>
        /// 获取或设置数据集种子。
        /// </summary>
        public int Seed { get; set; }
        /// <summary>
        /// 获取或设置数据行数。
        /// </summary>
        public int RowCount { get; set; }
        /// <summary>
        /// 获取或设置固定列数。
        /// </summary>
        public int ColumnCount { get; set; }
        /// <summary>
        /// 获取或设置每个场景的重复次数。
        /// </summary>
        public int Repetitions { get; set; }
        /// <summary>
        /// 获取或设置参与测试的 Provider 列表。
        /// </summary>
        public IReadOnlyList<string> Providers { get; set; } = Array.Empty<string>();
        /// <summary>
        /// 获取或设置参与测试的执行模式列表。
        /// </summary>
        public IReadOnlyList<string> Modes { get; set; } = Array.Empty<string>();
        /// <summary>
        /// 获取或设置参与测试的操作列表。
        /// </summary>
        public IReadOnlyList<string> Operations { get; set; } = Array.Empty<string>();
        /// <summary>
        /// 获取或设置探针工作负载描述。
        /// </summary>
        public string Workload { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置固定输入数据集身份。
        /// </summary>
        public string DatasetIdentity { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置候选程序集身份哈希。
        /// </summary>
        public IReadOnlyDictionary<string, string> CandidateIdentity { get; set; } =
            new Dictionary<string, string>();
        /// <summary>
        /// 获取或设置基线状态。
        /// </summary>
        public string BaselineStatus { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置基线状态说明。
        /// </summary>
        public string BaselineReason { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置各 Provider 输入夹具身份。
        /// </summary>
        public IReadOnlyDictionary<string, string> FixtureIdentities { get; set; } =
            new Dictionary<string, string>();
        /// <summary>
        /// 获取或设置性能阈值审批状态。
        /// </summary>
        public string ThresholdStatus { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置证据采集状态。
        /// </summary>
        public string EvidenceStatus { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置原始测量样本。
        /// </summary>
        public IReadOnlyList<Sample> Samples { get; set; } = Array.Empty<Sample>();
        /// <summary>
        /// 获取或设置聚合统计结果。
        /// </summary>
        public IReadOnlyList<Summary> Summaries { get; set; } = Array.Empty<Summary>();
    }

    /// <summary>
    /// 单次 Provider 测量结果。
    /// </summary>
    private sealed class Sample
    {
        /// <summary>
        /// 获取或设置 Provider 名称。
        /// </summary>
        public string Provider { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置执行模式。
        /// </summary>
        public string Mode { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置操作类型。
        /// </summary>
        public string Operation { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置测量重复序号。
        /// </summary>
        public int Repetition { get; set; }
        /// <summary>
        /// 获取或设置样本状态。
        /// </summary>
        public string Status { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置总耗时，单位为毫秒。
        /// </summary>
        public double? ElapsedMilliseconds { get; set; }
        /// <summary>
        /// 获取或设置导出耗时，单位为毫秒。
        /// </summary>
        public double? ExportElapsedMilliseconds { get; set; }
        /// <summary>
        /// 获取或设置导入耗时，单位为毫秒。
        /// </summary>
        public double? ImportElapsedMilliseconds { get; set; }
        /// <summary>
        /// 获取或设置托管堆分配字节数。
        /// </summary>
        public long? AllocatedBytes { get; set; }
        /// <summary>
        /// 获取或设置第 0 代 GC 次数。
        /// </summary>
        public int? Gen0Collections { get; set; }
        /// <summary>
        /// 获取或设置第 1 代 GC 次数。
        /// </summary>
        public int? Gen1Collections { get; set; }
        /// <summary>
        /// 获取或设置第 2 代 GC 次数。
        /// </summary>
        public int? Gen2Collections { get; set; }
        /// <summary>
        /// 获取或设置进程峰值工作集字节数。
        /// </summary>
        public long? PeakWorkingSetBytes { get; set; }
        /// <summary>
        /// 获取或设置输出工作簿字节数。
        /// </summary>
        public int? OutputBytes { get; set; }
        /// <summary>
        /// 获取或设置输出工作簿身份哈希。
        /// </summary>
        public string? OutputIdentity { get; set; }
        /// <summary>
        /// 获取或设置输入夹具身份哈希。
        /// </summary>
        public string FixtureIdentity { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置输入夹具字节数。
        /// </summary>
        public int? FixtureBytes { get; set; }
        /// <summary>
        /// 获取或设置导出的行数。
        /// </summary>
        public int? ExportedRows { get; set; }
        /// <summary>
        /// 获取或设置导入的行数。
        /// </summary>
        public int? ImportedRows { get; set; }
        /// <summary>
        /// 获取或设置导入错误数量。
        /// </summary>
        public int? ImportErrorCount { get; set; }
        /// <summary>
        /// 获取或设置往返校验是否成功。
        /// </summary>
        public bool? RoundTrip { get; set; }
        /// <summary>
        /// 获取或设置输出结构是否有效。
        /// </summary>
        public bool? OutputStructureValid { get; set; }
        /// <summary>
        /// 获取或设置行完整性结果。
        /// </summary>
        public Integrity? Integrity { get; set; }
        /// <summary>
        /// 获取或设置失败信息。
        /// </summary>
        public Failure? Failure { get; set; }

        /// <summary>
        /// 创建失败状态的测量样本。
        /// </summary>
        /// <param name="provider">Provider 名称。</param>
        /// <param name="mode">执行模式。</param>
        /// <param name="operation">操作类型。</param>
        /// <param name="repetition">测量重复序号。</param>
        /// <param name="exception">导致失败的异常。</param>
        /// <param name="fixtureIdentity">输入夹具身份哈希。</param>
        /// <param name="fixtureBytes">输入夹具字节数。</param>
        /// <returns>包含失败信息的样本。</returns>
        public static Sample Failed(string provider, string mode, string operation, int repetition,
            Exception exception, string fixtureIdentity = "", int fixtureBytes = 0) => new()
        {
            Provider = provider,
            Mode = mode,
            Operation = operation,
            Repetition = repetition,
            Status = "failed",
            RoundTrip = false,
            FixtureIdentity = fixtureIdentity,
            FixtureBytes = fixtureBytes,
            OutputStructureValid = false,
            Integrity = new Integrity
            {
                ExpectedRows = 0,
                ActualRows = 0,
                ExpectedColumns = ColumnCount,
                ImportErrorCount = 0,
                RoundTrip = false,
                Mismatch = null
            },
            Failure = new Failure
            {
                Type = exception.GetType().FullName ?? exception.GetType().Name,
                Message = exception.Message,
                StackTrace = exception.StackTrace
            }
        };
    }

    /// <summary>
    /// 导入行与预期数据的完整性结果。
    /// </summary>
    private sealed class Integrity
    {
        /// <summary>
        /// 获取或设置往返校验是否成功。
        /// </summary>
        public bool RoundTrip { get; set; }
        /// <summary>
        /// 获取或设置预期行数。
        /// </summary>
        public int ExpectedRows { get; set; }
        /// <summary>
        /// 获取或设置实际行数。
        /// </summary>
        public int ActualRows { get; set; }
        /// <summary>
        /// 获取或设置预期列数。
        /// </summary>
        public int ExpectedColumns { get; set; }
        /// <summary>
        /// 获取或设置导入错误数量。
        /// </summary>
        public int ImportErrorCount { get; set; }
        /// <summary>
        /// 获取或设置首个差异描述；没有差异时为 <see langword="null"/>。
        /// </summary>
        public string? Mismatch { get; set; }
    }

    /// <summary>
    /// 探针失败信息。
    /// </summary>
    private sealed class Failure
    {
        /// <summary>
        /// 获取或设置异常类型名称。
        /// </summary>
        public string Type { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置异常消息。
        /// </summary>
        public string Message { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置异常堆栈；不可用时为 <see langword="null"/>。
        /// </summary>
        public string? StackTrace { get; set; }
    }

    /// <summary>
    /// 按 Provider、模式和操作聚合的测量统计。
    /// </summary>
    private sealed class Summary
    {
        /// <summary>
        /// 获取或设置 Provider 名称。
        /// </summary>
        public string Provider { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置执行模式。
        /// </summary>
        public string Mode { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置操作类型。
        /// </summary>
        public string Operation { get; set; } = string.Empty;
        /// <summary>
        /// 获取或设置样本总数。
        /// </summary>
        public int SampleCount { get; set; }
        /// <summary>
        /// 获取或设置成功测量样本数。
        /// </summary>
        public int MeasuredSampleCount { get; set; }
        /// <summary>
        /// 获取或设置失败样本数。
        /// </summary>
        public int FailedSampleCount { get; set; }
        /// <summary>
        /// 获取或设置耗时中位数，单位为毫秒。
        /// </summary>
        public double? MedianElapsedMilliseconds { get; set; }
        /// <summary>
        /// 获取或设置分配字节中位数。
        /// </summary>
        public long? MedianAllocatedBytes { get; set; }
        /// <summary>
        /// 获取或设置峰值工作集字节中位数。
        /// </summary>
        public long? MedianPeakWorkingSetBytes { get; set; }
        /// <summary>
        /// 获取或设置输出字节中位数。
        /// </summary>
        public long? MedianOutputBytes { get; set; }
    }

    /// <summary>
    /// Provider 使用的固定输入工作簿夹具。
    /// </summary>
    private sealed class Fixture
    {
        /// <summary>
        /// 初始化工作簿夹具。
        /// </summary>
        /// <param name="provider">Provider 名称。</param>
        /// <param name="bytes">工作簿字节。</param>
        /// <param name="identity">工作簿身份哈希。</param>
        /// <param name="rowCount">工作簿中的数据行数。</param>
        public Fixture(string provider, byte[] bytes, string identity, int rowCount)
        {
            Provider = provider;
            Bytes = bytes;
            Identity = identity;
            RowCount = rowCount;
        }

        /// <summary>
        /// 获取 Provider 名称。
        /// </summary>
        public string Provider { get; }

        /// <summary>
        /// 获取工作簿字节。
        /// </summary>
        public byte[] Bytes { get; }

        /// <summary>
        /// 获取工作簿身份哈希。
        /// </summary>
        public string Identity { get; }

        /// <summary>
        /// 获取工作簿中的数据行数。
        /// </summary>
        public int RowCount { get; }
    }

    /// <summary>
    /// materialization 探针的导入工作簿模型。
    /// </summary>
    private sealed class MaterializationWorkbook
    {
        /// <summary>
        /// 获取导入的数据行集合。
        /// </summary>
        public List<MaterializationRow> Rows { get; } = new();
    }

    /// <summary>
    /// materialization 探针的固定数据行模型。
    /// </summary>
    private sealed class MaterializationRow
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
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
        /// <summary>
        /// 获取或设置发生时间。
        /// </summary>
        public DateTime OccurredAt { get; set; }
        /// <summary>
        /// 获取或设置启用状态。
        /// </summary>
        public bool Enabled { get; set; }
        /// <summary>
        /// 获取或设置描述文本。
        /// </summary>
        public string Description { get; set; } = string.Empty;
    }
}
