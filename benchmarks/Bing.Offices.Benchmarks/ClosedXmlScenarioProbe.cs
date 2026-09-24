using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Styles;
using ClosedXML.Excel;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// 为 ClosedXML 单独采集导入、导出、文件和富 XLSX 场景。
/// </summary>
internal static class ClosedXmlScenarioProbe
{
    /// <summary>
    /// 支持的 ClosedXML 基准场景名称。
    /// </summary>
    private static readonly string[] ScenarioNames =
    {
        "export", "import", "file", "multisheet", "style", "template", "formula"
    };

    /// <summary>
    /// 执行指定的 ClosedXML 基准场景并写入 JSONL 结果。
    /// </summary>
    /// <param name="artifactPath">结果文件路径。</param>
    /// <param name="rowCount">每个场景生成或导入的数据行数。</param>
    /// <param name="repetitions">每种操作模式的测量次数，至少为 3。</param>
    /// <param name="scenario">场景名称，或表示运行全部场景的 <c>all</c>。</param>
    public static async Task RunAsync(string artifactPath, int rowCount, int repetitions, string scenario)
    {
        if (rowCount < 1)
            throw new ArgumentOutOfRangeException(nameof(rowCount));
        if (repetitions < 3)
            throw new ArgumentOutOfRangeException(nameof(repetitions), "正式场景至少需要三次重复。");

        var scenarios = string.Equals(scenario, "all", StringComparison.OrdinalIgnoreCase)
            ? ScenarioNames
            : new[] { scenario.ToLowerInvariant() };
        foreach (var item in scenarios)
        {
            if (!ScenarioNames.Contains(item, StringComparer.OrdinalIgnoreCase))
                throw new ArgumentException($"未知 ClosedXML 场景: {item}。", nameof(scenario));
        }

        var fullPath = Path.GetFullPath(artifactPath);
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
        using var writer = new StreamWriter(fullPath, false, new UTF8Encoding(false));
        await writer.WriteLineAsync(JsonSerializer.Serialize(new
        {
            kind = "closedxml-scenario-probe",
            schema = 1,
            generatedUtc = DateTimeOffset.UtcNow,
            provider = "closedxml",
            framework = RuntimeInformation.FrameworkDescription,
            os = RuntimeInformation.OSDescription,
            rowCount,
            repetitions,
            requestedScenario = scenario,
            scenarios,
            beforeData = "NOT_APPLICABLE",
            workload = "independent export/import/file and rich XLSX scenarios"
        })).ConfigureAwait(false);

        var exporter = new ClosedXmlExcelExporter();
        var importer = new ClosedXmlExcelImporter();
        foreach (var currentScenario in scenarios)
        {
            var workload = new ScenarioWorkload(currentScenario, rowCount, exporter);
            foreach (var mode in new[] { "sync", "async" })
            {
                var samples = new List<Sample>(repetitions);
                await MeasureAsync(workload, exporter, importer, mode, 0, false).ConfigureAwait(false);
                for (var repetition = 1; repetition <= repetitions; repetition++)
                {
                    var sample = await MeasureAsync(workload, exporter, importer, mode, repetition, true)
                        .ConfigureAwait(false);
                    samples.Add(sample);
                    await writer.WriteLineAsync(JsonSerializer.Serialize(sample)).ConfigureAwait(false);
                    await writer.FlushAsync().ConfigureAwait(false);
                }

                await writer.WriteLineAsync(JsonSerializer.Serialize(new
                {
                    kind = "closedxml-scenario-summary",
                    provider = "closedxml",
                    scenario = currentScenario,
                    operation = workload.Operation,
                    mode,
                    rowCount,
                    repetitions,
                    medianElapsedMilliseconds = Median(samples.Select(x => x.ElapsedMilliseconds)),
                    medianAllocatedBytes = Median(samples.Select(x => x.AllocatedBytes)),
                    medianOutputBytes = Median(samples.Select(x => (long)x.OutputBytes)),
                    medianRowsPerSecond = Median(samples.Select(x => x.RowsPerSecond)),
                    medianPeakWorkingSetBytes = Median(samples.Select(x => x.PeakWorkingSetBytes)),
                    beforeData = "NOT_APPLICABLE"
                })).ConfigureAwait(false);
                await writer.FlushAsync().ConfigureAwait(false);
            }
        }

        Console.WriteLine($"CLOSEDXML_SCENARIO_PROBE artifact={fullPath} scenarios={scenarios.Length} "
            + $"rows={rowCount} repetitions={repetitions} status=passed");
    }

    /// <summary>
    /// 测量一次场景操作并收集运行时指标。
    /// </summary>
    /// <param name="workload">待执行的场景工作负载。</param>
    /// <param name="exporter">用于导出的 ClosedXML 导出器。</param>
    /// <param name="importer">用于导入的 ClosedXML 导入器。</param>
    /// <param name="mode">执行模式，取值为 <c>sync</c> 或 <c>async</c>。</param>
    /// <param name="repetition">当前测量序号。</param>
    /// <param name="capture">是否将本次测量标记为正式样本。</param>
    private static async Task<Sample> MeasureAsync(ScenarioWorkload workload,
        ClosedXmlExcelExporter exporter, ClosedXmlExcelImporter importer, string mode, int repetition,
        bool capture)
    {
        GC.Collect(2, GCCollectionMode.Forced, true, true);
        var process = Process.GetCurrentProcess();
        process.Refresh();
        var allocatedBefore = GC.GetTotalAllocatedBytes(true);
        var gen0Before = GC.CollectionCount(0);
        var gen1Before = GC.CollectionCount(1);
        var gen2Before = GC.CollectionCount(2);
        var stopwatch = Stopwatch.StartNew();
        var outputBytes = 0;
        var inputBytes = 0;
        var importedRows = 0;

        if (workload.Operation == "export")
        {
            using var output = new MemoryStream();
            var request = workload.CreateExportRequest();
            if (mode == "async")
                await exporter.ExportAsync(request, output).ConfigureAwait(false);
            else
                exporter.Export(request, output);
            outputBytes = checked((int)output.Length);
        }
        else if (workload.Operation == "import")
        {
            using var input = new MemoryStream(workload.Payload!, writable: false);
            inputBytes = workload.Payload!.Length;
            var request = workload.CreateImportRequest();
            var result = mode == "async"
                ? await importer.ImportAsync(input, request).ConfigureAwait(false)
                : importer.Import(input, request);
            importedRows = result.Workbook.Rows.Count;
        }
        else if (workload.Operation == "file")
        {
            var path = Path.Combine(Path.GetTempPath(), $"bing-offices-cx-benchmark-{Guid.NewGuid():N}.xlsx");
            try
            {
                var request = workload.CreateExportRequest();
                if (mode == "async")
                    await exporter.ExportToFileAsync(request, path).ConfigureAwait(false);
                else
                    exporter.ExportToFile(request, path);
                outputBytes = checked((int)new FileInfo(path).Length);
                await using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                    64 * 1024, useAsync: mode == "async");
                inputBytes = checked((int)input.Length);
                var importRequest = workload.CreateImportRequest();
                var result = mode == "async"
                    ? await importer.ImportAsync(input, importRequest).ConfigureAwait(false)
                    : importer.Import(input, importRequest);
                importedRows = result.Workbook.Rows.Count;
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }
        else
        {
            using var output = new MemoryStream();
            var exportRequest = workload.CreateExportRequest();
            if (mode == "async")
                await exporter.ExportAsync(exportRequest, output).ConfigureAwait(false);
            else
                exporter.Export(exportRequest, output);
            outputBytes = checked((int)output.Length);
            output.Position = 0;
            var importRequest = workload.CreateImportRequest();
            var result = mode == "async"
                ? await importer.ImportAsync(output, importRequest).ConfigureAwait(false)
                : importer.Import(output, importRequest);
            importedRows = result.Workbook.Rows.Count;
        }

        stopwatch.Stop();
        if (workload.Operation != "export" && importedRows != workload.RowCount)
            throw new InvalidOperationException($"ClosedXML {workload.Scenario} 导入行数不匹配: {importedRows}。");
        process.Refresh();
        return new Sample
        {
            Kind = "closedxml-scenario-sample",
            Provider = "closedxml",
            Scenario = workload.Scenario,
            Operation = workload.Operation,
            Mode = mode,
            Repetition = repetition,
            RowCount = workload.RowCount,
            ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds,
            AllocatedBytes = GC.GetTotalAllocatedBytes(true) - allocatedBefore,
            OutputBytes = outputBytes,
            InputBytes = inputBytes,
            RowsPerSecond = workload.Operation == "export"
                ? workload.RowCount / stopwatch.Elapsed.TotalSeconds
                : importedRows / stopwatch.Elapsed.TotalSeconds,
            Gen0Collections = GC.CollectionCount(0) - gen0Before,
            Gen1Collections = GC.CollectionCount(1) - gen1Before,
            Gen2Collections = GC.CollectionCount(2) - gen2Before,
            PeakWorkingSetBytes = process.PeakWorkingSet64,
            Captured = capture,
            BeforeData = "NOT_APPLICABLE"
        };
    }

    /// <summary>
    /// 计算双精度数值序列的中位数。
    /// </summary>
    /// <param name="values">待计算的数值序列。</param>
    /// <returns>序列的中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0)
            return 0;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 1 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2;
    }

    /// <summary>
    /// 计算长整数序列的中位数。
    /// </summary>
    /// <param name="values">待计算的数值序列。</param>
    /// <returns>序列的中位数；空序列返回 0。</returns>
    private static double Median(IEnumerable<long> values) => Median(values.Select(value => (double)value));

    /// <summary>
    /// 创建包含公式、批注和样式的模板工作簿字节。
    /// </summary>
    /// <returns>模板工作簿的 XLSX 字节。</returns>
    private static byte[] CreateTemplateBytes()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("A1").Value = "Template";
        sheet.Cell("H1").FormulaA1 = "=1+1";
        sheet.Cell("H1").CreateComment().AddText("benchmark template");
        sheet.Range("A1:H1").Style.Font.Bold = true;
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// ClosedXML 场景执行所需的固定工作负载。
    /// </summary>
    private sealed class ScenarioWorkload
    {
        /// <summary>
        /// 初始化场景工作负载。
        /// </summary>
        /// <param name="scenario">场景名称。</param>
        /// <param name="rowCount">工作负载的数据行数。</param>
        /// <param name="exporter">用于生成导入场景输入数据的导出器。</param>
        public ScenarioWorkload(string scenario, int rowCount, ClosedXmlExcelExporter exporter)
        {
            Scenario = scenario;
            RowCount = rowCount;
            var rows = Enumerable.Range(0, rowCount).Select(index => new ScenarioRow
            {
                Code = $"SCENARIO-{index:D7}",
                Quantity = index,
                Description = $"scenario-row-{index}",
                Formula = string.Equals(scenario, "formula", StringComparison.OrdinalIgnoreCase)
                    ? "=1+1"
                    : string.Empty
            }).ToArray();
            var templateBytes = scenario == "template" ? CreateTemplateBytes() : null;
            CreateExportRequest = () => CreateRequest(scenario, rows, templateBytes);
            CreateImportRequest = () => ExcelImport.Workbook<ScenarioWorkbook>(workbook =>
                workbook.Sheet("Data", root => root.Rows));
            Operation = scenario switch
            {
                "export" => "export",
                "import" => "import",
                "file" => "file",
                _ => "roundtrip"
            };
            if (Operation == "import")
            {
                using var payload = new MemoryStream();
                exporter.Export(CreateExportRequest(), payload);
                Payload = payload.ToArray();
            }
        }

        /// <summary>
        /// 获取场景名称。
        /// </summary>
        public string Scenario { get; }

        /// <summary>
        /// 获取工作负载的数据行数。
        /// </summary>
        public int RowCount { get; }

        /// <summary>
        /// 获取场景执行的操作类型。
        /// </summary>
        public string Operation { get; }

        /// <summary>
        /// 获取导入场景使用的预生成工作簿字节。
        /// </summary>
        public byte[]? Payload { get; }

        /// <summary>
        /// 获取创建导出请求的工厂。
        /// </summary>
        public Func<ExcelWorkbookExportRequest> CreateExportRequest { get; }

        /// <summary>
        /// 获取创建导入请求的工厂。
        /// </summary>
        public Func<ExcelWorkbookImportRequest<ScenarioWorkbook>> CreateImportRequest { get; }
    }

    /// <summary>
    /// 按场景构造导出请求。
    /// </summary>
    /// <param name="scenario">场景名称。</param>
    /// <param name="rows">待写入的数据行。</param>
    /// <param name="templateBytes">模板场景使用的模板字节；其他场景可为 <see langword="null"/>。</param>
    /// <returns>与场景匹配的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(string scenario, ScenarioRow[] rows,
        byte[]? templateBytes)
    {
        return scenario switch
        {
            "multisheet" => ExcelExport.Workbook(workbook => workbook
                .AddSheet("Data", rows)
                .AddSheet("Archive", rows)),
            "style" => ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows,
                sheet => sheet.HeaderStyle(new ExcelCellStyle { Bold = true, WrapText = true })
                    .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00", WrapText = true })
                    .ColumnWidth(new ExcelColumnWidthOptions
                    {
                        Mode = ExcelColumnWidthMode.Fixed,
                        FixedWidth = 18
                    }))),
            "template" => ExcelExport.Workbook(workbook => workbook
                .UseTemplate(new MemoryStream(templateBytes!), leaveOpen: false)
                .AddSheet("Data", rows)),
            "formula" => ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows)),
            _ => ExcelExport.Workbook(workbook => workbook.AddSheet("Data", rows))
        };
    }

    /// <summary>
    /// ClosedXML 场景使用的示例数据行。
    /// </summary>
    private sealed class ScenarioRow
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

        /// <summary>
        /// 获取或设置公式文本。
        /// </summary>
        public string Formula { get; set; } = string.Empty;
    }

    /// <summary>
    /// ClosedXML 场景导入使用的工作簿模型。
    /// </summary>
    private sealed class ScenarioWorkbook
    {
        /// <summary>
        /// 获取导入的场景行集合。
        /// </summary>
        public List<ScenarioRow> Rows { get; } = new();
    }

    /// <summary>
    /// ClosedXML 场景的一次测量样本。
    /// </summary>
    private sealed class Sample
    {
        /// <summary>
        /// 获取样本记录类型。
        /// </summary>
        public string Kind { get; init; } = string.Empty;

        /// <summary>
        /// 获取提供样本的 Provider 名称。
        /// </summary>
        public string Provider { get; init; } = string.Empty;

        /// <summary>
        /// 获取基准场景名称。
        /// </summary>
        public string Scenario { get; init; } = string.Empty;

        /// <summary>
        /// 获取场景执行的操作类型。
        /// </summary>
        public string Operation { get; init; } = string.Empty;

        /// <summary>
        /// 获取执行模式。
        /// </summary>
        public string Mode { get; init; } = string.Empty;

        /// <summary>
        /// 获取测量重复序号。
        /// </summary>
        public int Repetition { get; init; }

        /// <summary>
        /// 获取本次测量处理的数据行数。
        /// </summary>
        public int RowCount { get; init; }

        /// <summary>
        /// 获取耗时，单位为毫秒。
        /// </summary>
        public double ElapsedMilliseconds { get; init; }

        /// <summary>
        /// 获取托管堆分配字节数。
        /// </summary>
        public long AllocatedBytes { get; init; }

        /// <summary>
        /// 获取输出工作簿字节数。
        /// </summary>
        public int OutputBytes { get; init; }

        /// <summary>
        /// 获取输入工作簿字节数。
        /// </summary>
        public int InputBytes { get; init; }

        /// <summary>
        /// 获取处理吞吐量，单位为行每秒。
        /// </summary>
        public double RowsPerSecond { get; init; }

        /// <summary>
        /// 获取测量期间的第 0 代 GC 次数。
        /// </summary>
        public int Gen0Collections { get; init; }

        /// <summary>
        /// 获取测量期间的第 1 代 GC 次数。
        /// </summary>
        public int Gen1Collections { get; init; }

        /// <summary>
        /// 获取测量期间的第 2 代 GC 次数。
        /// </summary>
        public int Gen2Collections { get; init; }

        /// <summary>
        /// 获取进程峰值工作集字节数。
        /// </summary>
        public long PeakWorkingSetBytes { get; init; }

        /// <summary>
        /// 获取是否为正式采集的样本。
        /// </summary>
        public bool Captured { get; init; }

        /// <summary>
        /// 获取历史对照数据标记。
        /// </summary>
        public string BeforeData { get; init; } = "NOT_APPLICABLE";
    }
}
