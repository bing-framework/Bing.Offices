using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

if (args.Length == 1 && args[0] == "--npoi-font-probe")
    return RunNpoiFontProbe();

var options = ProbeOptions.Parse(args);
if (options == null)
    return 2;

var outputPath = Path.GetFullPath(options.OutputPath);
var outputDirectory = Path.GetDirectoryName(outputPath);
if (string.IsNullOrWhiteSpace(outputDirectory))
    return 2;

Directory.CreateDirectory(outputDirectory);
var request = ExcelExport.Workbook(workbook => workbook.AddSheet(
    "Rows",
    CreateRows(options.Rows),
    sheet => sheet.BodyStyle(new Bing.Offices.Styles.ExcelCellStyle
    {
        NumberFormat = "0.00"
    })));

var exporter = new SpreadCheetahStreamingExcelExporter();
var stopwatch = Stopwatch.StartNew();
try
{
    await exporter.ExportBatchesToFileAsync(
        request,
        outputPath,
        new ExcelStreamingExportOptions { BatchSize = options.BatchSize },
        CancellationToken.None).ConfigureAwait(false);
    stopwatch.Stop();

    var result = new
    {
        status = "success",
        rows = options.Rows,
        batchSize = options.BatchSize,
        elapsedMs = stopwatch.Elapsed.TotalMilliseconds,
        outputBytes = new FileInfo(outputPath).Length,
        peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64
    };
    Console.WriteLine(JsonSerializer.Serialize(result));
    return 0;
}
catch (Exception exception)
{
    stopwatch.Stop();
    Console.Error.WriteLine($"streaming probe failed: {exception.GetType().Name}: {exception.Message}");
    return 1;
}

static int RunNpoiFontProbe()
{
    var npoiRequest = ExcelExport.Workbook(workbook => workbook.AddSheet(
        "FontProbe",
        new[] { new ProbeRow(1, "Font probe 中文", 12.34) },
        sheet => sheet.ColumnWidth(new ExcelColumnWidthOptions
        {
            Mode = ExcelColumnWidthMode.AutoFit
        })));

    object npoiResult;
    var npoiMatchedExpectedFailure = false;
    try
    {
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(npoiRequest, output);
        npoiResult = new
        {
            outcome = "success",
            outputBytes = output.Length
        };
    }
    catch (BingOfficesException exception)
    {
        npoiMatchedExpectedFailure = exception.Code == BingOfficesErrorCode.ExportFailed
            && exception.Provider == "NPOI"
            && exception.Operation == BingOfficesOperation.Export
            && exception.Stage == BingOfficesStage.Write;
        npoiResult = new
        {
            outcome = "failure",
            exceptionType = exception.GetType().FullName,
            exceptionMessage = exception.Message,
            code = exception.Code.ToString(),
            provider = exception.Provider,
            operation = exception.Operation.ToString(),
            stage = exception.Stage.ToString()
        };
    }
    catch (Exception exception)
    {
        npoiResult = new
        {
            outcome = "unexpected-failure",
            exceptionType = exception.GetType().FullName,
            exceptionMessage = exception.Message
        };
    }

    object spreadCheetahResult;
    var spreadCheetahReadbackSucceeded = false;
    try
    {
        var spreadRequest = ExcelExport.Workbook(workbook => workbook.AddSheet(
            "Rows", new[] { new ProbeRow(1, "Font probe 中文", 12.34) }));
        using var output = new MemoryStream();
        new SpreadCheetahStreamingExcelExporter().ExportBatches(spreadRequest, output);
        var bytes = output.ToArray();
        using var workbook = new XSSFWorkbook(new MemoryStream(bytes));
        var row = workbook.GetSheet("Rows").GetRow(1);
        spreadCheetahReadbackSucceeded = row.Cells.Any(cell =>
            cell.CellType == CellType.String && cell.StringCellValue == "Font probe 中文");
        spreadCheetahResult = new
        {
            outcome = spreadCheetahReadbackSucceeded ? "success" : "failure",
            outputBytes = bytes.Length,
            readback = spreadCheetahReadbackSucceeded
        };
    }
    catch (Exception exception)
    {
        spreadCheetahResult = new
        {
            outcome = "failure",
            exceptionType = exception.GetType().FullName,
            exceptionMessage = exception.Message
        };
    }

    Console.WriteLine(JsonSerializer.Serialize(new
    {
        probe = "fontless-provider-comparison",
        npoiAutoFit = npoiResult,
        spreadCheetah = spreadCheetahResult
    }));
    return npoiMatchedExpectedFailure && spreadCheetahReadbackSucceeded ? 0 : 1;
}

static IEnumerable<ProbeRow> CreateRows(int count)
{
    for (var index = 1; index <= count; index++)
        yield return new ProbeRow(index, $"Row {index}", index / 100.0);
}

/// <summary>
/// 流式导出资源探针的数据行。
/// </summary>
sealed class ProbeRow
{
    /// <summary>
    /// 初始化一个 <see cref="ProbeRow"/> 类型的实例。
    /// </summary>
    public ProbeRow()
    {
    }

    /// <summary>
    /// 初始化一个 <see cref="ProbeRow"/> 类型的实例。
    /// </summary>
    /// <param name="id">行标识。</param>
    /// <param name="name">行名称。</param>
    /// <param name="amount">行金额。</param>
    public ProbeRow(int id, string name, double amount)
    {
        Id = id;
        Name = name;
        Amount = amount;
    }

    /// <summary>
    /// 获取行标识。
    /// </summary>
    public int Id { get; }
    /// <summary>
    /// 获取行名称。
    /// </summary>
    public string Name { get; }
    /// <summary>
    /// 获取行金额。
    /// </summary>
    public double Amount { get; }
}

/// <summary>
/// 流式导出资源探针的命令行选项。
/// </summary>
sealed class ProbeOptions
{
    /// <summary>
    /// 初始化一个 <see cref="ProbeOptions"/> 类型的实例。
    /// </summary>
    /// <param name="rows">生成的数据行数。</param>
    /// <param name="batchSize">每批生成的行数。</param>
    /// <param name="outputPath">输出工作簿路径。</param>
    private ProbeOptions(int rows, int batchSize, string outputPath)
    {
        Rows = rows;
        BatchSize = batchSize;
        OutputPath = outputPath;
    }

    /// <summary>
    /// 获取生成的数据行数。
    /// </summary>
    public int Rows { get; }
    /// <summary>
    /// 获取每批生成的行数。
    /// </summary>
    public int BatchSize { get; }
    /// <summary>
    /// 获取输出工作簿路径。
    /// </summary>
    public string OutputPath { get; }

    /// <summary>
    /// 解析探针命令行选项。
    /// </summary>
    /// <param name="args">命令行参数。</param>
    /// <returns>解析后的选项；参数缺失、无效或行数超限时返回 <see langword="null"/>。</returns>
    public static ProbeOptions Parse(string[] args)
    {
        int? rows = null;
        int? batchSize = null;
        string outputPath = null;
        for (var index = 0; index < args.Length; index++)
        {
            switch (args[index])
            {
                case "--rows" when TryReadPositiveInt(args, ref index, out var rowValue):
                    rows = rowValue;
                    break;
                case "--batch-size" when TryReadPositiveInt(args, ref index, out var batchValue):
                    batchSize = batchValue;
                    break;
                case "--output" when index + 1 < args.Length:
                    outputPath = args[++index];
                    break;
                default:
                    return null;
            }
        }

        if (!rows.HasValue || !batchSize.HasValue || string.IsNullOrWhiteSpace(outputPath)
            || rows.Value > 1_048_575)
            return null;
        return new ProbeOptions(rows.Value, batchSize.Value, outputPath);
    }

    /// <summary>
    /// 尝试读取选项后的正整数。
    /// </summary>
    /// <param name="args">命令行参数。</param>
    /// <param name="index">选项位置；存在后续参数时推进到其位置。</param>
    /// <param name="value">解析得到的整数；未成功解析整数时为零。</param>
    /// <returns>后续参数是正整数时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool TryReadPositiveInt(string[] args, ref int index, out int value)
    {
        value = 0;
        return index + 1 < args.Length
            && int.TryParse(args[++index], NumberStyles.Integer, CultureInfo.InvariantCulture, out value)
            && value > 0;
    }
}
