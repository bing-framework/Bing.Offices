using BenchmarkDotNet.Attributes;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Styles;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;

namespace Bing.Offices.Benchmarks;

/// <summary>
/// Stream-first 导入与单工作簿导出性能基准。
/// </summary>
[MemoryDiagnoser]
public class StreamPipelineBenchmarks
{
    /// <summary>
    /// 测量的数据行数。
    /// </summary>
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    /// <summary>
    /// NPOI 导入器。
    /// </summary>
    private IExcelImporter _importer = null!;

    /// <summary>
    /// NPOI 导出器。
    /// </summary>
    private IExcelExporter _exporter = null!;

    private IServiceProvider _serviceProvider = null!;

    /// <summary>
    /// 基准行集合。
    /// </summary>
    private IReadOnlyList<BenchmarkRow> _rows = Array.Empty<BenchmarkRow>();

    private ExcelWorkbookExportRequest _exportRequest = null!;

    private ExcelWorkbookImportRequest<BenchmarkWorkbook> _importRequest = null!;

    /// <summary>
    /// 预生成的导入工作簿字节。
    /// </summary>
    private byte[] _sourceBytes = Array.Empty<byte>();

    /// <summary>
    /// 初始化基准数据和导入工作簿。
    /// </summary>
    [GlobalSetup]
    public void Setup()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        _serviceProvider = services.BuildServiceProvider();
        _importer = _serviceProvider.GetRequiredService<IExcelImporter>();
        _exporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _rows = Enumerable.Range(0, RowCount)
            .Select(index => new BenchmarkRow
            {
                Code = $"CODE-{index:D6}",
                Quantity = index,
                Amount = index + 0.25m,
                OccurredAt = new DateTime(2026, 8, 13).AddMinutes(index)
            })
            .ToArray();
        _exportRequest = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", _rows));
        _importRequest = ExcelImport.Workbook<BenchmarkWorkbook>(workbook =>
            workbook.Sheet("Sheet1", root => root.Items));

        using var source = new MemoryStream();
        _exporter.Export(_exportRequest, source);
        _sourceBytes = source.ToArray();
    }

    /// <summary>
    /// 释放基准目标流。
    /// </summary>
    [GlobalCleanup]
    public void Cleanup()
    {
        (_serviceProvider as IDisposable)?.Dispose();
    }

    /// <summary>
    /// 测量从 XLSX 输入流导入全部行的成本。
    /// </summary>
    /// <returns>成功导入的行数。</returns>
    [Benchmark]
    public int Import()
    {
        using var source = new MemoryStream(_sourceBytes, writable: false);
        return _importer.Import(source, _importRequest).Workbook.Items.Count;
    }

    /// <summary>
    /// 测量向 XLSX 目标流导出全部行的成本。
    /// </summary>
    /// <returns>生成工作簿的字节数。</returns>
    [Benchmark]
    public long Export()
    {
        using var destination = new MemoryStream();
        _exporter.Export(_exportRequest, destination);
        return destination.Length;
    }

    /// <summary>
    /// 记录新建目标流完成导出后的容量，避免把 retained capacity 隐藏在复用流中。
    /// </summary>
    [Benchmark]
    public long ExportDestinationCapacity()
    {
        using var destination = new MemoryStream();
        _exporter.Export(_exportRequest, destination);
        return destination.Capacity;
    }

    /// <summary>
    /// 基准行模型。
    /// </summary>
    private sealed class BenchmarkRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;

        /// <summary>
        /// 数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 发生时间。
        /// </summary>
        public DateTime OccurredAt { get; set; }
    }

    private sealed class BenchmarkWorkbook
    {
        public List<BenchmarkRow> Items { get; } = new();
    }
}

/// <summary>
/// 随失败行数增长的失败工作簿导出基准。
/// </summary>
[MemoryDiagnoser]
public class FailureWorkbookBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int FailureRowCount { get; set; }

    private IExcelImporter _importer = null!;
    private IServiceProvider _serviceProvider = null!;
    private byte[] _failureBytes = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        _serviceProvider = services.BuildServiceProvider();
        _importer = _serviceProvider.GetRequiredService<IExcelImporter>();
        using var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        for (var row = 1; row <= FailureRowCount; row++)
            sheet.CreateRow(row).CreateCell(0).SetCellValue($"BAD-{row}");
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        _failureBytes = stream.ToArray();
    }

    [GlobalCleanup]
    public void Cleanup() => (_serviceProvider as IDisposable)?.Dispose();

    /// <summary>
    /// 测量随失败行数增长的筛选、重排和受限写出成本。
    /// </summary>
    [Benchmark]
    public long FailureWorkbookExport()
    {
        using var source = new MemoryStream(_failureBytes, writable: false);
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<FailureWorkbook>(workbook =>
            workbook.FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet("Sheet1", root => root.Items));
        var result = _importer.Import(source, request);
        return result.Errors.Count;
    }

    private sealed class FailureWorkbook
    {
        public List<FailureRow> Items { get; } = new();
    }

    private sealed class FailureRow
    {
        [ExcelRegex("^OK-")]
        public string Code { get; set; } = string.Empty;
    }
}

/// <summary>
/// 使用真实 HeaderAttribute 宽表头模型的样式缓存基准。
/// </summary>
[MemoryDiagnoser]
public class HeaderStyleBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int RowCount { get; set; }

    private IExcelExporter _exporter = null!;
    private IServiceProvider _serviceProvider = null!;
    private IReadOnlyList<HeaderStyleRow> _rows = Array.Empty<HeaderStyleRow>();
    private ExcelWorkbookExportRequest _request = null!;

    [GlobalSetup]
    public void Setup()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        _serviceProvider = services.BuildServiceProvider();
        _exporter = _serviceProvider.GetRequiredService<IExcelExporter>();
        _rows = Enumerable.Range(0, RowCount)
            .Select(index => new HeaderStyleRow
            {
                A = $"A-{index}", B = $"B-{index}", C = $"C-{index}", D = $"D-{index}",
                E = $"E-{index}", F = $"F-{index}", G = $"G-{index}", H = $"H-{index}",
                I = $"I-{index}", J = $"J-{index}", K = $"K-{index}", L = $"L-{index}"
            })
            .ToArray();
        _request = ExcelExport.Workbook(workbook => workbook.AddSheet("Sheet1", _rows,
            sheet => sheet.HeaderStyle(new ExcelCellStyle
            {
                FillPattern = ExcelFillPattern.Solid,
                ForegroundColor = new ExcelColor("FF112233")
            })));
    }

    [GlobalCleanup]
    public void Cleanup() => (_serviceProvider as IDisposable)?.Dispose();

    /// <summary>
    /// 测量真实 HeaderAttribute 和请求样式共同作用于宽表头数据的成本。
    /// </summary>
    [Benchmark]
    public long HeaderStyle()
    {
        using var destination = new MemoryStream();
        _exporter.Export(_request, destination);
        return destination.Length;
    }

    [Header(FontName = "Arial", FontSize = 11, Bold = false, Color = Color.Blue)]
    private sealed class HeaderStyleRow
    {
        public string A { get; set; } = string.Empty;
        public string B { get; set; } = string.Empty;
        public string C { get; set; } = string.Empty;
        public string D { get; set; } = string.Empty;
        public string E { get; set; } = string.Empty;
        public string F { get; set; } = string.Empty;
        public string G { get; set; } = string.Empty;
        public string H { get; set; } = string.Empty;
        public string I { get; set; } = string.Empty;
        public string J { get; set; } = string.Empty;
        public string K { get; set; } = string.Empty;
        public string L { get; set; } = string.Empty;
    }
}

/// <summary>
/// 使用随行数扩大的真实 Workbook Data Validation 区间的基准。
/// </summary>
[MemoryDiagnoser]
public class ValidationRangeBenchmarks
{
    [Params(1000, 10000, 100000)]
    public int ValidationRowCount { get; set; }

    private IExcelImporter _importer = null!;
    private IServiceProvider _serviceProvider = null!;
    private byte[] _validationBytes = Array.Empty<byte>();
    private ExcelWorkbookImportRequest<ValidationWorkbook> _request = null!;

    [GlobalSetup]
    public void Setup()
    {
        var services = new ServiceCollection();
        services.AddBingOfficesNpoi();
        _serviceProvider = services.BuildServiceProvider();
        _importer = _serviceProvider.GetRequiredService<IExcelImporter>();
        _request = ExcelImport.Workbook<ValidationWorkbook>(workbook =>
            workbook.ValidationMode(ExcelImportValidationMode.WorkbookRules)
                .Sheet("Sheet1", root => root.Items));
        using var workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
        var sheet = workbook.CreateSheet("Sheet1");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        for (var row = 1; row <= ValidationRowCount; row++)
            sheet.CreateRow(row).CreateCell(0).SetCellValue(5);
        var helper = sheet.GetDataValidationHelper();
        var validationCount = Math.Max(4, ValidationRowCount / 10000 + 4);
        for (var index = 0; index < validationCount; index++)
        {
            var firstRow = index == 0 ? 1 : Math.Max(1, ValidationRowCount / (index + 1));
            var lastRow = index == 0 ? ValidationRowCount
                : Math.Min(ValidationRowCount, ValidationRowCount - ValidationRowCount / (index + 2));
            var constraint = helper.CreateintConstraint(OperatorType.BETWEEN, "1", "10");
            sheet.AddValidationData(helper.CreateValidation(constraint,
                new NPOI.SS.Util.CellRangeAddressList(firstRow, lastRow, 0, 0)));
        }
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        _validationBytes = stream.ToArray();
    }

    [GlobalCleanup]
    public void Cleanup() => (_serviceProvider as IDisposable)?.Dispose();

    /// <summary>
    /// 测量随验证区间和数据行数增长的 Workbook 校验成本。
    /// </summary>
    [Benchmark]
    public int ValidationRange()
    {
        using var source = new MemoryStream(_validationBytes, writable: false);
        return _importer.Import(source, _request).Workbook.Items.Count;
    }

    private sealed class ValidationWorkbook
    {
        public List<ValidationRow> Items { get; } = new();
    }

    private sealed class ValidationRow
    {
        public string Code { get; set; } = string.Empty;
    }
}
