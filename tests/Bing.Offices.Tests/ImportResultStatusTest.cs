using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Bing.Offices.Csv;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>通过真实 Excel/CSV 输入验证完成状态、停止原因及旧截断字段的语义。</summary>
public sealed class ImportResultStatusTest
{
    public static IEnumerable<object[]> ExcelCases()
    {
        foreach (var format in new[] { "xls", "xlsx" })
        foreach (var asynchronous in new[] { false, true })
        foreach (var scenario in new[] { "success", "row-error", "header", "rows", "errors", "sheet-boundary", "exact", "empty-stop" })
            yield return new object[] { format, asynchronous, scenario };
    }

    /// <summary>普通行失败不等于未完成；资源和错误上限应保留部分成功结果并说明原因。</summary>
    [Theory]
    [MemberData(nameof(ExcelCases))]
    public async Task Excel_Import_ShouldReportCompletionWithoutChangingLegacyFlags(string format,
        bool asynchronous, string scenario)
    {
        using var stream = new MemoryStream(Workbook(format, scenario));
        var request = ExcelImport.Workbook<Root>(builder =>
        {
            builder.Sheet("Data", root => root.Rows, sheet =>
            {
                if (scenario == "empty-stop") sheet.StopAtFirstEmptyRow();
            });
            if (scenario == "sheet-boundary")
                builder.Sheet("Other", root => root.OtherRows);
            if (scenario is "rows" or "sheet-boundary" or "exact")
                builder.ResourceLimits(new ExcelResourceLimits { MaxRows = 1 });
            if (scenario == "errors")
                builder.ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 });
        });
        var importer = new NpoiExcelImporter();
        var result = asynchronous ? await importer.ImportAsync(stream, request) : importer.Import(stream, request);
        var reason = scenario switch
        {
            "header" => ImportStopReason.InvalidHeader,
            "rows" or "sheet-boundary" => ImportStopReason.ResourceLimit,
            "errors" => ImportStopReason.ErrorLimit,
            _ => ImportStopReason.None
        };
        Assert.Equal(reason, result.StopReason);
        Assert.Equal(reason == ImportStopReason.None, result.IsComplete);
        Assert.Equal(scenario is "success" or "exact" or "empty-stop", result.IsSuccess);
        Assert.Equal(scenario == "errors", result.ErrorsTruncated);
        Assert.Equal(scenario == "errors" ? 1 : (int?)null, result.MaxErrors);
        var expected = scenario == "header" ? Array.Empty<int>()
            : scenario is "rows" or "errors" or "sheet-boundary" or "exact" or "empty-stop" ? new[] { 7 } : new[] { 7, 9 };
        Assert.Equal(expected, result.Workbook.Rows.Select(row => row.Count));
        Assert.Empty(result.Workbook.OtherRows);
        Assert.Equal(scenario == "row-error" ? new[] { 1, 3 } : Enumerable.Range(1, expected.Length),
            Assert.Single(result.Sheets).SourceRows);
        if (scenario is "success" or "exact" or "empty-stop")
            Assert.Empty(result.Errors);
        else
        {
            var error = Assert.Single(result.Errors);
            Assert.Equal(scenario == "header" ? ExcelImportErrorCode.InvalidHeader
                : scenario is "rows" or "sheet-boundary" ? ExcelImportErrorCode.ResourceLimit
                : ExcelImportErrorCode.ValueConversion, error.Code);
            Assert.Equal(scenario == "sheet-boundary" ? "Other" : "Data", error.SheetName);
            Assert.Equal(scenario == "sheet-boundary" ? 0 : scenario == "header" ? 1 : 3, error.RowIndex);
            Assert.Equal(scenario is "row-error" or "errors" ? 1 : 0, error.ColumnIndex);
            Assert.Equal(scenario is "row-error" or "errors" ? nameof(Row.Count) : null, error.PropertyName);
            if (scenario == "sheet-boundary")
                Assert.Equal("Workbook 数据行数达到限制: 1，仍有未处理的 Sheet。", error.Message);
        }
        Assert.True(stream.CanRead);
    }

    public static IEnumerable<object[]> CsvCases()
    {
        foreach (var asynchronous in new[] { false, true })
        foreach (var scenario in new[] { "success", "row-error", "header", "rows", "errors", "bytes" })
            yield return new object[] { asynchronous, scenario };
    }

    /// <summary>CSV 默认实现应把普通错误、结构拒绝和各种旧截断归纳为与 Excel 一致的摘要。</summary>
    [Theory]
    [MemberData(nameof(CsvCases))]
    public async Task Csv_Import_ShouldReportCompletionAndPreservePartialItems(bool asynchronous, string scenario)
    {
        using var provider = new ServiceCollection().AddBingOfficesCore().BuildServiceProvider();
        var importer = provider.GetRequiredService<ICsvImporter>();
        var options = new CsvImportOptions<Row>
        {
            MaxRows = scenario == "rows" ? 1 : null,
            MaxErrors = scenario == "errors" ? 1 : null,
            MaxInputBytes = scenario == "bytes" ? 1 : null
        };
        var text = scenario == "header" ? "Wrong\r\n7\r\n"
            : scenario is "row-error" or "errors" ? "Count\r\n7\r\nbad\r\n9\r\n" : "Count\r\n7\r\n9\r\n";
        var bytes = Encoding.UTF8.GetBytes(text);
        var result = asynchronous ? await importer.ImportFromBytesAsync<Row>(bytes, options) : importer.ImportFromBytes<Row>(bytes, options);
        var reason = scenario switch
        {
            "header" => ImportStopReason.InvalidHeader,
            "rows" or "bytes" => ImportStopReason.ResourceLimit,
            "errors" => ImportStopReason.ErrorLimit,
            _ => ImportStopReason.None
        };
        Assert.Equal(reason, result.StopReason);
        Assert.Equal(reason == ImportStopReason.None, result.IsComplete);
        Assert.Equal(scenario == "success", result.IsSuccess);
        Assert.Equal(scenario is "rows" or "errors" or "bytes", result.IsTruncated);
        Assert.Equal(options.MaxErrors, result.MaxErrors);
        Assert.Equal(scenario is "header" or "bytes" ? Array.Empty<int>()
            : scenario is "rows" or "errors" ? new[] { 7 } : new[] { 7, 9 }, result.Items.Select(row => row.Count));
        if (scenario == "success")
            Assert.Empty(result.Errors);
        else
        {
            var error = Assert.Single(result.Errors);
            Assert.Equal(scenario == "header" ? CsvImportErrorCode.InvalidHeader
                : scenario is "rows" or "bytes" ? CsvImportErrorCode.ResourceLimit : CsvImportErrorCode.ValueConversion, error.Code);
            Assert.Equal(scenario == "header" ? 1 : scenario == "bytes" ? 0 : 3, error.RowIndex);
            Assert.Equal(scenario is "row-error" or "errors" ? 1 : 0, error.ColumnIndex);
            Assert.Equal(scenario is "row-error" or "errors" ? nameof(Row.Count) : null, error.PropertyName);
            Assert.Null(error.FirstRowNumber);
            Assert.False(string.IsNullOrWhiteSpace(error.Message));
        }
    }

    /// <summary>旧构造器无原因截断和多原因结果必须有稳定优先级，不把不完整结果伪装成成功。</summary>
    [Fact]
    public void ResultConstructors_ShouldSummarizeMultipleReasonsAndUnspecifiedTruncation()
    {
        var csv = new CsvImportResult<Row>(Array.Empty<Row>(), Array.Empty<CsvImportError>(), true);
        Assert.Equal(ImportStopReason.Unspecified, csv.StopReason);
        Assert.False(csv.IsSuccess);
        Assert.False(csv.IsComplete);
        var csvErrors = new[]
        {
            new CsvImportError("header", 1, 0, null, code: CsvImportErrorCode.InvalidHeader),
            new CsvImportError("limit", 0, 0, null, code: CsvImportErrorCode.ResourceLimit)
        };
        Assert.Equal(ImportStopReason.ResourceLimit,
            new CsvImportResult<Row>(Array.Empty<Row>(), csvErrors, true, 1).StopReason);
        Assert.Equal(ImportStopReason.ErrorLimit,
            new CsvImportResult<Row>(Array.Empty<Row>(), csvErrors.Take(1).ToArray(), true, 1).StopReason);
        var excelErrors = new[]
        {
            new ExcelImportError(ExcelImportErrorCode.InvalidHeader, "header", "Data", 1, 0, null),
            new ExcelImportError(ExcelImportErrorCode.ResourceLimit, "limit", "Data", 0, 0, null)
        };
        var workbook = new ExcelWorkbookImportResult<Root>(new Root(), Array.Empty<ExcelSheetImportResult>(), excelErrors, true, 1);
        var sheet = new ExcelSingleSheetImportResult<Row>(Array.Empty<Row>(), excelErrors, Array.Empty<int>(), true, 1);
        Assert.Equal(ImportStopReason.ResourceLimit, workbook.StopReason);
        Assert.Equal(workbook.StopReason, sheet.StopReason);
        Assert.False(sheet.IsSuccess);
        Assert.False(sheet.IsComplete);
        Assert.Same(excelErrors, sheet.Errors);
        Assert.Equal(ImportStopReason.ErrorLimit, new ExcelWorkbookImportResult<Root>(new Root(),
            Array.Empty<ExcelSheetImportResult>(), excelErrors.Take(1).ToArray(), true, 1).StopReason);
    }

    private static byte[] Workbook(string format, string scenario)
    {
        using IWorkbook workbook = format == "xls" ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue(scenario == "header" ? "Wrong" : "Count");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("7");
        if (scenario is "row-error" or "errors")
        {
            sheet.CreateRow(2).CreateCell(0).SetCellValue("bad");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("9");
        }
        else if (scenario == "empty-stop")
            sheet.CreateRow(3).CreateCell(0).SetCellValue("9");
        else if (scenario is not "exact" and not "sheet-boundary")
            sheet.CreateRow(2).CreateCell(0).SetCellValue("9");
        var other = workbook.CreateSheet("Other");
        other.CreateRow(0).CreateCell(0).SetCellValue("Count");
        other.CreateRow(1).CreateCell(0).SetCellValue("42");
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private sealed class Row { public int Count { get; set; } }
    private sealed class Root
    {
        public List<Row> Rows { get; } = new();
        public List<Row> OtherRows { get; } = new();
    }
}
