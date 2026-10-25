using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>验证 Workbook 行数限额在 Sheet 边界产生的诊断及失败工作簿内容。</summary>
public sealed class SheetBoundaryDiagnosticTest
{
    private const string ResourceLimitMessage = "Workbook 数据行数达到限制: 1，仍有未处理的 Sheet。";

    public static IEnumerable<object[]> SheetBoundaryCases()
    {
        foreach (var format in new[] { "xls", "xlsx" })
        foreach (var asynchronous in new[] { false, true })
        foreach (var selectorCase in new[] { "index", "different-case-name", "missing-name" })
            yield return new object[] { format, asynchronous, selectorCase };
    }

    /// <summary>限额恰在第一张表结束时耗尽，应报告实际解析表名或缺失选择器回退值。</summary>
    [Theory]
    [MemberData(nameof(SheetBoundaryCases))]
    public async Task Import_RowLimitAtSheetBoundary_ShouldReportResolvedNameAndCompleteFailureSummary(
        string format, bool asynchronous, string selectorCase)
    {
        using var source = new MemoryStream(CreateWorkbook(format), writable: false);
        using var failure = new MemoryStream();
        var secondSelector = selectorCase switch
        {
            "index" => ExcelSheetSelector.ByIndex(1),
            "different-case-name" => ExcelSheetSelector.ByName("oRdErS"),
            "missing-name" => ExcelSheetSelector.ByName("Missing"),
            _ => throw new Xunit.Sdk.XunitException($"Unknown selector case: {selectorCase}")
        };
        var expectedSheetName = selectorCase == "missing-name" ? "Missing" : "Orders";
        var request = ExcelImport.Workbook<BoundaryWorkbook>(builder => builder
            .SheetNameComparison(ExcelNameComparison.OrdinalIgnoreCase)
            .Sheet(ExcelSheetSelector.ByIndex(0), workbook => workbook.First)
            .Sheet(secondSelector, workbook => workbook.Second)
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure
            }));

        var result = asynchronous
            ? await new NpoiExcelImporter().ImportAsync(source, request)
            : new NpoiExcelImporter().Import(source, request);

        Assert.True(source.CanRead);
        Assert.True(failure.CanWrite);
        Assert.False(result.IsSuccess);
        Assert.False(result.IsComplete);
        Assert.Equal(ImportStopReason.ResourceLimit, result.StopReason);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);

        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        Assert.Equal(ResourceLimitMessage, error.Message);
        Assert.Equal(expectedSheetName, error.SheetName);
        Assert.Equal(0, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        Assert.Null(error.ColumnKey);
        Assert.Null(error.Header);
        Assert.Null(error.RawValue);
        Assert.Null(error.FirstRowNumber);

        Assert.Equal(new[] { 7 }, result.Workbook.First.Select(item => item.Count));
        Assert.Empty(result.Workbook.Second);
        var importedSheet = Assert.Single(result.Sheets);
        Assert.Equal("Data", importedSheet.Name);
        Assert.Equal(new[] { 1 }, importedSheet.SourceRows);
        Assert.Empty(importedSheet.Errors);

        failure.Position = 0;
        using var failedWorkbook = WorkbookFactory.Create(failure);
        var summary = Assert.IsAssignableFrom<ISheet>(failedWorkbook.GetSheet("_ImportErrors"));
        Assert.Equal(1, summary.LastRowNum);
        Assert.Equal(2, summary.PhysicalNumberOfRows);
        Assert.Equal(new[]
        {
            "Code", "Message", "Sheet", "Row", "Column", "Property", "Header", "RawValue"
        }, summary.GetRow(0).Cells.Select(cell => cell.StringCellValue));
        var summaryError = summary.GetRow(1);
        Assert.Equal(8, summaryError.PhysicalNumberOfCells);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit.ToString(), summaryError.GetCell(0).StringCellValue);
        Assert.Equal(ResourceLimitMessage, summaryError.GetCell(1).StringCellValue);
        Assert.Equal(expectedSheetName, summaryError.GetCell(2).StringCellValue);
        Assert.Equal(0d, summaryError.GetCell(3).NumericCellValue);
        Assert.Equal(0d, summaryError.GetCell(4).NumericCellValue);
        Assert.Equal(string.Empty, summaryError.GetCell(5).StringCellValue);
        Assert.Equal(string.Empty, summaryError.GetCell(6).StringCellValue);
        Assert.Equal(string.Empty, summaryError.GetCell(7).StringCellValue);
    }

    private static byte[] CreateWorkbook(string format)
    {
        using IWorkbook workbook = format == "xls" ? new HSSFWorkbook() : new XSSFWorkbook();
        AddSheet(workbook, "Data", 7);
        AddSheet(workbook, "Orders", 42);
        using var stream = new MemoryStream();
        workbook.Write(stream, true);
        return stream.ToArray();
    }

    private static void AddSheet(IWorkbook workbook, string name, int value)
    {
        var sheet = workbook.CreateSheet(name);
        sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(BoundaryRow.Count));
        sheet.CreateRow(1).CreateCell(0).SetCellValue(value);
    }

    private sealed class BoundaryRow
    {
        public int Count { get; set; }
    }

    private sealed class BoundaryWorkbook
    {
        public List<BoundaryRow> First { get; } = new();
        public List<BoundaryRow> Second { get; } = new();
    }
}
