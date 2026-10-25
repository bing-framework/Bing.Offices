using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 Excel 导入默认校验范围和表头列数上限。
/// </summary>
public sealed class UsabilityDefaultsTest
{
    /// <summary>
    /// 测试 - 默认 StopOnFirstFailure 只跳过失败行，并继续导入后续行。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls, false)]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    public async Task Import_DefaultValidation_ShouldContinueAfterInvalidRow(
        ExcelFormat format, bool useAsync)
    {
        using var source = new MemoryStream(CreateWorkbook(format, new[] { "Code" },
            new[] { "OK-1" }, new[] { "BAD" }, new[] { "OK-2" }));
        var request = ExcelImport.Workbook<RowsWorkbook<ValidatedRow>>(builder =>
            builder.Sheet("Data", workbook => workbook.Rows));
        var importer = new NpoiExcelImporter();

        var result = useAsync
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        Assert.Equal(new[] { "OK-1", "OK-2" }, result.Workbook.Rows.Select(row => row.Code));

        var sheetResult = Assert.Single(result.Sheets);
        Assert.Equal("Data", sheetResult.Name);
        Assert.Equal(typeof(ValidatedRow), sheetResult.ItemType);
        Assert.Equal(new[] { 1, 3 }, sheetResult.SourceRows);

        var error = Assert.Single(result.Errors);
        Assert.Same(error, Assert.Single(sheetResult.Errors));
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("格式不正确", error.Message);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(ValidatedRow.Code), error.PropertyName);
        Assert.Equal(nameof(ValidatedRow.Code), error.ColumnKey);
        Assert.Equal("Code", error.Header);
        Assert.Equal("BAD", error.RawValue);
        Assert.Null(error.FirstRowNumber);
    }

    /// <summary>
    /// 测试 - 默认允许读取 100 列，拒绝 101 列并可通过显式上限读取宽表。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls, false)]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    public async Task Import_DefaultColumnLimit_ShouldAllow100Reject101AndHonorExpansion(
        ExcelFormat format, bool useAsync)
    {
        var importer = new NpoiExcelImporter();
        var hundredHeaders = CreateHeaders(100);
        using var hundredColumnSource = new MemoryStream(CreateWorkbook(format, hundredHeaders,
            CreateRowValues(100, "OK-100")));
        var hundredColumnResult = await ImportColumnLimitAsync(
            importer, hundredColumnSource, CreateColumnLimitRequest(), useAsync);
        AssertSuccessfulColumnLimitResult(hundredColumnResult, "OK-100");

        var hundredOneHeaders = CreateHeaders(101);
        using var rejectedSource = new MemoryStream(CreateWorkbook(format, hundredOneHeaders,
            CreateRowValues(101, "OK-101")));
        var rejectedResult = await ImportColumnLimitAsync(
            importer, rejectedSource, CreateColumnLimitRequest(), useAsync);
        AssertInvalidHeaderResult(rejectedResult, 100);

        using var expandedSource = new MemoryStream(CreateWorkbook(format, hundredOneHeaders,
            CreateRowValues(101, "OK-101")));
        var expandedResult = await ImportColumnLimitAsync(
            importer, expandedSource, CreateColumnLimitRequest(101), useAsync);
        AssertSuccessfulColumnLimitResult(expandedResult, "OK-101");
    }

    private static async Task<ExcelWorkbookImportResult<RowsWorkbook<ColumnLimitRow>>> ImportColumnLimitAsync(
        NpoiExcelImporter importer, Stream source,
        ExcelWorkbookImportRequest<RowsWorkbook<ColumnLimitRow>> request, bool useAsync) =>
        useAsync
            ? await importer.ImportAsync(source, request)
            : importer.Import(source, request);

    private static ExcelWorkbookImportRequest<RowsWorkbook<ColumnLimitRow>> CreateColumnLimitRequest(
        int? maxReadColumns = null) =>
        ExcelImport.Workbook<RowsWorkbook<ColumnLimitRow>>(builder => builder.Sheet("Data",
            workbook => workbook.Rows, sheet =>
            {
                if (maxReadColumns.HasValue)
                    sheet.MaxReadColumns(maxReadColumns.Value);
            }));

    private static void AssertSuccessfulColumnLimitResult(
        ExcelWorkbookImportResult<RowsWorkbook<ColumnLimitRow>> result, string expectedCode)
    {
        Assert.True(result.IsSuccess, string.Join("; ", result.Errors.Select(error => error.Message)));
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        Assert.Equal(expectedCode, Assert.Single(result.Workbook.Rows).Code);
        Assert.Empty(result.Errors);

        var sheetResult = Assert.Single(result.Sheets);
        Assert.Equal("Data", sheetResult.Name);
        Assert.Equal(typeof(ColumnLimitRow), sheetResult.ItemType);
        Assert.Equal(new[] { 1 }, sheetResult.SourceRows);
        Assert.Empty(sheetResult.Errors);
    }

    private static void AssertInvalidHeaderResult(
        ExcelWorkbookImportResult<RowsWorkbook<ColumnLimitRow>> result, int maxColumns)
    {
        Assert.False(result.IsSuccess);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        Assert.Empty(result.Workbook.Rows);

        var sheetResult = Assert.Single(result.Sheets);
        Assert.Equal("Data", sheetResult.Name);
        Assert.Equal(typeof(ColumnLimitRow), sheetResult.ItemType);
        Assert.Empty(sheetResult.SourceRows);

        var error = Assert.Single(result.Errors);
        Assert.Same(error, Assert.Single(sheetResult.Errors));
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, error.Code);
        Assert.Equal($"导入表头超过最大列长度: {maxColumns}", error.Message);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        Assert.Null(error.ColumnKey);
        Assert.Null(error.Header);
        Assert.Null(error.RawValue);
        Assert.Null(error.FirstRowNumber);
    }

    private static string[] CreateHeaders(int count) =>
        Enumerable.Range(0, count)
            .Select(index => index == 0 ? "Code" : $"Extra{index}")
            .ToArray();

    private static string[] CreateRowValues(int columnCount, string code)
    {
        var values = new string[columnCount];
        values[0] = code;
        return values;
    }

    private static byte[] CreateWorkbook(ExcelFormat format, IReadOnlyList<string> headers,
        params string[][] dataRows)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        var header = sheet.CreateRow(0);
        for (var column = 0; column < headers.Count; column++)
            header.CreateCell(column).SetCellValue(headers[column]);

        for (var rowIndex = 0; rowIndex < dataRows.Length; rowIndex++)
        {
            var row = sheet.CreateRow(rowIndex + 1);
            for (var column = 0; column < dataRows[rowIndex].Length; column++)
                row.CreateCell(column).SetCellValue(dataRows[rowIndex][column]);
        }

        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private sealed class RowsWorkbook<T> where T : class, new()
    {
        public List<T> Rows { get; } = new();
    }

    private sealed class ValidatedRow
    {
        [ExcelRegex("^OK-")]
        public string Code { get; set; } = string.Empty;
    }

    private sealed class ColumnLimitRow
    {
        public string Code { get; set; } = string.Empty;
    }
}
