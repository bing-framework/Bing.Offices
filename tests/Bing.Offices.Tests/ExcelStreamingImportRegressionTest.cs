using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using static Bing.Offices.Tests.NpoiExcelStreamingImporterTest;

namespace Bing.Offices.Tests;

/// <summary>低内存导入与既有标量、选列、错误位置和消费者决定契约的回归测试。</summary>
public class ExcelStreamingImportRegressionTest
{
    public static IEnumerable<object[]> DateCases()
    {
        foreach (var asynchronous in new[] { false, true })
        foreach (var date1904 in new[] { false, true })
        foreach (var customFormat in new[] { false, true })
        foreach (var formula in new[] { false, true })
            yield return new object[] { asynchronous, date1904, customFormat, formula };
    }

    [Theory]
    [MemberData(nameof(DateCases))]
    public async Task Import_DateTextAndContexts_ShouldMatchDom(bool asynchronous, bool date1904, bool customFormat, bool formula)
    {
        using var book = new XSSFWorkbook();
        book.GetCTWorkbook().workbookPr.date1904 = date1904;
        var sheet = book.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
        var cell = sheet.CreateRow(1).CreateCell(0);
        if (formula) cell.SetCellFormula("1+0.5");
        cell.SetCellValue(1.5);
        var style = book.CreateCellStyle();
        style.DataFormat = customFormat ? book.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss") : (short)14;
        cell.CellStyle = style;
        using var source = new MemoryStream();
        book.Write(source, true);
        source.Position = 0;
        var domRule = new CaptureRule();
        var domConverter = new CaptureConverter();
        var dom = new NpoiExcelImporter(validationRules: new[] { domRule }, valueConverters: new[] { domConverter })
            .ImportSheet<TextValue>(source);
        Assert.True(dom.IsSuccess);
        var expectedDate = date1904 ? new DateTime(1904, 1, 2, 12, 0, 0) : new DateTime(1900, 1, 1, 12, 0, 0);
        var expectedText = expectedDate.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
        Assert.Equal(expectedText, Assert.Single(dom.Items).Value);
        source.Position = 0;
        var rule = new CaptureRule();
        var converter = new CaptureConverter();
        var importer = new NpoiExcelStreamingImporter(new[] { rule }, new[] { converter });
        ExcelImportRow<TextValue> actual = null;
        var summary = await Read(importer, ExcelStreamingImport.Sheet<TextValue>(0), source, asynchronous,
            row => { actual = row; return ExcelRowReadDecision.Continue; });
        Assert.True(summary.IsSuccess);
        Assert.Equal(expectedText, actual.Item.Value);
        Assert.Equal(expectedText, Assert.Single(rule.Contexts).Value);
        Assert.Equal(expectedText, Assert.Single(converter.Contexts).Cell.Text);
        foreach (var value in new[] { rule.Contexts[0].Cell, converter.Contexts[0].Cell })
        {
            Assert.Equal(expectedDate, Assert.IsType<DateTime>(value.Value));
            Assert.Equal(date1904, value.IsDate1904);
            Assert.Equal(formula ? ExcelCellKind.Formula : ExcelCellKind.DateTime, value.Kind);
            Assert.Equal(formula ? ExcelCellKind.DateTime : (ExcelCellKind?)null, value.CachedKind);
            Assert.Equal(formula ? "1+0.5" : null, value.Formula);
            Assert.Equal(style.DataFormat, value.FormatIndex);
        }
        Assert.Equal(domRule.Contexts[0].Value, rule.Contexts[0].Value);
        Assert.Equal(domConverter.Contexts[0].Value, converter.Contexts[0].Value);
        source.Position = 0;
        DateTime date = default;
        summary = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<DateValue>(0), source,
            asynchronous, row => { date = row.Item.Value; return ExcelRowReadDecision.Continue; });
        Assert.True(summary.IsSuccess); Assert.Equal(expectedDate, date);
    }

    [Theory]
    [InlineData(false, false, false)] [InlineData(true, false, false)]
    [InlineData(false, true, false)] [InlineData(true, true, false)]
    [InlineData(false, false, true)] [InlineData(true, false, true)]
    [InlineData(false, true, true)] [InlineData(true, true, true)]
    public async Task Import_UnboundCellFailures_ShouldNotRejectSelectedValues(bool asynchronous, bool explicitRange, bool missingCache)
    {
        var bad = missingCache ? "<c r=\"B2\"><f>1+1</f></c>" : "<c r=\"B2\" t=\"e\"><v>#DIV/0!</v></c>";
        using var source = new MemoryStream(Package(Header("Count", "Unused") + "<row r=\"2\"><c r=\"A2\"><v>7</v></c>" + bad + "</row>"));
        var rows = new List<ExcelImportRow<NumberRow>>();
        var result = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0,
            s => { if (explicitRange) s.ReadColumns(0, 1); }), source, asynchronous,
            row => { rows.Add(row); return ExcelRowReadDecision.Continue; });
        Assert.True(result.IsSuccess); Assert.Empty(result.Errors);
        var actual = Assert.Single(rows); Assert.Equal(1, actual.RowIndex);
        Assert.Equal(7, actual.Item.Count); Assert.Empty(actual.Errors);
        Assert.Equal(1, result.RowsRead); Assert.Equal(1, result.SuccessfulRows);
        Assert.Equal(0, result.FailedRows); Assert.Equal(1, result.DeliveredRows);
    }

    [Theory]
    [InlineData(false, false)] [InlineData(true, false)]
    [InlineData(false, true)] [InlineData(true, true)]
    public async Task Import_BoundCellFailures_ShouldRetainStructuredDiagnostics(bool asynchronous, bool missingCache)
    {
        var bad = missingCache ? "<c r=\"A2\"><f>1+1</f></c>" : "<c r=\"A2\" t=\"e\"><v>#DIV/0!</v></c>";
        using var source = new MemoryStream(Package(Header("Count") + "<row r=\"2\">" + bad + "</row>"));
        ExcelImportRow<NumberRow> row = null;
        var result = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0), source,
            asynchronous, value => { row = value; return ExcelRowReadDecision.Continue; });
        Assert.True(result.IsComplete); Assert.False(result.IsSuccess); Assert.Null(row.Item);
        var error = Assert.Single(row.Errors); Assert.Same(error, Assert.Single(result.Errors));
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("数据", error.SheetName); Assert.Equal(2, error.RowIndex); Assert.Equal(1, error.ColumnIndex);
        Assert.Equal("Count", error.PropertyName); Assert.Equal("Count", error.ColumnKey); Assert.Equal("Count", error.Header);
        Assert.Equal(missingCache ? "" : "7", error.RawValue);
        Assert.Equal(missingCache ? "公式缺少缓存结果，流式导入不执行公式计算。" : "单元格包含公式错误码: 7", error.Message);
        Assert.Null(error.FirstRowNumber); Assert.Equal(1, result.TotalErrors);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Import_UnknownExcelError_ShouldDeliverBoundDiagnostic(bool asynchronous)
    {
        using var source = new MemoryStream(Package(Header("Count") + "<row r=\"2\"><c r=\"A2\" t=\"e\"><v>#UNKNOWN!</v></c></row>"));
        var result = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0), source,
            asynchronous, row => { Assert.Null(row.Item); return ExcelRowReadDecision.Continue; });
        Assert.True(result.IsComplete); Assert.Equal(1, result.FailedRows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code); Assert.Equal("Count", error.PropertyName);
        Assert.Equal("#UNKNOWN!", error.RawValue); Assert.Equal(2, error.RowIndex); Assert.Equal(1, error.ColumnIndex);
    }

    [Theory]
    [InlineData(false, 1, 3)] [InlineData(true, 1, 3)]
    [InlineData(false, 4, 9)] [InlineData(true, 4, 9)]
    public async Task Import_RowLimit_ShouldLocateActualPhysicalRow(bool asynchronous, int start, int overflowRow)
    {
        using var source = new MemoryStream(Package(Header("Count") +
            $"<row r=\"{start + 1}\"><c r=\"A{start + 1}\"><v>7</v></c></row><row r=\"{overflowRow}\"><c r=\"A{overflowRow}\"><v>8</v></c></row>"));
        var calls = 0;
        var result = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0,
            s => s.DataRowStartIndex(start)), source, asynchronous, row => { calls++; Assert.Equal(7, row.Item.Count); return ExcelRowReadDecision.Continue; },
            new() { ResourceLimits = new ExcelResourceLimits { MaxRows = 1 } });
        Assert.Equal(1, calls); Assert.Equal(1, result.DeliveredRows); Assert.Equal(1, result.RowsRead);
        Assert.Equal(ImportStopReason.ResourceLimit, result.StopReason); Assert.False(result.IsComplete);
        var error = Assert.Single(result.Errors); Assert.Equal(overflowRow, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex); Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
    }

    [Theory]
    [InlineData(false, false)] [InlineData(true, false)]
    [InlineData(false, true)] [InlineData(true, true)]
    public async Task Import_IgnoredColumnStillObeysTextQuota_WithSparseRowLocation(bool asynchronous, bool rowQuota)
    {
        using var source = new MemoryStream(Package(Header("Count", "Unused") +
            "<row r=\"20\"><c r=\"A20\"><v>7</v></c><c r=\"B20\" t=\"inlineStr\"><is><t>abcdefghijklmnop</t></is></c></row>"));
        var result = await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0, s => s.ReadColumns(0, 1)),
            source, asynchronous, _ => throw new InvalidOperationException("不得交付超额行"),
            rowQuota ? new() { MaxRowCharacters = 12 } : new() { MaxCellCharacters = 12 });
        Assert.False(result.IsComplete); Assert.Equal(ImportStopReason.ResourceLimit, result.StopReason);
        var error = Assert.Single(result.Errors); Assert.Equal(20, error.RowIndex);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code); Assert.Equal(0, result.DeliveredRows);
    }

    [Theory]
    [InlineData(false, false, false)] [InlineData(true, false, false)]
    [InlineData(false, true, false)] [InlineData(true, true, false)]
    [InlineData(false, false, true)] [InlineData(true, false, true)]
    [InlineData(false, true, true)] [InlineData(true, true, true)]
    public async Task Import_ConsumerDecision_ShouldTakePrecedenceOverThreshold(bool asynchronous, bool stop, bool uniqueLimit)
    {
        var xml = uniqueLimit
            ? Header("Code") + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>AB</t></is></c></row>"
            : Header("Count") + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>bad</t></is></c></row>";
        // 后续畸形行若被预读会抛异常，确保停止决定之后没有继续读取。
        using var source = new MemoryStream(Package(xml + "<row r=\"1\"/>"));
        var calls = 0;
        var decision = stop ? ExcelRowReadDecision.Stop : ExcelRowReadDecision.Continue;
        var options = new ExcelStreamingImportOptions { MaxErrors = 1, MaxTrackedUniqueCharacters = 1 };
        ExcelRowReadDecision Consume<T>(ExcelImportRow<T> row) where T : class, new()
        { calls++; Assert.False(row.IsValid); Assert.Null(row.Item); Assert.Single(row.Errors); return decision; }
        var result = uniqueLimit
            ? await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<UniqueRow>(0), source, asynchronous, Consume, options)
            : await Read(new NpoiExcelStreamingImporter(), ExcelStreamingImport.Sheet<NumberRow>(0), source, asynchronous, Consume, options);
        Assert.Equal(stop ? ImportStopReason.ConsumerStopped : uniqueLimit ? ImportStopReason.ResourceLimit : ImportStopReason.ErrorLimit, result.StopReason);
        Assert.Equal(1, calls); Assert.Equal(1, result.RowsRead); Assert.Equal(1, result.DeliveredRows);
        Assert.Equal(1, result.FailedRows); Assert.Equal(0, result.SuccessfulRows); Assert.Equal(1, result.TotalErrors);
        Assert.False(result.IsComplete); Assert.False(result.IsSuccess);
    }

    private static async Task<ExcelStreamingImportSummary> Read<T>(IExcelStreamingImporter importer,
        ExcelStreamingImportRequest<T> request, Stream input, bool asynchronous,
        Func<ExcelImportRow<T>, ExcelRowReadDecision> consumer, ExcelStreamingImportOptions options = null) where T : class, new()
        => asynchronous ? await importer.ImportAsync(request, input, (row, _) => Task.FromResult(consumer(row)), options)
            : importer.Import(request, input, consumer, options);

    public class TextValue { [Capture] public string Value { get; set; } }
    public class DateValue { public DateTime Value { get; set; } }
    public sealed class CaptureAttribute : FilterAttributeBase { }
    private sealed class CaptureRule : IExcelValidationRule
    {
        internal List<ExcelValidationContext> Contexts { get; } = new();
        public bool CanValidate(FilterAttributeBase attribute) => attribute is CaptureAttribute;
        public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) { Contexts.Add(context); return true; }
    }
    private sealed class CaptureConverter : IExcelValueConverter
    {
        internal List<ExcelConversionContext> Contexts { get; } = new();
        public bool CanConvert(Type type) => type == typeof(string);
        public bool TryConvertFrom(ExcelConversionContext context, out object value) { Contexts.Add(context); value = null; return false; }
        public bool TryConvertTo(ExcelConversionContext context, out object value) { value = null; return false; }
    }
}
