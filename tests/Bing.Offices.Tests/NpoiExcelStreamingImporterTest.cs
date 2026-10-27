using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Imports;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Streaming;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>使用真实 ZIP/XML 验证逐行导入契约，不依赖 DOM 导出构造测试输入。</summary>
public class NpoiExcelStreamingImporterTest
{
    [Fact]
    public void Import_ShouldReuseMappingAndPredeclaredDynamicConversion()
    {
        using var input = new MemoryStream(Package(Header("编码", "扩展") +
            "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>甲</t></is></c><c r=\"B2\"><v>42</v></c></row>"));
        var request = ExcelStreamingImport.Sheet<DynamicRow>(0, sheet => sheet.Mapping(new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration> { new() { PropertyName = nameof(DynamicRow.Code), Title = "编码" } }
        }).DynamicColumns(row => row.Values, new[] { new ExcelDynamicColumnDefinition { Key = "extra", Title = "扩展", DataType = typeof(int) } }));
        var result = new NpoiExcelStreamingImporter().Import(request, input, row =>
        {
            Assert.True(row.IsValid); Assert.Equal("甲", row.Item.Code);
            var entry = Assert.Single(row.Item.Values); Assert.Equal("extra", entry.Key); Assert.Equal(42, entry.Value);
            return ExcelRowReadDecision.Continue;
        });
        Assert.True(result.IsSuccess); Assert.Equal(1, result.DeliveredRows);
        Assert.Null(request.Sheet.Target);
    }

    [Fact]
    public void Import_ShouldReuseValidationAndStopAtUniqueCharacterBudget()
    {
        var rows = Header("Code") + string.Concat(new[] { "A", "A", "B", "C" }.Select((value, index) =>
            $"<row r=\"{index + 2}\"><c r=\"A{index + 2}\" t=\"inlineStr\"><is><t>{value}</t></is></c></row>"));
        using var input = new MemoryStream(Package(rows));
        var results = new List<ExcelImportRow<UniqueRow>>();
        var summary = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<UniqueRow>(0), input,
            row => { results.Add(row); return ExcelRowReadDecision.Continue; }, new() { MaxTrackedUniqueCharacters = 2 });
        Assert.Equal(new[] { true, false, true, false }, results.Select(row => row.IsValid));
        Assert.Equal(2, Assert.Single(results[1].Errors).FirstRowNumber);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, Assert.Single(results[3].Errors).Code);
        Assert.Equal(ImportStopReason.ResourceLimit, summary.StopReason);
        Assert.Equal(2, summary.SuccessfulRows); Assert.Equal(2, summary.FailedRows); Assert.Equal(2, summary.TotalErrors);
    }

    [Theory]
    [InlineData("input")]
    [InlineData("temporary")]
    [InlineData("styles")]
    [InlineData("cell")]
    [InlineData("row")]
    public void Import_ShouldEnforceIndependentResourceBudgets(string budget)
    {
        using var source = new MemoryStream(Package(Header("Count") + NumericRows("123456")));
        var options = budget switch
        {
            "input" => new ExcelStreamingImportOptions { ResourceLimits = new ExcelResourceLimits { MaxInputBytes = 1 } },
            "temporary" => new ExcelStreamingImportOptions { MaxTemporaryBytes = 1 },
            "styles" => new ExcelStreamingImportOptions { MaxStyles = 1 },
            "cell" => new ExcelStreamingImportOptions { MaxCellCharacters = 5 },
            _ => new ExcelStreamingImportOptions { MaxRowCharacters = 5 }
        };
        var importer = new NpoiExcelStreamingImporter();
        var calls = 0;
        ExcelStreamingImportSummary Invoke() => importer.Import(ExcelStreamingImport.Sheet<NumberRow>(0), source,
            _ => { calls++; return ExcelRowReadDecision.Continue; }, options);
        if (budget is "cell" or "row")
        {
            var result = Invoke(); Assert.Equal(ImportStopReason.ResourceLimit, result.StopReason);
            var error = Assert.Single(result.Errors); Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        }
        else Assert.Throws<StreamingResourceLimitException>(() => Invoke());
        Assert.Equal(0, calls); Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_ShouldReadAllRows_FromCurrentPosition_AndKeepSourceOpen(bool asynchronous)
    {
        var bytes = Package(Header("Count", "Name") +
            "<row r=\"2\"><c r=\"A2\"><v>7</v></c><c r=\"B2\" t=\"inlineStr\"><is><t>中文_x005F_x0041_</t></is></c></row>" +
            "<row r=\"4\"><c r=\"A4\"><f>1+8</f><v>9</v></c><c r=\"B4\" t=\"s\"><v>0</v></c></row>",
            "<si><r><t>富</t></r><rPh sb=\"0\" eb=\"1\"><t>fu</t></rPh><r><t>文本</t></r></si>");
        using var source = new MemoryStream(new byte[] { 1, 2, 3 }.Concat(bytes).ToArray());
        source.Position = 3;
        var rows = new List<ExcelImportRow<Row>>();
        IExcelStreamingImporter importer = new NpoiExcelStreamingImporter();
        var request = ExcelStreamingImport.Sheet<Row>("数据");
        var summary = asynchronous
            ? await importer.ImportAsync(request, source, (row, _) => { rows.Add(row); return Task.FromResult(ExcelRowReadDecision.Continue); })
            : importer.Import(request, source, row => { rows.Add(row); return ExcelRowReadDecision.Continue; });
        Assert.Equal(new[] { 1, 3 }, rows.Select(row => row.RowIndex));
        Assert.Equal(new[] { 7, 9 }, rows.Select(row => row.Item.Count));
        Assert.Equal(new[] { "中文_x0041_", "富文本" }, rows.Select(row => row.Item.Name));
        Assert.All(rows, row => { Assert.True(row.IsValid); Assert.Empty(row.Errors); });
        Assert.True(summary.IsComplete); Assert.True(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead); Assert.Equal(2, summary.SuccessfulRows);
        Assert.Equal(0, summary.FailedRows); Assert.Equal(2, summary.DeliveredRows);
        Assert.Equal(0, summary.TotalErrors); Assert.Empty(summary.Errors);
        Assert.False(summary.ErrorsTruncated); Assert.Equal(ImportStopReason.None, summary.StopReason);
        Assert.True(source.CanRead);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Import_ShouldConvertBothDateSystems(bool date1904)
    {
        using var source = new MemoryStream(Package(Header("When") +
            "<row r=\"2\"><c r=\"A2\" s=\"1\"><v>1.5</v></c></row>", date1904: date1904));
        DateTime actual = default;
        var summary = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<DateRow>(0), source,
            row => { Assert.True(row.IsValid); actual = row.Item.When; return ExcelRowReadDecision.Continue; });
        Assert.Equal(date1904 ? new DateTime(1904, 1, 2, 12, 0, 0) : new DateTime(1900, 1, 1, 12, 0, 0), actual);
        Assert.True(summary.IsSuccess);
    }

    [Fact]
    public void Import_Errors_ShouldDeliverNoEntity_AndTruncateSamplesWithoutStopping()
    {
        using var source = new MemoryStream(Package(Header("Count") + NumericRows("bad", "worse", "3")));
        var rows = new List<ExcelImportRow<NumberRow>>();
        var summary = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<NumberRow>(0), source,
            row => { rows.Add(row); return ExcelRowReadDecision.Continue; }, new() { MaxErrorSamples = 1 });
        Assert.Equal(3, rows.Count); Assert.Null(rows[0].Item); Assert.Null(rows[1].Item);
        Assert.Equal(3, rows[2].Item.Count);
        Assert.True(summary.IsComplete); Assert.False(summary.IsSuccess); Assert.True(summary.ErrorsTruncated);
        Assert.Equal(2, summary.TotalErrors); Assert.Equal(2, summary.FailedRows); Assert.Equal(1, summary.SuccessfulRows);
        var error = Assert.Single(summary.Errors);
        Assert.Equal(2, error.RowIndex); Assert.Equal(1, error.ColumnIndex);
        Assert.Equal("数据", error.SheetName); Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(ImportStopReason.None, summary.StopReason);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Import_ShouldSeparateErrorLimitAndConsumerStop(bool errorLimit)
    {
        using var source = new MemoryStream(Package(Header("Count") + NumericRows(errorLimit ? "bad" : "1", "2")));
        var summary = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<NumberRow>(0), source,
            _ => errorLimit ? ExcelRowReadDecision.Continue : ExcelRowReadDecision.Stop, new() { MaxErrors = 1 });
        Assert.False(summary.IsComplete); Assert.False(summary.IsSuccess); Assert.Equal(1, summary.DeliveredRows);
        Assert.Equal(errorLimit ? ImportStopReason.ErrorLimit : ImportStopReason.ConsumerStopped, summary.StopReason);
        Assert.Equal(errorLimit ? 1 : 0, summary.TotalErrors);
    }

    [Fact]
    public async Task ImportAsync_ShouldUseAsyncInput_AndAwaitConsumerBeforeNextRow()
    {
        using var source = new AsyncOnlyInput(Package(Header("Count") + NumericRows("1", "2")));
        var first = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var release = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var calls = 0;
        var task = new NpoiExcelStreamingImporter().ImportAsync(ExcelStreamingImport.Sheet<NumberRow>(0), source,
            async (row, _) =>
            {
                calls++;
                if (calls == 1) { first.SetResult(true); await release.Task; }
                Assert.Equal(calls, row.Item.Count);
                return ExcelRowReadDecision.Continue;
            });
        await first.Task.WaitAsync(TimeSpan.FromSeconds(10));
        Assert.Equal(1, calls); Assert.False(task.IsCompleted);
        release.SetResult(true);
        Assert.True((await task).IsSuccess); Assert.Equal(2, calls); Assert.True(source.CanRead);
    }

    [Fact]
    public async Task ImportAsync_ShouldPropagateConsumerExceptionAndCancellation_AndCleanTemporaryFiles()
    {
        var folder = Path.Combine(Path.GetTempPath(), "bing-import-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            var importer = new NpoiExcelStreamingImporter();
            var options = new ExcelStreamingImportOptions { TempDirectory = folder, SharedStringMemoryBytes = 1 };
            var expected = new IOException("consumer failure");
            using var source = new MemoryStream(Package(Header("Count") + NumericRows("1", "2"), "<si><t>共享字符串</t></si>"));
            var actual = await Assert.ThrowsAsync<IOException>(() => importer.ImportAsync(ExcelStreamingImport.Sheet<NumberRow>(0), source,
                (_, _) => throw expected, options));
            Assert.Same(expected, actual); Assert.Empty(Directory.GetFileSystemEntries(folder));
            source.Position = 0;
            using var cancellation = new CancellationTokenSource();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => importer.ImportAsync(ExcelStreamingImport.Sheet<NumberRow>(0), source,
                (_, _) => { cancellation.Cancel(); return Task.FromResult(ExcelRowReadDecision.Continue); }, options, cancellation.Token));
            Assert.Empty(Directory.GetFileSystemEntries(folder)); Assert.True(source.CanRead);
        }
        finally { Directory.Delete(folder); }
    }

    [Fact]
    public void Import_ShouldRejectMissingFormulaCacheAsRowError_AndContinue()
    {
        using var source = new MemoryStream(Package(Header("Count") +
            "<row r=\"2\"><c r=\"A2\"><f>1+1</f></c></row><row r=\"3\"><c r=\"A3\"><v>3</v></c></row>"));
        var rows = new List<ExcelImportRow<NumberRow>>();
        var summary = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<NumberRow>(0), source,
            row => { rows.Add(row); return ExcelRowReadDecision.Continue; });
        Assert.Null(rows[0].Item); Assert.Equal(2, Assert.Single(rows[0].Errors).RowIndex);
        Assert.Equal(3, rows[1].Item.Count); Assert.True(summary.IsComplete); Assert.False(summary.IsSuccess);
    }

    internal static string Header(params string[] values) => "<row r=\"1\">" + string.Concat(values.Select((value, index) =>
        $"<c r=\"{(char)('A' + index)}1\" t=\"inlineStr\"><is><t>{value}</t></is></c>")) + "</row>";
    internal static string NumericRows(params string[] values) => string.Concat(values.Select((value, index) =>
        $"<row r=\"{index + 2}\"><c r=\"A{index + 2}\"><v>{value}</v></c></row>"));

    internal static byte[] Package(string rows, string shared = null, bool date1904 = false)
    {
        using var result = new MemoryStream();
        using (var zip = new ZipArchive(result, ZipArchiveMode.Create, true))
        {
            const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
            const string packageRel = "http://schemas.openxmlformats.org/package/2006/relationships";
            // 非标准部件路径确保实现按真实关系查找，不硬编码 xl/workbook.xml。
            Write("_rels/.rels", $"<Relationships xmlns=\"{packageRel}\"><Relationship Id=\"root\" Type=\"{rel}officeDocument\" Target=\"book/main.xml\"/></Relationships>");
            Write("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/book/main.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/></Types>");
            Write("book/main.xml", $"<workbook xmlns=\"{main}\" xmlns:r=\"{rel.TrimEnd('/')}\"><workbookPr date1904=\"{(date1904 ? 1 : 0)}\"/><sheets><sheet name=\"数据\" sheetId=\"1\" r:id=\"sheet\"/></sheets></workbook>");
            Write("book/_rels/main.xml.rels", $"<Relationships xmlns=\"{packageRel}\"><Relationship Id=\"sheet\" Type=\"{rel}worksheet\" Target=\"../data/rows.xml\"/><Relationship Id=\"style\" Type=\"{rel}styles\" Target=\"style.xml\"/>" +
                (shared == null ? "" : $"<Relationship Id=\"strings\" Type=\"{rel}sharedStrings\" Target=\"strings.xml\"/>") + "</Relationships>");
            Write("data/rows.xml", $"<worksheet xmlns=\"{main}\"><sheetData>{rows}</sheetData></worksheet>");
            Write("book/style.xml", $"<styleSheet xmlns=\"{main}\"><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\"/></cellXfs></styleSheet>");
            if (shared != null) Write("book/strings.xml", $"<sst xmlns=\"{main}\">{shared}</sst>");
            void Write(string path, string text)
            { using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false)); writer.Write(text); }
        }
        return result.ToArray();
    }

    public class Row { public int Count { get; set; } public string Name { get; set; } }
    public class NumberRow { public int Count { get; set; } }
    public class DateRow { public DateTime When { get; set; } }
    public class UniqueRow { [ExcelUnique] public string Code { get; set; } }
    public class DynamicRow
    {
        public string Code { get; set; }
        [DynamicColumn] public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }
    private sealed class AsyncOnlyInput : MemoryStream
    {
        internal AsyncOnlyInput(byte[] bytes) : base(bytes) { }
        public override int Read(byte[] buffer, int offset, int count) => throw new InvalidOperationException("必须异步读取输入");
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
        { token.ThrowIfCancellationRequested(); return Task.FromResult(base.Read(buffer, offset, count)); }
    }
}
