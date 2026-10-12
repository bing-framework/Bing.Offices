using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证透视表来源公式、类型、缓存和资源行为的公共合同。
/// </summary>
public sealed class PivotTableSourceContractTest
{
    /// <summary>
    /// 获取应在来源预检阶段拒绝的公式位置、函数和缓存状态。
    /// </summary>
    public static IEnumerable<object[]> FormulaSourceCases
    {
        get
        {
            var cases = new[]
            {
                new FormulaSourceCase("C5", "1+2", "3", false, 5, 3),
                new FormulaSourceCase("D5", "NO_SUCH_FUNCTION(1)", null, false, 5, 4),
                new FormulaSourceCase("E6", "1+2", "3", false, 6, 5),
                new FormulaSourceCase("C4", "1+2", "3", false, 4, 3),
                new FormulaSourceCase("D5", "1+2", "3", false, 5, 4),
                new FormulaSourceCase("D5", "1+2", "999", false, 5, 4),
                new FormulaSourceCase("D5", "1+2", null, false, 5, 4),
                new FormulaSourceCase("E6", "NO_SUCH_FUNCTION(1)", "#NAME?", true, 6, 5)
            };

            foreach (var item in cases)
            {
                yield return new object[]
                {
                    item.Address, item.Formula, item.CachedValue, item.CachedError, item.Row, item.Column, false
                };
                yield return new object[]
                {
                    item.Address, item.Formula, item.CachedValue, item.CachedError, item.Row, item.Column, true
                };
            }
        }
    }

    /// <summary>
    /// 验证支持函数、未知函数及各种模板缓存状态的公式均在取值前拒绝。
    /// </summary>
    /// <param name="address">模板来源区域内的公式单元格。</param>
    /// <param name="formula">公式文本。</param>
    /// <param name="cachedValue">可选的原始缓存文本。</param>
    /// <param name="cachedError">缓存是否是 Excel 错误类型。</param>
    /// <param name="row">预期的一基来源行。</param>
    /// <param name="column">预期的一基来源列。</param>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [MemberData(nameof(FormulaSourceCases))]
    public async Task FormulaInSource_ShouldRejectBeforeEvaluation(string address, string formula,
        string cachedValue, bool cachedError, int row, int column, bool asynchronous)
    {
        var bytes = WithFormula(CreateTemplate(), address, formula, cachedValue, cachedError);
        using var template = new MemoryStream(bytes, writable: false);
        using var output = SeededStream();
        var original = output.ToArray();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        var request = CreateRequest(template);

        var exception = await CaptureConfigurationException(() => Export(exporter, request, output, asynchronous));

        AssertConfigurationLocation(exception, row, column);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证带空格的字段名按原文匹配并写入完整缓存。
    /// </summary>
    /// <remarks>
    /// 表头与字段配置都保留首尾空格，并检查缓存来源记录和透视字段索引。
    /// </remarks>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PaddedHeaderFieldNames_ShouldBindExactlyAndPreserveCache(bool asynchronous)
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("C4").Value = " Region ";
            sheet.Cell("D4").Value = " Amount ";
        }), writable: false);
        using var output = new MemoryStream();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        var definition = CreateDefinition(rowField: " Region ", valueField: " Amount ");

        await Export(exporter, CreateRequest(template, definition), output, asynchronous);

        Assert.Empty(observer.Exceptions);
        AssertCompleteSourceAndCache(output.ToArray(), expectedHeaders: new[]
            { " Region ", " Amount ", "Memo" });
        using var archive = new ZipArchive(new MemoryStream(output.ToArray(), writable: false), ZipArchiveMode.Read);
        var pivotPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var pivot = ReadXml(archive, pivotPath).Root;
        var rowField = Assert.Single(pivot.Element(SpreadsheetNs + "rowFields").Elements());
        var dataField = Assert.Single(pivot.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField"));
        Assert.Equal("0", (string)rowField.Attribute("x"));
        Assert.Equal("1", (string)dataField.Attribute("fld"));
    }

    /// <summary>
    /// 验证字段匹配不会去除表头两端的空格。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnpaddedFieldNames_ShouldNotMatchPaddedHeaders(bool asynchronous)
    {
        using var template = new MemoryStream(CreateTemplate(sheet =>
        {
            sheet.Cell("C4").Value = " Region ";
            sheet.Cell("D4").Value = " Amount ";
        }), writable: false);
        using var output = SeededStream();
        var original = output.ToArray();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        var exception = await CaptureConfigurationException(() =>
            Export(exporter, CreateRequest(template), output, asynchronous));

        AssertConfigurationLocation(exception, 4, 3);
        Assert.Contains("Region", exception.Message);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证后续透视来源失败时保留调用方流和已有文件。
    /// </summary>
    /// <remarks>
    /// 第一个来源有效，第二个来源在 G5 含错误值；失败后只保留原目标文件。
    /// </remarks>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LaterPivotSourceFailure_ShouldPreserveStreamAndFile(bool asynchronous)
    {
        var original = new byte[] { 9, 2, 7, 1, 8 };
        using var streamTemplate = new MemoryStream(CreateMultiPivotTemplate(), writable: false);
        using var output = SeededStream();
        var originalStream = output.ToArray();
        var streamObserver = new RecordingObserver();
        var streamExporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { streamObserver });

        var streamException = await CaptureConfigurationException(() => Export(streamExporter,
            CreateMultiPivotRequest(streamTemplate), output, asynchronous));

        AssertConfigurationLocation(streamException, 5, 7, "SalesByMemo");
        Assert.Same(streamException, Assert.Single(streamObserver.Exceptions));
        Assert.Equal(originalStream, output.ToArray());
        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        Assert.True(streamTemplate.CanRead);

        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        File.WriteAllBytes(target, original);
        using var fileTemplate = new MemoryStream(CreateMultiPivotTemplate(), writable: false);
        var fileObserver = new RecordingObserver();
        var fileExporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { fileObserver });

        var fileException = await CaptureFileConfigurationException(() => ExportToFile(fileExporter,
            CreateMultiPivotRequest(fileTemplate), target, asynchronous));

        AssertConfigurationLocation(fileException, 5, 7, "SalesByMemo");
        Assert.Same(fileException, Assert.Single(fileObserver.Exceptions));
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        Assert.True(fileTemplate.CanRead);
    }

    /// <summary>
    /// 验证区域外公式保留，纯值来源与全部缓存字段及记录保持一致。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PureValues_ShouldWriteCompleteSourceAndCache(bool asynchronous)
    {
        var bytes = WithFormula(CreateTemplate(sheet => sheet.Cell("H10").Value = "outside"),
            "H10", "1+2", "3", false);
        using var template = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        await Export(exporter, CreateRequest(template), output, asynchronous);

        Assert.Empty(observer.Exceptions);
        Assert.True(template.CanRead);
        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        AssertCompleteSourceAndCache(output.ToArray());
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        Assert.Equal("1+2", workbook.Worksheet("Data").Cell("H10").FormulaA1);
        Assert.Equal("Template sentinel", workbook.Worksheet("Data").Cell("A20").GetString());
    }

    /// <summary>
    /// 验证混合类型分组在透视来源和缓存记录中保留原始类型。
    /// </summary>
    /// <remarks>
    /// 来源包含数值、日期和布尔分组；断言同时检查重新打开的单元格类型与缓存字段值类型。
    /// </remarks>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TypedGroups_ShouldPreserveSourceAndCacheTypes(bool asynchronous)
    {
        var bytes = CreateTemplate(sheet =>
        {
            sheet.Cell("C5").Value = 42;
            sheet.Cell("C6").Value = new DateTime(2026, 10, 5);
            sheet.Cell("C7").Value = true;
        });
        using var template = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        await Export(exporter, CreateRequest(template), output, asynchronous);

        Assert.Empty(observer.Exceptions);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var source = workbook.Worksheet("Data");
        Assert.Equal(XLDataType.Number, source.Cell("C5").DataType);
        Assert.Equal(XLDataType.DateTime, source.Cell("C6").DataType);
        Assert.Equal(XLDataType.Boolean, source.Cell("C7").DataType);
        Assert.Equal("10", source.Cell("D5").GetString());
        Assert.Equal("", source.Cell("D6").GetString());
        Assert.Equal("20", source.Cell("D7").GetString());

        using var archive = new ZipArchive(new MemoryStream(output.ToArray(), writable: false), ZipArchiveMode.Read);
        var cachePath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheDefinition1.xml", StringComparison.Ordinal)).FullName;
        var recordsPath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheRecords1.xml", StringComparison.Ordinal)).FullName;
        var cache = ReadXml(archive, cachePath).Root;
        var fields = cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        Assert.Equal(new[] { "Region", "Amount", "Memo" },
            fields.Select(field => (string)field.Attribute("name")));
        var records = ReadXml(archive, recordsPath).Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, records.Length);
        Assert.Equal(new[] { "n", "d", "b" }, records.Select(record =>
            CacheCellType(record.Elements().First(), fields[0])));
        Assert.Equal(new[] { "10", "", "20" }, records.Select(record =>
            DecodeCacheCell(record.Elements().ElementAt(1), fields[1])));
    }

    /// <summary>
    /// 验证小写字段名仍可绑定原始表头，并保留分组值的缓存类型。
    /// </summary>
    /// <remarks>
    /// 该用例同时检查数值、日期、布尔分组值及金额字段，确保字段名匹配不改变源记录。
    /// </remarks>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LowercaseFieldNames_ShouldSelectTypedSourceAndCache(bool asynchronous)
    {
        var bytes = CreateTemplate(sheet =>
        {
            sheet.Cell("C5").Value = 42;
            sheet.Cell("C6").Value = new DateTime(2026, 10, 5);
            sheet.Cell("C7").Value = true;
        });
        using var template = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        await Export(exporter, CreateRequest(template, CreateDefinition(rowField: "region", valueField: "amount")),
            output, asynchronous);

        Assert.Empty(observer.Exceptions);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray(), writable: false), ZipArchiveMode.Read);
        var pivotPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var cachePath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheDefinition1.xml", StringComparison.Ordinal)).FullName;
        var recordsPath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheRecords1.xml", StringComparison.Ordinal)).FullName;
        var pivot = ReadXml(archive, pivotPath).Root;
        var rowField = Assert.Single(pivot.Element(SpreadsheetNs + "rowFields").Elements());
        var dataField = Assert.Single(pivot.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField"));
        Assert.Equal("0", (string)rowField.Attribute("x"));
        Assert.Equal("1", (string)dataField.Attribute("fld"));

        var cache = ReadXml(archive, cachePath).Root;
        var fields = cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        Assert.Equal(new[] { "Region", "Amount", "Memo" },
            fields.Select(field => (string)field.Attribute("name")));
        var records = ReadXml(archive, recordsPath).Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, records.Length);
        Assert.All(records, record => Assert.Equal(3, record.Elements().Count()));
        Assert.Equal(new[] { "n", "d", "b" }, records.Select(record =>
            CacheCellType(record.Elements().First(), fields[0])));
        Assert.Equal(new[] { "42", "2026-10-05T00:00:00", "1" }, records.Select(record =>
            DecodeCacheCell(record.Elements().First(), fields[0])));
        Assert.Equal(new[] { "10", "", "20" }, records.Select(record =>
            DecodeCacheCell(record.Elements().ElementAt(1), fields[1])));
        Assert.Equal(new[] { "Fresh", "Empty", "Special" }, records.Select(record =>
            DecodeCacheCell(record.Elements().ElementAt(2), fields[2])));

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        var source = workbook.Worksheet("Data");
        Assert.Equal(XLDataType.Number, source.Cell("C5").DataType);
        Assert.Equal(XLDataType.DateTime, source.Cell("C6").DataType);
        Assert.Equal(XLDataType.Boolean, source.Cell("C7").DataType);
        Assert.Equal(new[] { "Region", "Amount", "Memo" },
            source.Range("C4:E4").Cells().Select(cell => cell.GetString()));
    }

    /// <summary>
    /// 验证大小写不同的文本分组按现有比较规则合并且完整来源仍进缓存。
    /// </summary>
    /// <remarks>
    /// 将输出起点放在恰好容纳两个分组的位置；若 North 与 NORTH 未去重，预算预检必须失败。
    /// </remarks>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CaseInsensitiveGroups_ShouldKeepCompleteSourceCacheAtExactBoundary(bool asynchronous)
    {
        var bytes = CreateTemplate(sheet =>
        {
            sheet.Cell("C6").Value = "NORTH";
            sheet.Cell("D6").Value = 15;
            sheet.Cell("E6").Value = "Uppercase";
        });
        using var template = new MemoryStream(bytes, writable: false);
        using var output = new MemoryStream();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        await Export(exporter, CreateRequest(template, CreateDefinition(targetRow: 1048572)),
            output, asynchronous);

        Assert.Empty(observer.Exceptions);
        AssertCompleteSourceAndCache(output.ToArray(), new[] { "North", "NORTH", "South" },
            new[] { "10", "15", "20" }, new[] { "Fresh", "Uppercase", "Special" }, "C1048573");
    }

    /// <summary>
    /// 验证不符合透视来源值合同的数据返回精确坐标并保全调用方流。
    /// </summary>
    /// <param name="scenario">要注入的来源错误类别。</param>
    /// <param name="row">预期的一基来源行。</param>
    /// <param name="column">预期的一基来源列。</param>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData("all-empty-sum", 5, 4, false)]
    [InlineData("all-empty-sum", 5, 4, true)]
    [InlineData("empty-row", 6, 3, false)]
    [InlineData("empty-row", 6, 3, true)]
    [InlineData("numeric-text", 5, 4, false)]
    [InlineData("numeric-text", 5, 4, true)]
    [InlineData("boolean-sum", 5, 4, false)]
    [InlineData("boolean-sum", 5, 4, true)]
    [InlineData("date-sum", 5, 4, false)]
    [InlineData("date-sum", 5, 4, true)]
    [InlineData("excel-error", 5, 4, false)]
    [InlineData("excel-error", 5, 4, true)]
    [InlineData("duplicate-header", 4, 5, false)]
    [InlineData("duplicate-header", 4, 5, true)]
    [InlineData("blank-header", 4, 5, false)]
    [InlineData("blank-header", 4, 5, true)]
    [InlineData("whitespace-header", 4, 5, false)]
    [InlineData("whitespace-header", 4, 5, true)]
    [InlineData("case-only-duplicate-header", 4, 5, false)]
    [InlineData("case-only-duplicate-header", 4, 5, true)]
    [InlineData("excel-error-header", 4, 3, false)]
    [InlineData("excel-error-header", 4, 3, true)]
    [InlineData("excel-error-group", 5, 3, false)]
    [InlineData("excel-error-group", 5, 3, true)]
    [InlineData("excel-error-unselected", 6, 5, false)]
    [InlineData("excel-error-unselected", 6, 5, true)]
    public async Task InvalidSource_ShouldReportCoordinateAndPreserveOutput(string scenario, int row,
        int column, bool asynchronous)
    {
        using var template = new MemoryStream(CreateInvalidTemplate(scenario), writable: false);
        using var output = SeededStream();
        var original = output.ToArray();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        var exception = await CaptureConfigurationException(() =>
            Export(exporter, CreateRequest(template), output, asynchronous));

        AssertConfigurationLocation(exception, row, column);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证透视展开预算错误指向造成越界的来源分组单元格。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExpandedOutputBudget_ShouldReportSourceCoordinate(bool asynchronous)
    {
        using var template = new MemoryStream(CreateTemplate(), writable: false);
        using var output = SeededStream();
        var original = output.ToArray();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        var request = CreateRequest(template, CreateDefinition(targetRow: 1048572));

        var exception = await CaptureConfigurationException(() => Export(exporter, request, output, asynchronous));

        AssertConfigurationLocation(exception, 7, 3);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证来源公式文件预检失败保留旧目标并清理同目录暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_FormulaSourceFailure_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 6, 2, 6, 4, 3 };
        File.WriteAllBytes(target, original);
        using var template = new MemoryStream(
            WithFormula(CreateTemplate(), "D5", "1+2", "3", false), writable: false);
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        var exception = await CaptureFileConfigurationException(() =>
            ExportToFile(exporter, CreateRequest(template), target, asynchronous));

        AssertConfigurationLocation(exception, 5, 4);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证预取消请求不会触发来源公式预检或修改输出流。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreCanceledSourceExport_ShouldPreserveStreamAndTemplate(bool asynchronous)
    {
        using var template = new MemoryStream(
            WithFormula(CreateTemplate(), "D5", "NO_SUCH_FUNCTION(1)", null, false), writable: false);
        using var output = SeededStream();
        var original = output.ToArray();
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Export(exporter, CreateRequest(template), output, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证来源数据生成期间取消会保留调用方流及真实文件目标。
    /// </summary>
    /// <param name="asynchronous">是否使用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SourceGenerationCancellation_ShouldPreserveStreamAndFile(bool asynchronous)
    {
        using var cancellation = new CancellationTokenSource();
        var request = CreateCancelDuringGenerationRequest(cancellation);
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        using var output = SeededStream();
        var original = output.ToArray();

        var streamException = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Export(exporter, request, output, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, streamException.CancellationToken);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanRead);
        Assert.Empty(observer.Exceptions);

        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        File.WriteAllBytes(target, original);
        using var fileCancellation = new CancellationTokenSource();
        var fileRequest = CreateCancelDuringGenerationRequest(fileCancellation);
        var fileObserver = new RecordingObserver();
        var fileExporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { fileObserver });

        var fileException = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(fileExporter, fileRequest, target, asynchronous, fileCancellation.Token));

        Assert.Equal(fileCancellation.Token, fileException.CancellationToken);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        Assert.Empty(fileObserver.Exceptions);
    }

    /// <summary>
    /// 验证暂存工作簿实际写入期间取消会保留旧目标并清理临时文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileWriteCancellation_ShouldPreserveTargetAndCleanTemporaryFile(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 8, 5, 3, 0, 9 };
        File.WriteAllBytes(target, original);
        using var template = new MemoryStream(CreateTemplate(), writable: false);
        var observer = new RecordingObserver();
        var committer = new CancelDuringWriteCommitter(cancellation);
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer },
            fileExportCommitter: committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(template), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.True(committer.BytesWritten > 0);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证目标工作簿包含完整来源范围、字段和三行缓存记录。
    /// </summary>
    /// <param name="bytes">导出的 XLSX 字节。</param>
    /// <param name="expectedGroups">预期的分组字段值。</param>
    /// <param name="expectedAmounts">预期的求和字段值。</param>
    /// <param name="expectedMemos">预期的未选字段值。</param>
    /// <param name="expectedPivotLocation">预期透视表的左上角单元格。</param>
    /// <param name="expectedHeaders">预期写入缓存的来源表头。</param>
    private static void AssertCompleteSourceAndCache(byte[] bytes,
        string[] expectedGroups = null, string[] expectedAmounts = null, string[] expectedMemos = null,
        string expectedPivotLocation = "C2", string[] expectedHeaders = null)
    {
        using var archive = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var pivotPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var cachePath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheDefinition1.xml", StringComparison.Ordinal)).FullName;
        var recordsPath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheRecords1.xml", StringComparison.Ordinal)).FullName;
        var pivot = ReadXml(archive, pivotPath).Root;
        var cache = ReadXml(archive, cachePath).Root;
        var records = ReadXml(archive, recordsPath).Root;

        Assert.Equal("SalesByRegion", (string)pivot.Attribute("name"));
        Assert.StartsWith(expectedPivotLocation,
            (string)pivot.Element(SpreadsheetNs + "location").Attribute("ref"));
        var source = cache.Element(SpreadsheetNs + "cacheSource")
            .Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Data", (string)source.Attribute("sheet"));
        Assert.Equal("C4:E7", (string)source.Attribute("ref"));
        var fields = cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        expectedHeaders ??= new[] { "Region", "Amount", "Memo" };
        Assert.Equal(expectedHeaders,
            fields.Select(field => (string)field.Attribute("name")));

        var rows = records.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, row => Assert.Equal(3, row.Elements().Count()));
        var values = rows.Select(row => row.Elements().Select((cell, index) =>
            DecodeCacheCell(cell, fields[index])).ToArray()).ToArray();
        expectedGroups ??= new[] { "North", "", "South" };
        expectedAmounts ??= new[] { "10", "", "20" };
        expectedMemos ??= new[] { "Fresh", "Empty", "Special" };
        Assert.Equal(expectedGroups, values.Select(row => row[0]));
        Assert.Equal(expectedAmounts, values.Select(row => row[1]));
        Assert.Equal(expectedMemos, values.Select(row => row[2]));

        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        var sourceSheet = workbook.Worksheet("Data");
        Assert.Equal(expectedHeaders,
            sourceSheet.Range("C4:E4").Cells().Select(cell => cell.GetString()));
        Assert.Equal(expectedGroups[0], sourceSheet.Cell("C5").GetString());
        Assert.Equal(expectedAmounts[0], sourceSheet.Cell("D5").GetString());
        Assert.Equal(expectedMemos[0], sourceSheet.Cell("E5").GetString());
        Assert.Equal(expectedGroups[1], sourceSheet.Cell("C6").GetString());
        Assert.Equal(expectedAmounts[1], sourceSheet.Cell("D6").GetString());
        Assert.Equal(expectedMemos[1], sourceSheet.Cell("E6").GetString());
        Assert.Equal(expectedGroups[2], sourceSheet.Cell("C7").GetString());
        Assert.Equal(expectedAmounts[2], sourceSheet.Cell("D7").GetString());
        Assert.Equal(expectedMemos[2], sourceSheet.Cell("E7").GetString());
    }

    /// <summary>
    /// 获取透视缓存记录单元格的有效 OOXML 值类型。
    /// </summary>
    /// <param name="cell">记录中的来源字段值。</param>
    /// <param name="cacheField">字段对应的共享缓存定义。</param>
    /// <returns>共享项或记录单元格的 SpreadsheetML 元素类型。</returns>
    private static string CacheCellType(XElement cell, XElement cacheField)
    {
        if (cell.Name.LocalName != "x")
            return cell.Name.LocalName;
        var sharedItems = cacheField.Element(SpreadsheetNs + "sharedItems").Elements().ToArray();
        var index = int.Parse((string)cell.Attribute("v"), System.Globalization.CultureInfo.InvariantCulture);
        return sharedItems[index].Name.LocalName;
    }

    /// <summary>
    /// 将透视请求中的公式缓存单元格解码为完整来源值。
    /// </summary>
    /// <param name="cell">缓存记录中的字段值。</param>
    /// <param name="cacheField">对应的缓存字段定义。</param>
    /// <returns>字段值文本；空值返回空文本。</returns>
    private static string DecodeCacheCell(XElement cell, XElement cacheField)
    {
        if (cell.Name.LocalName == "m")
            return string.Empty;
        var value = (string)cell.Attribute("v");
        if (cell.Name.LocalName != "x")
            return value ?? cell.Value;
        var sharedItems = cacheField.Element(SpreadsheetNs + "sharedItems").Elements().ToArray();
        var index = int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        var item = sharedItems[index];
        return (string)item.Attribute("v") ?? item.Value;
    }

    /// <summary>
    /// 构造一基坐标、阶段、错误码和透视标识均符合合同的异常断言。
    /// </summary>
    /// <param name="exception">来源预检抛出的配置异常。</param>
    /// <param name="row">预期的一基行号。</param>
    /// <param name="column">预期的一基列号。</param>
    /// <param name="pivotName">预期出现在错误消息中的透视表名称。</param>
    private static void AssertConfigurationLocation(BingOfficesConfigurationException exception, int row,
        int column, string pivotName = "SalesByRegion")
    {
        Assert.NotNull(exception);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal("Data", exception.SheetName);
        Assert.Equal(row, exception.RowIndex);
        Assert.Equal(column, exception.ColumnIndex);
        Assert.Contains(pivotName, exception.Message);
    }

    /// <summary>
    /// 捕获同步或异步透视流导出的配置异常。
    /// </summary>
    /// <param name="export">待执行的导出操作。</param>
    /// <returns>导出失败时返回的配置异常。</returns>
    private static async Task<BingOfficesConfigurationException> CaptureConfigurationException(
        Func<Task> export)
    {
        try
        {
            await export();
            return null;
        }
        catch (BingOfficesConfigurationException exception)
        {
            return exception;
        }
    }

    /// <summary>
    /// 捕获同步或异步透视文件导出的配置异常。
    /// </summary>
    /// <param name="export">待执行的文件导出操作。</param>
    /// <returns>导出失败时返回的配置异常。</returns>
    private static async Task<BingOfficesConfigurationException> CaptureFileConfigurationException(
        Func<Task> export)
    {
        try
        {
            await export();
            return null;
        }
        catch (BingOfficesConfigurationException exception)
        {
            return exception;
        }
    }

    /// <summary>
    /// 按请求选择真实同步或异步流导出入口。
    /// </summary>
    /// <param name="exporter">执行透视导出的 ClosedXML Provider。</param>
    /// <param name="request">透视表工作簿请求。</param>
    /// <param name="output">调用方目标流。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static async Task Export(ClosedXmlExcelExporter exporter, ExcelWorkbookExportRequest request,
        Stream output, bool asynchronous, CancellationToken cancellationToken = default)
    {
        if (asynchronous)
            await exporter.ExportAsync(request, output, cancellationToken);
        else
            exporter.Export(request, output, cancellationToken);
    }

    /// <summary>
    /// 按请求选择真实同步或异步文件导出入口。
    /// </summary>
    /// <param name="exporter">执行透视导出的 ClosedXML Provider。</param>
    /// <param name="request">透视表工作簿请求。</param>
    /// <param name="target">目标文件路径。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="cancellationToken">取消令牌。</param>
    private static async Task ExportToFile(ClosedXmlExcelExporter exporter,
        ExcelWorkbookExportRequest request, string target, bool asynchronous,
        CancellationToken cancellationToken = default)
    {
        if (asynchronous)
            await exporter.ExportToFileAsync(request, target, cancellationToken);
        else
            exporter.ExportToFile(request, target, cancellationToken);
    }

    /// <summary>
    /// 创建模板模式的透视工作簿请求。
    /// </summary>
    /// <param name="template">包含偏移来源区域的模板流。</param>
    /// <param name="definition">可选的透视定义。</param>
    /// <returns>已完成配置的公开工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template,
        ExcelPivotTableDefinition definition = null) => ExcelExport.Workbook(builder => builder
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", Array.Empty<PivotSourceRow>())
            .PivotTable(definition ?? CreateDefinition()));

    /// <summary>
    /// 创建包含两个透视定义的模板导出请求。
    /// </summary>
    /// <param name="template">包含两个来源区域的模板流。</param>
    /// <returns>按顺序处理两个透视来源的导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateMultiPivotRequest(Stream template) => ExcelExport.Workbook(builder =>
        builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Data", Array.Empty<PivotSourceRow>())
            .PivotTable(CreateDefinition())
            .PivotTable(CreateDefinition(name: "SalesByMemo", targetSheet: "Summary2",
                sourceStartColumn: 6, sourceEndColumn: 8)));

    /// <summary>
    /// 创建第二个来源区域在分组字段包含 Excel 错误的模板。
    /// </summary>
    /// <returns>包含两个来源区域的 XLSX 模板字节。</returns>
    private static byte[] CreateMultiPivotTemplate() => CreateTemplate(sheet =>
    {
        sheet.Cell("G4").Value = "Region";
        sheet.Cell("H4").Value = "Amount";
        sheet.Cell("I4").Value = "Memo";
        sheet.Cell("G5").Value = XLError.DivisionByZero;
        sheet.Cell("H5").Value = 5;
        sheet.Cell("I5").Value = "First";
        sheet.Cell("G6").Value = "West";
        sheet.Cell("H6").Value = 6;
        sheet.Cell("I6").Value = "Second";
        sheet.Cell("G7").Value = "South";
        sheet.Cell("H7").Value = 7;
        sheet.Cell("I7").Value = "Third";
    });

    /// <summary>
    /// 创建从模板 C4:E7 读取来源数据的透视定义。
    /// </summary>
    /// <param name="targetRow">透视结果的零基起始行。</param>
    /// <param name="rowField">分组字段名称。</param>
    /// <param name="valueField">求和字段名称。</param>
    /// <param name="name">透视表名称。</param>
    /// <param name="targetSheet">透视结果工作表名称。</param>
    /// <param name="sourceStartColumn">来源区域的零基起始列号。</param>
    /// <param name="sourceEndColumn">来源区域的零基结束列号。</param>
    /// <returns>包含来源范围和字段声明的透视定义。</returns>
    private static ExcelPivotTableDefinition CreateDefinition(int targetRow = 1,
        string rowField = "Region", string valueField = "Amount", string name = "SalesByRegion",
        string targetSheet = "Summary", int sourceStartColumn = 2, int sourceEndColumn = 4) => new()
    {
        Name = name,
        SourceSheet = "Data",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 3,
            StartColumn = sourceStartColumn,
            EndRow = 6,
            EndColumn = sourceEndColumn
        },
        TargetSheet = targetSheet,
        TargetRow = targetRow,
        TargetColumn = 2,
        RowField = rowField,
        ValueField = valueField
    };

    /// <summary>
    /// 创建含中文工作表、偏移来源区及空分组与空金额的模板。
    /// </summary>
    /// <param name="configure">可选的模板单元格覆盖操作。</param>
    /// <returns>完整模板工作簿的 XLSX 字节。</returns>
    private static byte[] CreateTemplate(Action<IXLWorksheet> configure = null)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("C4").Value = "Region";
        sheet.Cell("D4").Value = "Amount";
        sheet.Cell("E4").Value = "Memo";
        sheet.Cell("C5").Value = "North";
        sheet.Cell("D5").Value = 10;
        sheet.Cell("E5").Value = "Fresh";
        sheet.Cell("C6").Value = string.Empty;
        sheet.Cell("E6").Value = "Empty";
        sheet.Cell("C7").Value = "South";
        sheet.Cell("D7").Value = 20;
        sheet.Cell("E7").Value = "Special";
        sheet.Cell("A20").Value = "Template sentinel";
        configure?.Invoke(sheet);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建指定类型的无效来源模板。
    /// </summary>
    /// <param name="scenario">要注入的来源数据错误类型。</param>
    /// <returns>含指定错误的模板 XLSX 字节。</returns>
    private static byte[] CreateInvalidTemplate(string scenario) => CreateTemplate(sheet =>
    {
        switch (scenario)
        {
            case "all-empty-sum":
                sheet.Cell("D5").Clear(XLClearOptions.Contents);
                sheet.Cell("D7").Clear(XLClearOptions.Contents);
                break;
            case "empty-row":
                sheet.Range("C6:E6").Clear(XLClearOptions.Contents);
                break;
            case "numeric-text":
                sheet.Cell("D5").Value = "10";
                break;
            case "boolean-sum":
                sheet.Cell("D5").Value = true;
                break;
            case "date-sum":
                sheet.Cell("D5").Value = new DateTime(2026, 10, 5);
                break;
            case "excel-error":
                sheet.Cell("D5").Value = XLError.DivisionByZero;
                break;
            case "duplicate-header":
                sheet.Cell("E4").Value = "Region";
                break;
            case "blank-header":
                sheet.Cell("E4").Clear(XLClearOptions.Contents);
                break;
            case "whitespace-header":
                sheet.Cell("E4").Value = "   ";
                break;
            case "case-only-duplicate-header":
                sheet.Cell("E4").Value = "region";
                break;
            case "excel-error-header":
                sheet.Cell("C4").Value = XLError.DivisionByZero;
                break;
            case "excel-error-group":
                sheet.Cell("C5").Value = XLError.DivisionByZero;
                break;
            case "excel-error-unselected":
                sheet.Cell("E6").Value = XLError.DivisionByZero;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(scenario), scenario, "未知来源合同场景。");
        }
    });

    /// <summary>
    /// 在 XLSX 公式单元格中设置原始公式、可选缓存和缓存错误类型。
    /// </summary>
    /// <param name="template">来源模板 XLSX 字节。</param>
    /// <param name="address">公式单元格地址。</param>
    /// <param name="formula">不含等号的公式文本。</param>
    /// <param name="cachedValue">可选的原始缓存文本。</param>
    /// <param name="cachedError">缓存是否具有 Excel 错误单元格类型。</param>
    /// <returns>改写后的模板 XLSX 字节。</returns>
    private static byte[] WithFormula(byte[] template, string address, string formula,
        string cachedValue, bool cachedError)
    {
        using var input = new ZipArchive(new MemoryStream(template, writable: false), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        using (var destination = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var replacement = destination.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                using var sourceStream = entry.Open();
                using var destinationStream = replacement.Open();
                if (entry.FullName == "xl/worksheets/sheet1.xml")
                {
                    var document = XDocument.Load(sourceStream);
                    var cell = document.Descendants(SpreadsheetNs + "c")
                        .Single(item => (string)item.Attribute("r") == address);
                    cell.Attribute("t")?.Remove();
                    cell.Elements(SpreadsheetNs + "f").Remove();
                    cell.Elements(SpreadsheetNs + "v").Remove();
                    cell.Add(new XElement(SpreadsheetNs + "f", formula));
                    if (cachedValue != null)
                    {
                        if (cachedError)
                            cell.SetAttributeValue("t", "e");
                        cell.Add(new XElement(SpreadsheetNs + "v", cachedValue));
                    }
                    document.Save(destinationStream, System.Xml.Linq.SaveOptions.DisableFormatting);
                }
                else
                {
                    sourceStream.CopyTo(destinationStream);
                }
            }
        }
        return output.ToArray();
    }

    /// <summary>
    /// 创建在第二次枚举明细前触发取消的工作簿请求。
    /// </summary>
    /// <param name="cancellation">明细枚举中触发的取消源。</param>
    /// <returns>引用取消明细生成器的透视请求。</returns>
    private static ExcelWorkbookExportRequest CreateCancelDuringGenerationRequest(
        CancellationTokenSource cancellation) => ExcelExport.Workbook(builder => builder
            .AddSheet("Data", CancelAfterFirstRow(cancellation))
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 2,
                    EndColumn = 1
                },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = "Amount"
            }));

    /// <summary>
    /// 在生成第二条来源记录时触发取消。
    /// </summary>
    /// <param name="cancellation">生成期间取消的令牌源。</param>
    /// <returns>按需生成的两条来源记录。</returns>
    private static IEnumerable<PivotSourceRow> CancelAfterFirstRow(CancellationTokenSource cancellation)
    {
        yield return new PivotSourceRow { Region = "North", Amount = 10, Memo = "first" };
        cancellation.Cancel();
        yield return new PivotSourceRow { Region = "South", Amount = 20, Memo = "second" };
    }

    /// <summary>
    /// 创建带已有字节且位于起点的输出流。
    /// </summary>
    /// <returns>包含哨兵内容的可读写内存流。</returns>
    private static MemoryStream SeededStream()
    {
        var output = new MemoryStream();
        output.Write(new byte[] { 7, 11, 19, 23 });
        output.Position = 0;
        return output;
    }

    /// <summary>
    /// 读取 XLSX 压缩包中的 XML 部件。
    /// </summary>
    /// <param name="archive">输入工作簿压缩包。</param>
    /// <param name="path">部件相对路径。</param>
    /// <returns>已加载的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path);
        Assert.NotNull(entry);
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 获取 SpreadsheetML 主命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 透视公式场景的单行测试数据。
    /// </summary>
    private sealed class FormulaSourceCase
    {
        /// <summary>
        /// 初始化一个 <see cref="FormulaSourceCase" /> 类型的实例。
        /// </summary>
        /// <param name="address">公式单元格地址。</param>
        /// <param name="formula">公式文本。</param>
        /// <param name="cachedValue">公式缓存文本。</param>
        /// <param name="cachedError">缓存是否为错误类型。</param>
        /// <param name="row">预期一基行号。</param>
        /// <param name="column">预期一基列号。</param>
        public FormulaSourceCase(string address, string formula, string cachedValue, bool cachedError,
            int row, int column)
        {
            Address = address;
            Formula = formula;
            CachedValue = cachedValue;
            CachedError = cachedError;
            Row = row;
            Column = column;
        }

        /// <summary>
        /// 获取公式单元格地址。
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// 获取公式文本。
        /// </summary>
        public string Formula { get; }

        /// <summary>
        /// 获取可选的公式缓存文本。
        /// </summary>
        public string CachedValue { get; }

        /// <summary>
        /// 获取缓存是否为 Excel 错误类型。
        /// </summary>
        public bool CachedError { get; }

        /// <summary>
        /// 获取预期的一基来源行号。
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 获取预期的一基来源列号。
        /// </summary>
        public int Column { get; }
    }

    /// <summary>
    /// 表示生成透视来源时使用的最小明细行。
    /// </summary>
    private sealed class PivotSourceRow
    {
        /// <summary>
        /// 初始化一个 <see cref="PivotSourceRow" /> 类型的实例。
        /// </summary>
        public PivotSourceRow()
        {
        }

        /// <summary>
        /// 获取或设置来源分组文本。
        /// </summary>
        public string Region { get; set; }

        /// <summary>
        /// 获取或设置来源求和值。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置未选择的缓存来源字段。
        /// </summary>
        public string Memo { get; set; }
    }

    /// <summary>
    /// 提供独立的真实文件导出测试目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 保存当前测试使用的临时目录绝对路径。
        /// </summary>
        private readonly string _path;

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory()
        {
            _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "bing-pivot-source-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_path);
        }

        /// <summary>
        /// 获取测试临时目录绝对路径。
        /// </summary>
        public string Path => _path;

        /// <summary>
        /// 删除本实例创建的临时目录。
        /// </summary>
        public void Dispose()
        {
            if (Directory.Exists(_path))
                Directory.Delete(_path, recursive: true);
        }
    }

    /// <summary>
    /// 记录真实导出期间收到的公共异常通知。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 保存已经收到的公共异常。
        /// </summary>
        private readonly List<BingOfficesException> _exceptions = new();

        /// <summary>
        /// 获取观察到的公共异常列表。
        /// </summary>
        public IReadOnlyList<BingOfficesException> Exceptions => _exceptions;

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => _exceptions.Add(exception);
    }

    /// <summary>
    /// 在真实文件提交暂存流首次写入后取消导出。
    /// </summary>
    private sealed class CancelDuringWriteCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 保存负责暂存和失败清理的默认提交器。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>
        /// 保存首次暂存写入后触发的取消源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 保存当前包装的真实暂存流。
        /// </summary>
        private CancelAfterWriteStream _stream;

        /// <summary>
        /// 初始化一个 <see cref="CancelDuringWriteCommitter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">真实写入时触发的取消源。</param>
        public CancelDuringWriteCommitter(CancellationTokenSource cancellation)
        {
            _cancellation = cancellation;
        }

        /// <summary>
        /// 获取首次取消前已写入暂存文件的字节数。
        /// </summary>
        public long BytesWritten => _stream?.BytesWritten ?? 0;

        /// <summary>
        /// 获取同步提交调用次数。
        /// </summary>
        public int SyncCalls { get; private set; }

        /// <summary>
        /// 获取异步提交调用次数。
        /// </summary>
        public int AsyncCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            SyncCalls++;
            _inner.Commit(path, destination =>
            {
                _stream = new CancelAfterWriteStream(destination, _cancellation);
                write(_stream);
            }, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            return _inner.CommitAsync(path, async (destination, token) =>
            {
                _stream = new CancelAfterWriteStream(destination, _cancellation);
                await writeAsync(_stream, token);
            }, cancellationToken, format);
        }
    }

    /// <summary>
    /// 转发首次真实写入后触发取消且不取得暂存流所有权。
    /// </summary>
    private sealed class CancelAfterWriteStream : Stream
    {
        /// <summary>
        /// 保存由默认文件提交器管理的暂存流。
        /// </summary>
        private readonly Stream _inner;

        /// <summary>
        /// 保存写入后需要触发的取消源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 标记首次转发写入是否已触发取消。
        /// </summary>
        private bool _canceled;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="inner">由提交器拥有的真实暂存流。</param>
        /// <param name="cancellation">首次写入后触发的取消源。</param>
        public CancelAfterWriteStream(Stream inner, CancellationTokenSource cancellation)
        {
            _inner = inner;
            _cancellation = cancellation;
        }

        /// <summary>
        /// 获取转发到真实暂存流的字节数。
        /// </summary>
        public long BytesWritten { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;

        /// <inheritdoc />
        public override bool CanSeek => _inner.CanSeek;

        /// <inheritdoc />
        public override bool CanWrite => _inner.CanWrite;

        /// <inheritdoc />
        public override long Length => _inner.Length;

        /// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) =>
            _inner.FlushAsync(cancellationToken);

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);

        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            _inner.Write(buffer, offset, count);
            CancelAfterWrite(count);
        }

        /// <inheritdoc />
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            _inner.Write(buffer);
            CancelAfterWrite(buffer.Length);
        }

        /// <inheritdoc />
        public override async Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            await _inner.WriteAsync(buffer, offset, count, cancellationToken);
            CancelAfterWrite(count);
        }

        /// <inheritdoc />
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await _inner.WriteAsync(buffer, cancellationToken);
            CancelAfterWrite(buffer.Length);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
        }

        /// <summary>
        /// 记录已写字节并在首次调用时触发取消。
        /// </summary>
        /// <param name="count">刚刚写入的字节数。</param>
        private void CancelAfterWrite(int count)
        {
            BytesWritten += count;
            if (_canceled)
                return;
            _canceled = true;
            _cancellation.Cancel();
            throw new OperationCanceledException(_cancellation.Token);
        }
    }
}
