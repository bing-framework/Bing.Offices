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
using Bing.Offices.ClosedXml.Tests.AdvancedPivotResources;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证高级透视导出的完整缓存、文件提交和取消资源边界。
/// </summary>
public sealed class AdvancedPivotResourceTest
{
    /// <summary>
    /// SpreadsheetML 工作簿部件使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// 验证同步和异步流导出生成等价的高级透视结构及完整缓存。
    /// </summary>
    [Fact]
    public async Task StreamExport_SyncAndAsync_ShouldPreserveAdvancedPivotCache()
    {
        var request = CreateRequest();
        using var synchronous = new MemoryStream();
        using var asynchronous = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();

        exporter.Export(request, synchronous);
        await exporter.ExportAsync(request, asynchronous);

        Assert.True(synchronous.CanRead && synchronous.CanWrite);
        Assert.True(asynchronous.CanRead && asynchronous.CanWrite);
        AssertAdvancedStructure(synchronous.ToArray());
        AssertAdvancedStructure(asynchronous.ToArray());
        var syncParts = ReadParts(synchronous.ToArray());
        var asyncParts = ReadParts(asynchronous.ToArray());
        Assert.True(XNode.DeepEquals(syncParts.Table, asyncParts.Table));
        Assert.True(XNode.DeepEquals(syncParts.Cache, asyncParts.Cache));
        Assert.True(XNode.DeepEquals(syncParts.Records, asyncParts.Records));
    }

    /// <summary>
    /// 验证同步和异步文件提交保留高级透视数据并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_Success_ShouldCommitCompleteAdvancedPivot(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        var target = Path.Combine(directory.Path, "advanced.xlsx");
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);

        await ExportToFile(exporter, CreateRequest(), target, asynchronous);

        Assert.Empty(observer.Exceptions);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.Equal(CancellationToken.None, committer.LastToken);
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        AssertAdvancedStructure(File.ReadAllBytes(target));
    }

    /// <summary>
    /// 验证非法高级筛选在文件预检失败时保留目标并观察一次。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_InvalidAdvancedFilter_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        var target = Path.Combine(directory.Path, "advanced.xlsx");
        var original = new byte[] { 5, 8, 13, 21 };
        File.WriteAllBytes(target, original);
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
            ExportToFile(exporter, CreateRequest(invalidFilter: true), target, asynchronous));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Equal(CancellationToken.None, committer.LastToken);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证同步和异步高级透视导出不关闭调用方流。
    /// </summary>
    /// <param name="asynchronous">是否使用异步流导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StreamExport_ShouldLeaveCallerStreamOpen(bool asynchronous)
    {
        using var output = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();
        await ExportToStream(exporter, CreateRequest(), output, asynchronous);

        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        AssertAdvancedStructure(output.ToArray());
    }

    /// <summary>
    /// 验证预取消同步和异步提交保留旧目标并转发原取消令牌。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_PreCanceled_ShouldPreserveTargetAndForwardToken(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var target = SeedTarget(directory);
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancellation.Token, committer.LastToken);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(observer.Exceptions);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证来源枚举期间取消保留旧目标并清理真实暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_SourceEnumerationCanceled_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = SeedTarget(directory);
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode.Success);
        var exporter = CreateExporter(observer, committer);
        var request = CreateCanceledEnumerationRequest(cancellation);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, request, target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancellation.Token, committer.LastToken);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(observer.Exceptions);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证真实暂存写入期间取消保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledDuringTemporaryWrite_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = SeedTarget(directory);
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(AdvancedPivotCommitMode.CancelDuringWrite, cancellation);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancellation.Token, committer.LastToken);
        Assert.Equal(asynchronous ? 0 : 1, committer.SyncCalls);
        Assert.Equal(asynchronous ? 1 : 0, committer.AsyncCalls);
        Assert.True(committer.BytesWritten > 0);
        Assert.Empty(observer.Exceptions);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证完整序列化后取消仍保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledAfterSerialization_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = SeedTarget(directory);
        var observer = new AdvancedPivotRecordingObserver();
        var committer = new AdvancedPivotRecordingCommitter(
            AdvancedPivotCommitMode.CancelAfterSerialization, cancellation);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(cancellation.Token, committer.LastToken);
        Assert.True(committer.SerializedBytes > 0);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(observer.Exceptions);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 验证完整高级工作簿序列化后提交失败只观察一次并保留旧目标。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件提交入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CommitFailure_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new AdvancedPivotTestDirectory();
        var target = SeedTarget(directory);
        var observer = new AdvancedPivotRecordingObserver();
        var failure = new BingOfficesFileCommitException("模拟提交失败", provider: "ClosedXML");
        var committer = new AdvancedPivotRecordingCommitter(
            AdvancedPivotCommitMode.FailAfterSerialization, failure: failure);
        var exporter = CreateExporter(observer, committer);

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous));

        Assert.Same(failure, exception);
        Assert.Equal(BingOfficesStage.Commit, exception.Stage);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.True(committer.SerializedBytes > 0);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Equal(CancellationToken.None, committer.LastToken);
        AssertOldTargetAndNoTemporaryFile(directory, target);
    }

    /// <summary>
    /// 创建带高级轴、双度量和报表筛选的工作簿请求。
    /// </summary>
    /// <param name="invalidFilter">是否选择来源中不存在的筛选项。</param>
    /// <returns>包含完整来源记录的高级透视请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(bool invalidFilter = false) =>
        ExcelExport.Workbook(builder => builder
            .AddSheet("Data", Rows())
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 2
                },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowFields = new[]
                {
                    new ExcelPivotAxisFieldDefinition { Field = "Region", SortOrder = ExcelPivotSortOrder.Ascending }
                },
                ValueFields = new[]
                {
                    new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" },
                    new ExcelPivotValueDefinition
                    {
                        Field = "Amount", Caption = "Record count", Aggregation = ExcelPivotAggregation.Count
                    }
                },
                Filters = new[]
                {
                    new ExcelPivotFilterDefinition
                    {
                        Field = "Status", Values = new object[] { invalidFilter ? "Missing" : "Open" }
                    }
                }
            }));

    /// <summary>
    /// 创建在来源枚举第二行触发取消的高级透视请求。
    /// </summary>
    /// <param name="cancellation">取消来源枚举的令牌源。</param>
    /// <returns>引用取消枚举器的高级透视请求。</returns>
    private static ExcelWorkbookExportRequest CreateCanceledEnumerationRequest(
        CancellationTokenSource cancellation) => ExcelExport.Workbook(builder => builder
        .AddSheet("Data", CancelDuringEnumeration(cancellation))
        .PivotTable(new ExcelPivotTableDefinition
        {
            Name = "CanceledSource", SourceSheet = "Data", TargetSheet = "Summary",
            SourceRange = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 2 },
            RowFields = new[] { new ExcelPivotAxisFieldDefinition { Field = "Region" } },
            ValueFields = new[] { new ExcelPivotValueDefinition { Field = "Amount", Caption = "Total" } }
        }));

    /// <summary>
    /// 生成一行后取消来源枚举。
    /// </summary>
    /// <param name="cancellation">取消来源枚举的令牌源。</param>
    /// <returns>触发取消后停止的来源记录。</returns>
    private static IEnumerable<AdvancedPivotRow> CancelDuringEnumeration(CancellationTokenSource cancellation)
    {
        yield return new AdvancedPivotRow { Region = "North", Status = "Open", Amount = 10 };
        cancellation.Cancel();
        cancellation.Token.ThrowIfCancellationRequested();
        yield return new AdvancedPivotRow { Region = "South", Status = "Closed", Amount = 20 };
    }

    /// <summary>
    /// 生成用于缓存完整性验证的来源记录。
    /// </summary>
    /// <returns>三条含状态筛选维度的来源记录。</returns>
    private static IEnumerable<AdvancedPivotRow> Rows()
    {
        yield return new AdvancedPivotRow { Region = "North", Status = "Open", Amount = 10 };
        yield return new AdvancedPivotRow { Region = "South", Status = "Closed", Amount = 20 };
        yield return new AdvancedPivotRow { Region = "North", Status = "Open", Amount = 30 };
    }

    /// <summary>
    /// 创建接收异常观察器和文件提交器的 ClosedXML Provider。
    /// </summary>
    /// <param name="observer">记录导出异常的观察器。</param>
    /// <param name="committer">包装默认同目录提交行为的提交器。</param>
    /// <returns>配置完成的 ClosedXML Provider。</returns>
    private static ClosedXmlExcelExporter CreateExporter(AdvancedPivotRecordingObserver observer,
        AdvancedPivotRecordingCommitter committer) => new(exceptionObservers: new[] { observer },
        fileExportCommitter: committer);

    /// <summary>
    /// 根据参数调用同步或异步文件导出入口。
    /// </summary>
    /// <param name="exporter">执行导出的 ClosedXML Provider。</param>
    /// <param name="request">高级透视工作簿请求。</param>
    /// <param name="target">最终目标文件路径。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    /// <param name="cancellationToken">提交和导出的取消令牌。</param>
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
    /// 根据参数调用同步或异步流导出入口。
    /// </summary>
    /// <param name="exporter">执行导出的 ClosedXML Provider。</param>
    /// <param name="request">高级透视工作簿请求。</param>
    /// <param name="destination">由调用方拥有的目标流。</param>
    /// <param name="asynchronous">是否调用异步入口。</param>
    private static async Task ExportToStream(ClosedXmlExcelExporter exporter,
        ExcelWorkbookExportRequest request, Stream destination, bool asynchronous)
    {
        if (asynchronous)
            await exporter.ExportAsync(request, destination);
        else
            exporter.Export(request, destination);
    }

    /// <summary>
    /// 验证目标文件内容仍是旧字节且提交器没有遗留暂存文件。
    /// </summary>
    /// <param name="directory">隔离的临时输出目录。</param>
    /// <param name="target">目标文件路径。</param>
    private static void AssertOldTargetAndNoTemporaryFile(AdvancedPivotTestDirectory directory, string target)
    {
        Assert.Equal(new byte[] { 2, 3, 5, 7, 11 }, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 创建包含固定旧内容的目标文件。
    /// </summary>
    /// <param name="directory">隔离的临时输出目录。</param>
    /// <returns>目标文件绝对路径。</returns>
    private static string SeedTarget(AdvancedPivotTestDirectory directory)
    {
        var target = Path.Combine(directory.Path, "advanced.xlsx");
        File.WriteAllBytes(target, new byte[] { 2, 3, 5, 7, 11 });
        return target;
    }

    /// <summary>
    /// 验证高级透视表定义、字段、筛选、完整缓存记录和值。
    /// </summary>
    /// <param name="bytes">已保存的 XLSX 工作簿。</param>
    private static void AssertAdvancedStructure(byte[] bytes)
    {
        var parts = ReadParts(bytes);
        var table = parts.Table.Root;
        Assert.Equal("SalesByRegion", (string)table.Attribute("name"));
        Assert.Equal("C4", ((string)table.Element(SpreadsheetNs + "location").Attribute("ref"))
            .Split(':')[0]);
        Assert.Equal("1", (string)table.Element(SpreadsheetNs + "location").Attribute("rowPageCount"));
        Assert.Equal("0", (string)Assert.Single(table.Element(SpreadsheetNs + "rowFields").Elements())
            .Attribute("x"));
        Assert.Equal(new[] { "2", "2" }, table.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField").Select(field => (string)field.Attribute("fld")));
        Assert.Equal(new[] { "sum", "count" }, table.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField").Select(field => (string)field.Attribute("subtotal") ?? "sum"));
        var pageField = Assert.Single(table.Element(SpreadsheetNs + "pageFields").Elements());
        Assert.Equal("1", (string)pageField.Attribute("fld"));

        var cacheFields = parts.Cache.Root.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField").ToArray();
        Assert.Equal(new[] { "Region", "Status", "Amount" },
            cacheFields.Select(field => (string)field.Attribute("name")));
        var source = parts.Cache.Root.Element(SpreadsheetNs + "cacheSource")
            .Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Data", (string)source.Attribute("sheet"));
        Assert.Equal("A1:C4", (string)source.Attribute("ref"));

        var rows = parts.Records.Root.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, row => Assert.Equal(3, row.Elements().Count()));
        Assert.Equal(new[]
        {
            "North|Open|10", "South|Closed|20", "North|Open|30"
        }, rows.Select(row => string.Join("|", Enumerable.Range(0, 3)
            .Select(index => CacheValue(cacheFields[index], row.Elements().ElementAt(index))))));
        var statusField = table.Element(SpreadsheetNs + "pivotFields").Elements().ElementAt(1);
        var selectedItem = statusField.Element(SpreadsheetNs + "items").Elements()
            .ElementAt(int.Parse((string)pageField.Attribute("item")));
        var selectedIndex = int.Parse((string)selectedItem.Attribute("x"));
        Assert.Equal("Open", (string)cacheFields[1].Element(SpreadsheetNs + "sharedItems")
            .Elements().ElementAt(selectedIndex).Attribute("v"));

        var cachePart = AssertPackageRelationship(bytes, "xl/workbook.xml", "pivotCacheDefinition");
        AssertPackageRelationship(bytes, cachePart, "pivotCacheRecords");
        AssertPackageRelationship(bytes, ResolveWorksheetPart(bytes, "Summary"), "pivotTable");
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        Assert.Equal(new[] { "Data", "Summary" }, workbook.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "Region", "Status", "Amount", "North", "Open", "10", "South", "Closed", "20",
            "North", "Open", "30" }, workbook.Worksheet("Data").Range("A1:C4").Cells()
            .Select(cell => cell.GetString()));
        var pivot = Assert.Single(workbook.Worksheet("Summary").PivotTables);
        Assert.Equal(new[] { "Total", "Record count" }, pivot.Values.Select(value => value.CustomName));
        Assert.Equal(new[] { XLPivotSummary.Sum, XLPivotSummary.Count },
            pivot.Values.Select(value => value.SummaryFormula));
        Assert.Equal(XLPivotSortType.Ascending, pivot.RowLabels.Get("Region").SortType);
    }

    /// <summary>
    /// 解析完整工作簿中的透视表、缓存定义和缓存记录部件。
    /// </summary>
    /// <param name="bytes">完整 XLSX 字节。</param>
    /// <returns>三个透视 OOXML 文档。</returns>
    private static (XDocument Table, XDocument Cache, XDocument Records) ReadParts(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var table = ReadSingleXml(archive, path => path.StartsWith("xl/pivotTables/", StringComparison.Ordinal)
            && path.EndsWith(".xml", StringComparison.Ordinal));
        var cache = ReadSingleXml(archive, path => path.Split('/').Last()
            .StartsWith("pivotCacheDefinition", StringComparison.Ordinal)
            && path.EndsWith(".xml", StringComparison.Ordinal));
        var records = ReadSingleXml(archive, path => path.Split('/').Last()
            .StartsWith("pivotCacheRecords", StringComparison.Ordinal)
            && path.EndsWith(".xml", StringComparison.Ordinal));
        return (table, cache, records);
    }

    /// <summary>
    /// 读取压缩包中满足路径筛选条件的唯一 XML 部件。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="predicate">目标部件路径条件。</param>
    /// <returns>部件路径与解析后的文档。</returns>
    private static XDocument ReadSingleXml(ZipArchive archive, Func<string, bool> predicate)
    {
        var entry = Assert.Single(archive.Entries.Where(item => predicate(item.FullName)
            && (item.FullName.EndsWith(".xml", StringComparison.Ordinal)
                || item.FullName.EndsWith(".rels", StringComparison.Ordinal))));
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    /// <summary>
    /// 将缓存记录的字段编码解析为业务值。
    /// </summary>
    /// <param name="cacheField">对应来源字段缓存定义。</param>
    /// <param name="recordCell">缓存记录的字段单元格。</param>
    /// <returns>缓存记录中的文本、数字或空值标记。</returns>
    private static string CacheValue(XElement cacheField, XElement recordCell)
    {
        if (recordCell.Name == SpreadsheetNs + "m")
            return string.Empty;
        if (recordCell.Name == SpreadsheetNs + "x")
        {
            var index = int.Parse((string)recordCell.Attribute("v"));
            return (string)cacheField.Element(SpreadsheetNs + "sharedItems")
                .Elements().ElementAt(index).Attribute("v");
        }
        return (string)recordCell.Attribute("v") ?? recordCell.Value;
    }

    /// <summary>
    /// 验证包关系中包含指定关系类型。
    /// </summary>
    /// <param name="bytes">XLSX 工作簿字节。</param>
    /// <param name="sourcePart">关系源部件。</param>
    /// <param name="relationshipSuffix">关系类型末尾。</param>
    /// <returns>关系指向的实际部件路径。</returns>
    private static string AssertPackageRelationship(byte[] bytes, string sourcePart,
        string relationshipSuffix)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var slash = sourcePart.LastIndexOf('/');
        var relationshipPart = sourcePart.Substring(0, slash + 1) + "_rels/"
            + sourcePart.Substring(slash + 1) + ".rels";
        XNamespace relationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var relationships = ReadSingleXml(archive, path => path == relationshipPart).Root
            .Elements(relationshipNs + "Relationship").ToArray();
        var relationship = Assert.Single(relationships.Where(item =>
            ((string)item.Attribute("Type"))?.EndsWith("/" + relationshipSuffix,
                StringComparison.Ordinal) == true));
        var sourceDirectory = sourcePart.Substring(0, slash + 1);
        var target = ResolvePartPath(sourceDirectory, (string)relationship.Attribute("Target"));
        Assert.NotNull(archive.GetEntry(target));
        return target;
    }

    /// <summary>
    /// 按工作簿关系解析指定工作表的实际部件路径。
    /// </summary>
    /// <param name="bytes">XLSX 工作簿字节。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <returns>工作表在压缩包中的实际部件路径。</returns>
    private static string ResolveWorksheetPart(byte[] bytes, string sheetName)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        XNamespace spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var workbook = ReadSingleXml(archive, path => path == "xl/workbook.xml");
        var relationshipId = (string)workbook.Root.Element(spreadsheetNs + "sheets")
            .Elements(spreadsheetNs + "sheet")
            .Single(sheet => (string)sheet.Attribute("name") == sheetName)
            .Attribute(officeRelationshipNs + "id");
        var relationships = ReadSingleXml(archive, path => path == "xl/_rels/workbook.xml.rels");
        var target = (string)relationships.Root.Elements(packageRelationshipNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == relationshipId)
            .Attribute("Target");
        return ResolvePartPath("xl/", target);
    }

    /// <summary>
    /// 解析包关系目标中的当前目录与父目录段。
    /// </summary>
    /// <param name="sourceDirectory">关系源部件所在目录。</param>
    /// <param name="target">关系文件内的目标路径。</param>
    /// <returns>相对于 XLSX 包根目录的标准路径。</returns>
    private static string ResolvePartPath(string sourceDirectory, string target)
    {
        var segments = (target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : sourceDirectory + target)
            .Split('/');
        var resolved = new List<string>();
        foreach (var segment in segments)
        {
            if (segment == "..")
                resolved.RemoveAt(resolved.Count - 1);
            else if (segment != "." && segment.Length > 0)
                resolved.Add(segment);
        }
        return string.Join("/", resolved);
    }

}
