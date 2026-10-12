using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using ClosedXML.Excel;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 ClosedXML 透视表工作簿的流、文件提交和取消资源边界。
/// </summary>
public sealed class PivotTableResourcesTest
{
    /// <summary>
    /// 验证同步和异步导出生成相同的透视表、缓存定义和完整记录。
    /// </summary>
    [Fact]
    public async Task StreamExport_SyncAndAsync_ShouldPreserveCompletePivotCache()
    {
        var request = CreateRequest();
        using var synchronous = new MemoryStream();
        using var asynchronous = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();

        exporter.Export(request, synchronous);
        await exporter.ExportAsync(request, asynchronous);

        Assert.True(synchronous.CanRead);
        Assert.True(synchronous.CanWrite);
        Assert.True(asynchronous.CanRead);
        Assert.True(asynchronous.CanWrite);
        var syncStructure = Inspect(synchronous.ToArray());
        var asyncStructure = Inspect(asynchronous.ToArray());
        AssertCompletePivotStructure(syncStructure);
        AssertCompletePivotStructure(asyncStructure);
        Assert.True(XNode.DeepEquals(syncStructure.PivotTable, asyncStructure.PivotTable));
        Assert.True(XNode.DeepEquals(syncStructure.CacheDefinition, asyncStructure.CacheDefinition));
        Assert.True(XNode.DeepEquals(syncStructure.CacheRecords, asyncStructure.CacheRecords));
    }

    /// <summary>
    /// 验证同步和异步文件提交均保留完整透视缓存并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_Success_ShouldCommitCompletePivotWorkbook(bool asynchronous)
    {
        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        if (asynchronous)
            await exporter.ExportToFileAsync(CreateRequest(), target);
        else
            exporter.ExportToFile(CreateRequest(), target);

        Assert.Empty(observer.Exceptions);
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
        var bytes = File.ReadAllBytes(target);
        var structure = Inspect(bytes);
        AssertCompletePivotStructure(structure);
        using var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false));
        Assert.Equal(new[] { "Data", "Summary" }, workbook.Worksheets.Select(sheet => sheet.Name));
        Assert.Equal(new[] { "Region", "Amount", "North", "10", "South", "20", "North", "30" },
            workbook.Worksheet("Data").Range("A1:B4").Cells().Select(cell => cell.GetString()));
    }

    /// <summary>
    /// 验证预检错误通过同步和异步文件入口分类、观察一次并保留旧文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_InvalidPivot_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 8, 3, 5, 7, 9 };
        File.WriteAllBytes(target, original);
        var observer = new RecordingObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        var request = CreateRequest(valueField: "MissingValue");

        var exception = await ExportToFile(exporter, request, target, asynchronous);

        Assert.NotNull(exception);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证调用方流在同步或异步透视表导出后仍由调用方持有。
    /// </summary>
    /// <param name="asynchronous">是否调用异步流入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StreamExport_ShouldLeaveCallerStreamOpen(bool asynchronous)
    {
        using var output = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter();
        if (asynchronous)
            await exporter.ExportAsync(CreateRequest(), output);
        else
            exporter.Export(CreateRequest(), output);

        Assert.True(output.CanRead);
        Assert.True(output.CanWrite);
        AssertCompletePivotStructure(Inspect(output.ToArray()));
    }

    /// <summary>
    /// 验证完整透视表已写入暂存文件后取消仍保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledAfterSerialization_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 2, 4, 6, 8 };
        File.WriteAllBytes(target, original);
        var observer = new RecordingObserver();
        var committer = new CancelAfterWriteCommitter(cancellation);
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer },
            fileExportCommitter: committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(committer.SerializedBytes > 0);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证暂存文件实际写入期间取消会保留旧目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CanceledDuringWrite_ShouldPreserveTarget(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var cancellation = new CancellationTokenSource();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 3, 1, 4, 1, 5 };
        File.WriteAllBytes(target, original);
        var observer = new RecordingObserver();
        var committer = new CancelDuringWriteCommitter(cancellation);
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer },
            fileExportCommitter: committer);

        var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous, cancellation.Token));

        Assert.Equal(cancellation.Token, exception.CancellationToken);
        Assert.True(committer.BytesWritten > 0);
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.Empty(observer.Exceptions);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 验证完整透视表序列化后提交失败只观察一次并保留旧目标。
    /// </summary>
    /// <param name="asynchronous">是否调用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FileExport_CommitFailure_ShouldPreserveTargetAndObserveOnce(bool asynchronous)
    {
        using var directory = new TestDirectory();
        var target = Path.Combine(directory.Path, "pivot.xlsx");
        var original = new byte[] { 1, 4, 1, 4, 2 };
        File.WriteAllBytes(target, original);
        var observer = new RecordingObserver();
        var failure = new BingOfficesFileCommitException("提交失败", provider: "ClosedXML");
        var committer = new FailAfterWriteCommitter(failure);
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer },
            fileExportCommitter: committer);

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            ExportToFile(exporter, CreateRequest(), target, asynchronous));

        Assert.Same(failure, exception);
        Assert.Equal(BingOfficesStage.Commit, exception.Stage);
        Assert.Same(exception, Assert.Single(observer.Exceptions));
        Assert.Equal(1, committer.SyncCalls + committer.AsyncCalls);
        Assert.True(committer.SerializedBytes > 0);
        Assert.Equal(original, File.ReadAllBytes(target));
        Assert.Equal(new[] { target }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 创建含三条明细和一个区域求和透视表的请求。
    /// </summary>
    /// <param name="valueField">作为求和字段的表头。</param>
    /// <returns>已构建的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(string valueField = "Amount") =>
        ExcelExport.Workbook(builder => builder
            .AddSheet("Data", new[]
            {
                new PivotRow { Region = "North", Amount = 10 },
                new PivotRow { Region = "South", Amount = 20 },
                new PivotRow { Region = "North", Amount = 30 }
            })
            .PivotTable(new ExcelPivotTableDefinition
            {
                Name = "SalesByRegion",
                SourceSheet = "Data",
                SourceRange = new ExcelRangeDefinition
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 3,
                    EndColumn = 1
                },
                TargetSheet = "Summary",
                TargetRow = 1,
                TargetColumn = 2,
                RowField = "Region",
                ValueField = valueField
            }));

    /// <summary>
    /// 按指定入口导出文件并返回配置错误。
    /// </summary>
    /// <param name="exporter">执行导出的 ClosedXML Provider。</param>
    /// <param name="request">透视表工作簿请求。</param>
    /// <param name="target">目标文件路径。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="cancellationToken">导出取消令牌。</param>
    /// <returns>预检配置异常；导出成功时返回 null。</returns>
    private static async Task<BingOfficesConfigurationException> ExportToFile(
        ClosedXmlExcelExporter exporter, ExcelWorkbookExportRequest request, string target,
        bool asynchronous, CancellationToken cancellationToken = default)
    {
        if (asynchronous)
        {
            try
            {
                await exporter.ExportToFileAsync(request, target, cancellationToken);
            }
            catch (BingOfficesConfigurationException exception)
            {
                return exception;
            }
            return null;
        }

        try
        {
            exporter.ExportToFile(request, target, cancellationToken);
        }
        catch (BingOfficesConfigurationException exception)
        {
            return exception;
        }
        return null;
    }

    /// <summary>
    /// 从 XLSX 包中读取透视表定义和完整缓存部件。
    /// </summary>
    /// <param name="bytes">工作簿字节。</param>
    /// <returns>解析后的透视表结构快照。</returns>
    private static PivotStructure Inspect(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var pivotPath = archive.Entries.Single(entry => entry.FullName.Contains("pivotTables/")
            && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).FullName;
        var cachePath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheDefinition1.xml", StringComparison.Ordinal)).FullName;
        var recordsPath = archive.Entries.Single(entry => entry.FullName.EndsWith(
            "pivotCacheRecords1.xml", StringComparison.Ordinal)).FullName;
        AssertRelationship(archive, "xl/workbook.xml", "pivotCacheDefinition", cachePath);
        var worksheetPath = ResolveWorksheetPath(archive, "Summary");
        AssertRelationship(archive, worksheetPath, "pivotTable", pivotPath);
        AssertRelationship(archive, cachePath, "pivotCacheRecords", recordsPath);
        return new PivotStructure(ReadXml(archive, pivotPath), ReadXml(archive, cachePath),
            ReadXml(archive, recordsPath));
    }

    /// <summary>
    /// 验证指定 OOXML 部件通过关系文件指向预期目标部件。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="sourcePart">关系源部件路径。</param>
    /// <param name="relationshipSuffix">关系类型末尾名称。</param>
    /// <param name="expectedTarget">预期目标部件路径。</param>
    private static void AssertRelationship(ZipArchive archive, string sourcePart,
        string relationshipSuffix, string expectedTarget)
    {
        XNamespace relationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var directory = sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1);
        var file = sourcePart.Substring(sourcePart.LastIndexOf('/') + 1);
        var relationshipPath = directory + "_rels/" + file + ".rels";
        var relationships = ReadXml(archive, relationshipPath).Root
            .Elements(relationshipNs + "Relationship")
            .Where(item => ((string)item.Attribute("Type"))?.EndsWith("/" + relationshipSuffix,
                StringComparison.Ordinal) == true)
            .ToArray();
        var relationship = Assert.Single(relationships);
        Assert.Equal(expectedTarget, ResolvePartPath(directory, (string)relationship.Attribute("Target")));
    }

    /// <summary>
    /// 根据工作表名称解析其 OOXML 工作表部件路径。
    /// </summary>
    /// <param name="archive">XLSX 压缩包。</param>
    /// <param name="sheetName">工作表名称。</param>
    /// <returns>相对于压缩包根目录的工作表部件路径。</returns>
    private static string ResolveWorksheetPath(ZipArchive archive, string sheetName)
    {
        XNamespace spreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace officeRelationshipNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        var workbook = ReadXml(archive, "xl/workbook.xml");
        var relationshipId = (string)workbook.Root.Element(spreadsheetNs + "sheets")
            .Elements(spreadsheetNs + "sheet")
            .Single(sheet => (string)sheet.Attribute("name") == sheetName)
            .Attribute(officeRelationshipNs + "id");
        var relationships = ReadXml(archive, "xl/_rels/workbook.xml.rels");
        var target = (string)relationships.Root.Elements(packageRelationshipNs + "Relationship")
            .Single(item => (string)item.Attribute("Id") == relationshipId).Attribute("Target");
        return ResolvePartPath("xl/", target);
    }

    /// <summary>
    /// 将 OOXML 关系目标解析为压缩包内的规范部件路径。
    /// </summary>
    /// <param name="sourceDirectory">关系源部件所在目录。</param>
    /// <param name="target">关系文件中的目标路径。</param>
    /// <returns>相对于压缩包根目录的部件路径。</returns>
    private static string ResolvePartPath(string sourceDirectory, string target)
    {
        var segments = (target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : sourceDirectory + target)
            .Split('/');
        var resolved = new System.Collections.Generic.List<string>();
        foreach (var segment in segments)
        {
            if (segment == "..")
                resolved.RemoveAt(resolved.Count - 1);
            else if (segment != "." && segment.Length > 0)
                resolved.Add(segment);
        }
        return string.Join("/", resolved);
    }

    /// <summary>
    /// 验证透视表位置、字段、来源范围及全部缓存记录和值。
    /// </summary>
    /// <param name="structure">待验证的透视表结构。</param>
    private static void AssertCompletePivotStructure(PivotStructure structure)
    {
        var table = structure.PivotTable.Root;
        Assert.Equal("SalesByRegion", (string)table.Attribute("name"));
        var location = (string)table.Element(SpreadsheetNs + "location").Attribute("ref");
        Assert.Equal("C2", location.Split(':')[0]);
        var rowField = Assert.Single(table.Element(SpreadsheetNs + "rowFields").Elements());
        Assert.Equal("0", (string)rowField.Attribute("x"));
        var dataField = table.Element(SpreadsheetNs + "dataFields")
            .Elements(SpreadsheetNs + "dataField").Single();
        Assert.Equal("1", (string)dataField.Attribute("fld"));
        Assert.True(dataField.Attribute("subtotal") == null
            || (string)dataField.Attribute("subtotal") == "sum");

        var cache = structure.CacheDefinition.Root;
        var source = cache.Element(SpreadsheetNs + "cacheSource")
            .Element(SpreadsheetNs + "worksheetSource");
        Assert.Equal("Data", (string)source.Attribute("sheet"));
        Assert.Equal("A1:B4", (string)source.Attribute("ref"));
        Assert.Equal(new[] { "Region", "Amount" }, cache.Element(SpreadsheetNs + "cacheFields")
            .Elements(SpreadsheetNs + "cacheField")
            .Select(field => (string)field.Attribute("name")));

        var records = structure.CacheRecords.Root;
        Assert.True(records.Attribute("count") == null || (string)records.Attribute("count") == "3");
        var rows = records.Elements(SpreadsheetNs + "r").ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, row => Assert.Equal(2, row.Elements().Count()));
        var sharedRegions = cache.Element(SpreadsheetNs + "cacheFields").Elements().First()
            .Element(SpreadsheetNs + "sharedItems").Elements()
            .Select(item => (string)item.Attribute("v")).ToArray();
        Assert.Equal(new[] { "North", "South", "North" }, rows.Select(row => row.Elements().First())
            .Select(item => item.Name.LocalName == "x"
                ? sharedRegions[int.Parse((string)item.Attribute("v"))]
                : (string)item.Attribute("v")));
        Assert.Equal(new[] { "10", "20", "30" }, rows
            .Select(row => (string)row.Elements().Last().Attribute("v")));
        var cachedValues = cache.Descendants(SpreadsheetNs + "sharedItems")
            .SelectMany(items => items.Elements())
            .SelectMany(item => item.Attributes("v").Select(attribute => attribute.Value)
                .Concat(string.IsNullOrEmpty(item.Value) ? Array.Empty<string>() : new[] { item.Value }))
            .Concat(records.Descendants().SelectMany(item => item.Attributes("v")
                .Select(attribute => attribute.Value)))
            .ToArray();
        Assert.Contains("North", cachedValues);
        Assert.Contains("South", cachedValues);
        Assert.Contains("10", cachedValues);
        Assert.Contains("20", cachedValues);
        Assert.Contains("30", cachedValues);
    }

    /// <summary>
    /// 从 XLSX 压缩包读取指定 XML 部件。
    /// </summary>
    /// <param name="archive">工作簿压缩包。</param>
    /// <param name="name">部件路径。</param>
    /// <returns>已加载的 XML 文档。</returns>
    private static XDocument ReadXml(ZipArchive archive, string name)
    {
        var entry = archive.GetEntry(name);
        Assert.NotNull(entry);
        using var entryStream = entry.Open();
        return XDocument.Load(entryStream);
    }

    /// <summary>
    /// SpreadsheetML 工作簿部件使用的 XML 命名空间。
    /// </summary>
    private static readonly XNamespace SpreadsheetNs =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private sealed class PivotStructure
    {
        /// <summary>
        /// 初始化一个 <see cref="PivotStructure" /> 类型的实例。
        /// </summary>
        public PivotStructure(XDocument pivotTable, XDocument cacheDefinition, XDocument cacheRecords)
        {
            PivotTable = pivotTable;
            CacheDefinition = cacheDefinition;
            CacheRecords = cacheRecords;
        }

        /// <summary>获取透视表定义 XML。</summary>
        public XDocument PivotTable { get; }

        /// <summary>获取透视缓存定义 XML。</summary>
        public XDocument CacheDefinition { get; }

        /// <summary>获取完整透视缓存记录 XML。</summary>
        public XDocument CacheRecords { get; }
    }

    private sealed class PivotRow
    {
        /// <summary>
        /// 初始化一个 <see cref="PivotRow" /> 类型的实例。
        /// </summary>
        public PivotRow()
        {
        }

        /// <summary>获取或设置区域。</summary>
        public string Region { get; set; }

        /// <summary>获取或设置金额。</summary>
        public decimal Amount { get; set; }
    }

    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory()
        {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                "bing-pivot-resources-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        /// <summary>获取临时目录绝对路径。</summary>
        public string Path { get; }

        /// <inheritdoc />
        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
    }

    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>获取已观察的 Bing.Offices 异常。</summary>
        public System.Collections.Generic.List<BingOfficesException> Exceptions { get; } = new();

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }

    private sealed class CancelDuringWriteCommitter : IFileExportCommitter
    {
        /// <summary>执行真实暂存和清理的默认提交器。</summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>首次写入暂存文件后触发的取消源。</summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>最近一次提交中拦截暂存写入的流。</summary>
        private CancelingWriteStream _stream;

        /// <summary>
        /// 初始化一个 <see cref="CancelDuringWriteCommitter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">写入期间触发的取消源。</param>
        public CancelDuringWriteCommitter(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <summary>获取已转发到暂存文件的字节数。</summary>
        public long BytesWritten => _stream?.BytesWritten ?? 0;

        /// <summary>获取同步提交调用次数。</summary>
        public int SyncCalls { get; private set; }

        /// <summary>获取异步提交调用次数。</summary>
        public int AsyncCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            SyncCalls++;
            _inner.Commit(path, destination =>
            {
                _stream = new CancelingWriteStream(destination, _cancellation);
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
                _stream = new CancelingWriteStream(destination, _cancellation);
                await writeAsync(_stream, token).ConfigureAwait(false);
            }, cancellationToken, format);
        }
    }

    private sealed class CancelingWriteStream : Stream
    {
        /// <summary>由文件提交器拥有的暂存文件流。</summary>
        private readonly Stream _inner;

        /// <summary>首次转发写入后触发的取消源。</summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>是否已触发首次写入取消。</summary>
        private bool _canceled;

        /// <summary>
        /// 初始化一个 <see cref="CancelingWriteStream" /> 类型的实例。
        /// </summary>
        /// <param name="inner">文件提交器拥有的暂存流。</param>
        /// <param name="cancellation">写入后触发的取消源。</param>
        public CancelingWriteStream(Stream inner, CancellationTokenSource cancellation)
        {
            _inner = inner;
            _cancellation = cancellation;
        }

        /// <summary>获取已转发到暂存流的字节数。</summary>
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
            await _inner.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            CancelAfterWrite(count);
        }

        /// <inheritdoc />
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
            CancelAfterWrite(buffer.Length);
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            // 提交器负责释放真实暂存流，包装流不取得其所有权。
        }

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

    private sealed class CancelAfterWriteCommitter : IFileExportCommitter
    {
        /// <summary>执行真实暂存和原子替换的默认提交器。</summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>写入完整暂存内容后触发的取消源。</summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterWriteCommitter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">写完暂存文件后取消的令牌源。</param>
        public CancelAfterWriteCommitter(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <summary>获取同步序列化的完整字节数。</summary>
        public long SerializedBytes { get; private set; }

        /// <summary>获取同步提交调用次数。</summary>
        public int SyncCalls { get; private set; }

        /// <summary>获取异步提交调用次数。</summary>
        public int AsyncCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            SyncCalls++;
            _inner.Commit(path, destination =>
            {
                write(destination);
                SerializedBytes = destination.Length;
                _cancellation.Cancel();
            }, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            return _inner.CommitAsync(path, async (destination, token) =>
            {
                await writeAsync(destination, token).ConfigureAwait(false);
                SerializedBytes = destination.Length;
                _cancellation.Cancel();
            }, cancellationToken, format);
        }
    }

    private sealed class FailAfterWriteCommitter : IFileExportCommitter
    {
        /// <summary>执行真实暂存和清理的默认提交器。</summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>写出完整工作簿后作为提交失败返回的异常。</summary>
        private readonly BingOfficesFileCommitException _failure;

        /// <summary>
        /// 初始化一个 <see cref="FailAfterWriteCommitter" /> 类型的实例。
        /// </summary>
        /// <param name="failure">完整序列化后抛出的提交异常。</param>
        public FailAfterWriteCommitter(BingOfficesFileCommitException failure) => _failure = failure;

        /// <summary>获取完整序列化的暂存字节数。</summary>
        public long SerializedBytes { get; private set; }

        /// <summary>获取同步提交调用次数。</summary>
        public int SyncCalls { get; private set; }

        /// <summary>获取异步提交调用次数。</summary>
        public int AsyncCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            SyncCalls++;
            _inner.Commit(path, destination =>
            {
                write(destination);
                SerializedBytes = destination.Length;
                throw _failure;
            }, cancellationToken, format);
        }

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            return _inner.CommitAsync(path, async (destination, token) =>
            {
                await writeAsync(destination, token).ConfigureAwait(false);
                SerializedBytes = destination.Length;
                throw _failure;
            }, cancellationToken, format);
        }
    }
}
