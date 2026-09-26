using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.SpreadCheetah.Tests;

/// <summary>
/// SpreadCheetah 流式导出器职责级测试。
/// </summary>
public sealed class SpreadCheetahStreamingExcelExporterTest
{
    /// <summary>
    /// 验证同一区域的表格和筛选合并输出。
    /// </summary>
    /// <param name="style">待验证的表格样式名。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [InlineData("TableStyleMedium2", true)]
    [InlineData("Medium2", true)]
    [InlineData("TableStyleMedium2", false)]
    [InlineData("Medium2", false)]
    public async Task ExportBatchesAsync_ShouldMergeSameRangeTableAndFilter(string style, bool async)
    {
        var range = new ExcelRangeDefinition { EndRow = 1, EndColumn = 1 };
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows", new[] { new Row { Id = 1, Name = "中文" } },
            sheet => sheet.Table(new ExcelTableDefinition { Name = "RowsTable", Range = range, StyleName = style })
                .AutoFilter(new ExcelAutoFilterDefinition { Range = range })));
        using var output = new MemoryStream();
        var exporter = new SpreadCheetahStreamingExcelExporter();
        if (async) await exporter.ExportBatchesAsync(request, output);
        else exporter.ExportBatches(request, output);
        using var workbook = Open(output);
        var sheet = (XSSFSheet)workbook.GetSheetAt(0);
        Assert.Equal("中文", sheet.GetRow(1).GetCell(1).StringCellValue);
        var table = Assert.Single(sheet.GetTables());
        Assert.Equal("TableStyleMedium2", table.StyleName);
        Assert.Equal("A1:B2", table.CellReferences.FormatAsString());
    }

    /// <summary>
    /// 提供不受支持的报表请求预检用例。
    /// </summary>
    /// <returns>报表场景与同步或异步入口的组合用例。</returns>
    public static IEnumerable<object[]> ReportPreflightCases()
    {
        foreach (var scenario in new[] { "style", "table-count", "filters", "table-layout", "duplicate-name", "out-of-range" })
        foreach (var async in new[] { false, true })
            yield return new object[] { scenario, async };
    }

    /// <summary>
    /// 验证表格区域不匹配失败且不重复枚举数据。
    /// </summary>
    /// <param name="count">测试数据行数。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(0, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(2, true)]
    public async Task ExportBatches_TableRangeMismatch_ShouldFailWithoutSecondEnumeration(int count, bool async)
    {
        var data = new TrackingEnumerable<Row>(Enumerable.Range(1, count).Select(id => new Row { Id = id, Name = "值" }));
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows", data, sheet => sheet.Table(new ExcelTableDefinition
        {
            Name = "RowsTable", Range = new ExcelRangeDefinition { EndRow = 1, EndColumn = 1 }
        })));
        using var output = new MemoryStream();
        var exporter = new SpreadCheetahStreamingExcelExporter();
        var error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(async () =>
        {
            if (async) await exporter.ExportBatchesAsync(request, output);
            else exporter.ExportBatches(request, output);
        });
        Assert.Equal(BingOfficesStage.Write, error.Stage);
        Assert.Equal(1, data.EnumeratorCount);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证后续工作表的无效请求在任何输出前被拒绝。
    /// </summary>
    /// <param name="scenario">触发预检失败的报表场景标识。</param>
    /// <param name="async">是否调用异步导出入口。</param>
    [Theory]
    [MemberData(nameof(ReportPreflightCases))]
    public async Task ExportBatchesAsync_ShouldPreflightLaterSheetBeforeAnyOutput(string scenario, bool async)
    {
        var data = new TrackingEnumerable<Row>(new[] { new Row { Id = 1, Name = "中文" } });
        var range = new ExcelRangeDefinition { EndRow = scenario == "out-of-range" ? 1048576 : 1, EndColumn = 1 };
        var request = ExcelExport.Workbook(book => book.AddSheet("First", data, sheet =>
        {
            if (scenario == "duplicate-name")
                sheet.Table(new ExcelTableDefinition { Name = "TableOne", Range = range });
        }).AddSheet("Invalid", data, sheet =>
        {
            if (scenario == "filters")
                sheet.AutoFilter(new ExcelAutoFilterDefinition { Range = range })
                    .AutoFilter(new ExcelAutoFilterDefinition { Range = new ExcelRangeDefinition { EndRow = 2, EndColumn = 1 } });
            else
            {
                sheet.Table(new ExcelTableDefinition
                {
                    Name = "TableOne", Range = range,
                    ShowTotals = scenario == "table-layout", StyleName = scenario == "style" ? "BadStyle" : null
                });
                if (scenario == "table-count")
                    sheet.Table(new ExcelTableDefinition { Name = "TableTwo", Range = range });
            }
        }));
        using var output = new MemoryStream();
        var exporter = new SpreadCheetahStreamingExcelExporter();
        var error = await Record.ExceptionAsync(async () =>
        {
            if (async) await exporter.ExportBatchesAsync(request, output);
            else exporter.ExportBatches(request, output);
        });
        Assert.Equal(BingOfficesStage.Preflight, Assert.IsAssignableFrom<BingOfficesException>(error).Stage);
        Assert.Equal(0, data.EnumeratorCount);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证能力描述限定为支持异步写入的新 XLSX 导出。
    /// </summary>
    [Fact]
    public void Capabilities_ShouldDescribeAsyncWriteOnlyXlsxBoundary()
    {
        var exporter = new SpreadCheetahStreamingExcelExporter();

        Assert.Empty(exporter.ReadFormats);
        Assert.Equal(new[] { ExcelFormat.Xlsx }, exporter.WriteFormats);
        Assert.True(exporter.SupportsTrueAsyncIo);
        Assert.False(exporter.SupportsCompleteWorkbookExport);
        Assert.True(exporter.Supports(ExcelProviderCapabilities.List
            | ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Async
            | ExcelProviderCapabilities.Xlsx));
    }

    /// <summary>
    /// 验证多个工作表按请求顺序写入且仅枚举一次。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldWriteMultipleSheetsInOrder()
    {
        var first = new TrackingEnumerable<Row>(new[]
        {
            new Row { Id = 1, Name = "Alpha" },
            new Row { Id = 2, Name = "Beta" }
        });
        var second = new TrackingEnumerable<Row>(new[] { new Row { Id = 3, Name = "Gamma" } });
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("First", first);
            workbook.AddSheet("Second", second);
        });
        var destination = new MemoryStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination,
            new ExcelStreamingExportOptions { BatchSize = 1 });

        Assert.Equal(1, first.EnumeratorCount);
        Assert.Equal(1, second.EnumeratorCount);
        using var workbook = Open(destination);
        Assert.Equal(2, workbook.NumberOfSheets);
        Assert.Equal(1d, workbook.GetSheet("First").GetRow(1).GetCell(0).NumericCellValue);
        Assert.Equal("Alpha", workbook.GetSheet("First").GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("Gamma", workbook.GetSheet("Second").GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 验证导出使用异步写入且保持不可定位目标流打开。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldUseAsyncIoAndLeaveNonSeekableDestinationOpen()
    {
        var request = CreateRequest(2);
        var destination = new TrackingNonSeekableStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination);

        Assert.True(destination.AsyncWriteCount > 0);
        Assert.True(destination.SyncWriteCount > 0);
        Assert.False(destination.IsDisposed);
        using var workbook = Open(destination.Inner);
        Assert.Equal("Name 2", workbook.GetSheetAt(0).GetRow(2).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 验证取消导出后调用方目标流保持打开。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WhenCanceled_ShouldLeaveCallerStreamOpen()
    {
        var tokenSource = new CancellationTokenSource();
        tokenSource.Cancel();
        var destination = new TrackingMemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(
                CreateRequest(1), destination, cancellationToken: tokenSource.Token));

        Assert.False(destination.IsDisposed);
    }

    /// <summary>
    /// 验证文件导出提交成功且取消不损坏已有目标。
    /// </summary>
    [Fact]
    public async Task ExportBatchesToFileAsync_ShouldCommitAtomicallyAndPreserveOldTargetOnCancellation()
    {
        var directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "result.xlsx");
            var exporter = new SpreadCheetahStreamingExcelExporter();
            await exporter.ExportBatchesToFileAsync(CreateRequest(1), path);
            using (var workbook = new XSSFWorkbook(path))
                Assert.Equal("Name 1", workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue);

            var sentinel = await File.ReadAllBytesAsync(path);
            var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => exporter.ExportBatchesToFileAsync(
                CreateRequest(2), path, cancellationToken: cancellation.Token));
            Assert.Equal(sentinel, await File.ReadAllBytesAsync(path));
            Assert.Empty(Directory.EnumerateFiles(directory).Where(file => file != path));
        }
        finally
        {
            Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证同步导出成功且不支持的模板请求被拒绝。
    /// </summary>
    [Fact]
    public async Task ExportBatches_AndAdvancedRequests_ShouldRespectSyncAndUnsupportedBoundaries()
    {
        var exporter = new SpreadCheetahStreamingExcelExporter();
        using var syncDestination = new MemoryStream();
        exporter.ExportBatches(CreateRequest(1), syncDestination);
        using (var workbook = Open(syncDestination))
            Assert.Equal("Name 1", workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue);

        var template = new MemoryStream(new byte[] { 1 });
        var request = ExcelExport.Workbook(workbook => workbook.UseTemplate(template, true)
            .AddSheet("Rows", new[] { new Row { Id = 1, Name = "A" } }));
        var unsupported = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            exporter.ExportBatchesAsync(request, new MemoryStream()));
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, unsupported.Code);
        Assert.Equal("SpreadCheetah", unsupported.Provider);
        Assert.Equal(BingOfficesStage.Preflight, unsupported.Stage);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证动态列和原生类型单元格正确写入。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldWriteDynamicColumnsAndTypedCells()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new DynamicRow { Id = 7, Name = "动态" } }, sheet => sheet
                .DynamicColumns(row => row.Extras, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "Score", Title = "分数", Order = 0 }
                })));
        using var destination = new MemoryStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination);

        using var workbook = Open(destination);
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal(7, sheet.GetRow(1).GetCell(0).NumericCellValue);
        Assert.Equal("动态", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal(12.5, sheet.GetRow(1).GetCell(2).NumericCellValue);
    }

    /// <summary>
    /// 验证前向导出应用表格、冻结窗格和样式。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldApplyForwardReportSubsetAndStyles()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new Row { Id = 1, Name = "A" }, new Row { Id = 2, Name = "B" } }, sheet => sheet
                .HeaderStyle(new ExcelCellStyle { Bold = true, BackgroundColor = new ExcelColor("FFDDDDDD") })
                .BodyStyle(new ExcelCellStyle { NumberFormat = "0" })
                .FreezePane(new ExcelFreezePaneDefinition { Rows = 1, Columns = 1 })
                .Table(new ExcelTableDefinition
                {
                    Name = "RowsTable",
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 }
                })));
        using var destination = new MemoryStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination);

        using var workbook = Open(destination);
        var sheet = (NPOI.XSSF.UserModel.XSSFSheet)workbook.GetSheetAt(0);
        Assert.Equal(1, sheet.PaneInformation.HorizontalSplitPosition);
        Assert.Equal(1, sheet.PaneInformation.VerticalSplitPosition);
        Assert.Single(sheet.GetTables());
        Assert.True(sheet.GetRow(0).GetCell(0).CellStyle.GetFont(workbook).IsBold);
    }

    /// <summary>
    /// 验证未请求表格时输出独立自动筛选。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldApplyAutoFilterWhenNoTableIsRequested()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new Row { Id = 1, Name = "A" } }, sheet => sheet.AutoFilter(
                new ExcelAutoFilterDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1 }
                })));
        using var destination = new MemoryStream();

        await new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination);

        using var workbook = Open(destination);
        var sheet = (NPOI.XSSF.UserModel.XSSFSheet)workbook.GetSheetAt(0);
        Assert.NotNull(sheet.GetCTWorksheet().autoFilter);
    }

    /// <summary>
    /// 创建用于导出验证的工作簿请求。
    /// </summary>
    /// <param name="count">测试数据行数。</param>
    /// <returns>包含测试数据和所需映射的工作簿导出请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(int count)
        => ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            Enumerable.Range(1, count).Select(index => new Row { Id = index, Name = $"Name {index}" })));

    /// <summary>
    /// 从导出流的副本打开工作簿。
    /// </summary>
    /// <param name="stream">包含导出工作簿的内存流。</param>
    /// <returns>从独立流副本加载的工作簿。</returns>
    private static XSSFWorkbook Open(MemoryStream stream)
        => new(new MemoryStream(stream.ToArray()));

    /// <summary>
    /// 用于基本导出验证的数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置测试数据编号。
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 用于动态列和单元格类型验证的数据行。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置测试数据编号。
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 用于验证动态数值单元格导出的额外列值。
        /// </summary>
        public IDictionary<string, object> Extras = new Dictionary<string, object>
        {
            ["Score"] = 12.5
        };
    }

    /// <summary>
    /// 记录枚举器创建次数的测试序列。
    /// </summary>
    /// <typeparam name="T">测试序列的元素类型。</typeparam>
    private sealed class TrackingEnumerable<T> : IEnumerable<T>
    {
        /// <summary>
        /// 用于记录枚举次数的原始数据源。
        /// </summary>
        private readonly IEnumerable<T> _source;
        /// <summary>
        /// 初始化一个 <see cref="TrackingEnumerable{T}"/> 类型的实例。
        /// </summary>
        /// <param name="source">待跟踪枚举次数的数据源。</param>
        public TrackingEnumerable(IEnumerable<T> source) => _source = source;
        /// <summary>
        /// 获取或设置枚举器创建次数。
        /// </summary>
        public int EnumeratorCount { get; private set; }
        /// <inheritdoc />
        public IEnumerator<T> GetEnumerator()
        {
            EnumeratorCount++;
            return _source.GetEnumerator();
        }
        /// <inheritdoc />
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    /// <summary>
    /// 记录释放状态的内存流。
    /// </summary>
    private sealed class TrackingMemoryStream : MemoryStream
    {
        /// <summary>
        /// 获取或设置流是否已被释放。
        /// </summary>
        public bool IsDisposed { get; private set; }
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 记录同步和异步写入次数的不可定位流。
    /// </summary>
    private sealed class TrackingNonSeekableStream : Stream
    {
        /// <summary>
        /// 获取保存测试输出内容的底层内存流。
        /// </summary>
        public MemoryStream Inner { get; } = new();
        /// <summary>
        /// 获取或设置异步写入次数。
        /// </summary>
        public int AsyncWriteCount { get; private set; }
        /// <summary>
        /// 获取或设置同步写入次数。
        /// </summary>
        public int SyncWriteCount { get; private set; }
        /// <summary>
        /// 获取或设置流是否已被释放。
        /// </summary>
        public bool IsDisposed { get; private set; }
        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        /// <inheritdoc />
        public override void Flush() => Inner.Flush();
        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken) => Inner.FlushAsync(cancellationToken);
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            SyncWriteCount++;
            Inner.Write(buffer, offset, count);
        }
        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            AsyncWriteCount++;
            return Inner.WriteAsync(buffer, cancellationToken);
        }
        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            IsDisposed = true;
            base.Dispose(disposing);
        }
    }
}
