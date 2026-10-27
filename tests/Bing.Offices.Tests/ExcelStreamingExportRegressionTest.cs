using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.IO;
using Bing.Offices.Streaming;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>文件总额度与动态读取异常的四出口职责回归。</summary>
public class ExcelStreamingExportRegressionTest
{
    [Theory]
    [InlineData(false, false)] [InlineData(true, false)]
    [InlineData(false, true)] [InlineData(true, true)]
    public async Task FileExport_ShouldBudgetStagingAndCommitTogether_AndPreserveTarget(bool asynchronous, bool customCommitter)
    {
        var root = NewFolder();
        var temporary = Path.Combine(root, "sessions");
        Directory.CreateDirectory(temporary);
        var path = Path.Combine(root, "report.xlsx");
        var original = Encoding.UTF8.GetBytes("原文件必须完整保留");
        File.WriteAllBytes(path, original);
        var committer = new RecordingCommitter();
        IExcelStreamingExporter exporter = new NpoiExcelStreamingExporter(committer: customCommitter ? committer : null);
        var request = ExcelExport.Workbook(book => book.AddSheet("Data", Rows()));
        try
        {
            using var sample = new MemoryStream();
            exporter.Export(request, sample);
            var quota = sample.Length + 1024;
            var options = new ExcelStreamingExportOptions { TempDirectory = temporary, MaxTemporaryBytes = quota };
            // 调用方流只计库暂存；相同额度仍应成功，并完整回读。
            using var output = new MemoryStream();
            if (asynchronous) await exporter.ExportAsync(request, output, options);
            else exporter.Export(request, output, options);
            Assert.True(output.CanWrite); Assert.True(output.Length <= quota); Assert.True(output.Length * 2 > quota);
            CheckRows(output.ToArray());
            if (asynchronous)
                await Assert.ThrowsAsync<StreamingResourceLimitException>(() => exporter.ExportToFileAsync(request, path, options));
            else Assert.Throws<StreamingResourceLimitException>(() => exporter.ExportToFile(request, path, options));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFileSystemEntries(temporary));
            Assert.Equal(new[] { path }, Directory.GetFiles(root));
            Assert.Equal(customCommitter ? 1 : 0, committer.Calls);

            var sufficient = new ExcelStreamingExportOptions { TempDirectory = temporary, MaxTemporaryBytes = quota * 2 };
            if (asynchronous) await exporter.ExportToFileAsync(request, path, sufficient);
            else exporter.ExportToFile(request, path, sufficient);
            CheckRows(File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFileSystemEntries(temporary));
            Assert.Equal(new[] { path }, Directory.GetFiles(root));
        }
        finally { Directory.Delete(root, true); }
    }

    public static IEnumerable<object[]> ExceptionCases()
    {
        foreach (var entry in new[] { "stream", "stream-async", "file", "file-async" })
        foreach (var failure in new[] { "ordinary", "office", "cancel" })
            yield return new object[] { entry, failure };
    }

    [Theory]
    [MemberData(nameof(ExceptionCases))]
    public async Task Export_DynamicFailure_ShouldPreserveContextAndNotifyOnce(string entry, string failure)
    {
        var root = NewFolder();
        var temporary = Path.Combine(root, "sessions");
        Directory.CreateDirectory(temporary);
        var path = Path.Combine(root, "report.xlsx");
        var original = Encoding.UTF8.GetBytes("原有文件");
        File.WriteAllBytes(path, original);
        Exception expected = failure switch
        {
            "office" => new BingOfficesExportException("已分类异常", code: BingOfficesErrorCode.UserExtensionFailed),
            "cancel" => new OperationCanceledException("用户取消"),
            _ => new InvalidOperationException("动态读取失败")
        };
        var item = new DynamicRow { Failure = expected };
        var observer = new RecordingObserver();
        IExcelStreamingExporter exporter = new NpoiExcelStreamingExporter(observers: new[] { observer });
        var request = Request(item);
        var options = new ExcelStreamingExportOptions { TempDirectory = temporary };
        using var stream = new MemoryStream();
        try
        {
            var actual = await Record.ExceptionAsync(async () =>
            {
                switch (entry)
                {
                    case "stream": exporter.Export(request, stream, options); break;
                    case "stream-async": await exporter.ExportAsync(request, stream, options); break;
                    case "file": exporter.ExportToFile(request, path, options); break;
                    case "file-async": await exporter.ExportToFileAsync(request, path, options); break;
                }
            });
            Assert.Equal(1, item.Reads);
            if (failure == "ordinary")
            {
                var office = Assert.IsType<BingOfficesExportException>(actual);
                Assert.Same(expected, office.InnerException);
                Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, office.Code);
                Assert.Equal(BingOfficesStage.Validate, office.Stage);
                Assert.Equal(BingOfficesOperation.Export, office.Operation);
                Assert.Equal("NPOI", office.Provider); Assert.Equal("Data", office.SheetName);
                Assert.Equal(2, office.RowIndex); Assert.Null(office.ColumnIndex);
                Assert.Equal("Values", office.PropertyName);
            }
            else Assert.Same(expected, actual);
            if (failure == "cancel") Assert.Empty(observer.Errors);
            else Assert.Same(actual, Assert.Single(observer.Errors));
            Assert.Empty(Directory.GetFileSystemEntries(temporary));
            Assert.Equal(new[] { path }, Directory.GetFiles(root));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.True(stream.CanWrite); Assert.Empty(stream.ToArray());
        }
        finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(ExcelFormat.Xls, false)] [InlineData(ExcelFormat.Xlsx, false)]
    [InlineData(ExcelFormat.Xls, true)] [InlineData(ExcelFormat.Xlsx, true)]
    public void DomExport_SharedDynamicReader_ShouldPreserveBothProviders(ExcelFormat format, bool fail)
    {
        var expected = new InvalidOperationException("动态读取失败");
        var item = new DynamicRow { Failure = fail ? expected : null };
        var observer = new RecordingObserver();
        var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });
        using var output = new MemoryStream();
        var request = Request(item, format);
        if (fail)
        {
            var actual = Assert.Throws<BingOfficesExportException>(() => exporter.Export(request, output));
            Assert.Same(expected, actual.InnerException); Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, actual.Code);
            Assert.Equal("Data", actual.SheetName); Assert.Equal(2, actual.RowIndex); Assert.Equal("Values", actual.PropertyName);
            Assert.Same(actual, Assert.Single(observer.Errors));
        }
        else
        {
            exporter.Export(request, output); output.Position = 0;
            using var book = WorkbookFactory.Create(output);
            var sheet = book.GetSheetAt(0);
            Assert.Equal(2, sheet.PhysicalNumberOfRows);
            Assert.Equal(new[] { "Code", "Extra" }, sheet.GetRow(0).Cells.Select(c => c.StringCellValue));
            Assert.Equal(new[] { "A", "中文" }, sheet.GetRow(1).Cells.Select(c => c.StringCellValue));
            Assert.Empty(observer.Errors);
        }
        Assert.Equal(1, item.Reads);
    }

    private static ExcelWorkbookExportRequest Request(DynamicRow row, ExcelFormat format = ExcelFormat.Xlsx)
        => ExcelExport.Workbook(book => book.Format(format).AddSheet("Data", new[] { row },
            sheet => sheet.DynamicColumns(item => item.Values, new[] { new ExcelDynamicColumnDefinition { Key = "extra", Title = "Extra" } })));

    private static string NewFolder()
    { var path = Path.Combine(Path.GetTempPath(), "bing-stream-fix-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
    private static IEnumerable<TextRow> Rows()
    {
        var random = new Random(17);
        for (var i = 0; i < 1000; i++)
        { var bytes = new byte[24]; random.NextBytes(bytes); yield return new TextRow { Value = Convert.ToBase64String(bytes) + i }; }
    }
    private static void CheckRows(byte[] bytes)
    {
        using var input = new MemoryStream(bytes);
        using var book = WorkbookFactory.Create(input);
        Assert.Equal(1, book.NumberOfSheets);
        var sheet = book.GetSheetAt(0); Assert.Equal("Data", sheet.SheetName);
        Assert.Equal("Value", sheet.GetRow(0).GetCell(0).StringCellValue);
        var expected = Rows().ToArray(); Assert.Equal(expected.Length + 1, sheet.PhysicalNumberOfRows);
        for (var i = 0; i < expected.Length; i++)
        { var cell = Assert.Single(sheet.GetRow(i + 1).Cells); Assert.Equal(expected[i].Value, cell.StringCellValue); }
    }
    public class TextRow { public string Value { get; set; } }
    public class DynamicRow
    {
        internal Exception Failure;
        internal int Reads;
        public string Code { get; set; } = "A";
        [DynamicColumn] public IDictionary<string, object> Values
        { get { Reads++; if (Failure != null) throw Failure; return new Dictionary<string, object> { ["extra"] = "中文" }; } set { } }
    }
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        internal List<BingOfficesException> Errors { get; } = new();
        public void Observe(BingOfficesException exception) => Errors.Add(exception);
    }
    private sealed class RecordingCommitter : IFileExportCommitter
    {
        private readonly DefaultFileExportCommitter _inner = new();
        internal int Calls;
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        { Calls++; _inner.Commit(path, write, cancellationToken, format); }
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync, CancellationToken cancellationToken, string format)
        { Calls++; return _inner.CommitAsync(path, writeAsync, cancellationToken, format); }
    }
}
