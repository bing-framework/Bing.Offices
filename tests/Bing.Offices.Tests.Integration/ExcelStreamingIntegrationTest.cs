using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>真实文件上的跨调用隔离与异步取消提交验证。</summary>
public class ExcelStreamingIntegrationTest
{
    [Fact]
    public async Task StreamingServices_ShouldIsolateConcurrentCallsAndCleanAllFiles()
    {
        var folder = Path.Combine(Path.GetTempPath(), "bing-stream-integration-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            using var provider = new ServiceCollection().AddBingOfficesNpoiStreaming().BuildServiceProvider();
            var exporter = provider.GetRequiredService<IExcelStreamingExporter>();
            var importer = provider.GetRequiredService<IExcelStreamingImporter>();
            var paths = new[] { Path.Combine(folder, "one.xlsx"), Path.Combine(folder, "two.xlsx") };
            await Task.WhenAll(paths.Select((path, index) => exporter.ExportToFileAsync(
                Request(index * 1000), path, new ExcelStreamingExportOptions { RowWindowSize = 2, TempDirectory = folder })));
            var bothEntered = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var entered = 0;
            await Task.WhenAll(paths.Select(async (path, index) =>
            {
                using var source = File.OpenRead(path);
                var count = 0;
                var result = await importer.ImportAsync(ExcelStreamingImport.Sheet<Row>("Data"), source, async (row, token) =>
                {
                    if (count == 0)
                    {
                        if (Interlocked.Increment(ref entered) == 2) bothEntered.SetResult(true);
                        await bothEntered.Task.WaitAsync(TimeSpan.FromSeconds(10), token);
                    }
                    Assert.True(row.IsValid); Assert.Equal(index * 1000 + count, row.Item.Value);
                    Assert.Equal(++count, row.RowIndex);
                    return ExcelRowReadDecision.Continue;
                }, new ExcelStreamingImportOptions { TempDirectory = folder });
                Assert.Equal(20, count); Assert.True(result.IsSuccess); Assert.True(source.CanRead);
            }));
            Assert.Equal(paths.OrderBy(path => path), Directory.GetFileSystemEntries(folder).OrderBy(path => path));
            foreach (var path in paths) File.Delete(path);
        }
        finally { Directory.Delete(folder); }
    }

    [Fact]
    public async Task ExportToFileAsync_ShouldPreserveOldTargetWhenCancelledAcrossWindow()
    {
        var folder = Path.Combine(Path.GetTempPath(), "bing-stream-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "data.xlsx");
        var original = Encoding.UTF8.GetBytes("已有文件");
        File.WriteAllBytes(path, original);
        try
        {
            using var cancellation = new CancellationTokenSource();
            using var provider = new ServiceCollection().AddBingOfficesNpoiStreaming().BuildServiceProvider();
            var exporter = provider.GetRequiredService<IExcelStreamingExporter>();
            var request = ExcelExport.Workbook(book => book.AddSheet("Data", Rows()));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => exporter.ExportToFileAsync(request, path,
                new ExcelStreamingExportOptions { RowWindowSize = 1, TempDirectory = folder }, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path)); Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(folder));
            IEnumerable<Row> Rows()
            {
                for (var index = 0; index < 10; index++)
                { if (index == 5) cancellation.Cancel(); yield return new Row { Value = index }; }
            }
        }
        finally { File.Delete(path); Directory.Delete(folder); }
    }

    private static ExcelWorkbookExportRequest Request(int offset) => ExcelExport.Workbook(book => book.AddSheet("Data",
        Enumerable.Range(0, 20).Select(index => new Row { Value = offset + index })));
    public class Row { public int Value { get; set; } }
}
