using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Streaming;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.Streaming;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>流式 API 的资源所有权、清理失败和异步出口职责测试。</summary>
public class ExcelStreamingLifecycleTest
{
    [Fact]
    public void Sxssf_Dispose_ShouldDeleteCompressedFilesAcrossWindows()
    {
        var book = new SXSSFWorkbook(2) { CompressTempFiles = true };
        var sheet = (SXSSFSheet)book.CreateSheet("Data");
        for (var i = 0; i < 12; i++) sheet.CreateRow(i).CreateCell(0).SetCellValue("值" + i);
        var path = sheet.SheetDataWriter.TempFileInfo.FullName;
        Assert.True(File.Exists(path)); Assert.Null(sheet.GetRow(0));
        NpoiExcelStreamingExporter.DisposeWorkbook(book);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void Sxssf_CleanupFailure_ShouldContinueOtherWriters_AndBeObservable()
    {
        var book = new SXSSFWorkbook(1);
        var first = (SXSSFSheet)book.CreateSheet("First");
        var second = (SXSSFSheet)book.CreateSheet("Second");
        first.CreateRow(0).CreateCell(0).SetCellValue("一");
        second.CreateRow(0).CreateCell(0).SetCellValue("二");
        first.FlushRows(); second.FlushRows();
        var firstPath = first.SheetDataWriter.TempFileInfo.FullName;
        var secondPath = second.SheetDataWriter.TempFileInfo.FullName;
        // Windows 文件共享机制模拟删除失败，随后仍须清理其他工作表。
        using (var locked = new FileStream(firstPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            Assert.Throws<AggregateException>(() => NpoiExcelStreamingExporter.DisposeWorkbook(book));
        Assert.False(File.Exists(secondPath));
        File.Delete(firstPath);
        var primary = new IOException("主异常");
        var attempts = 0;
        StreamingCleanup.RunAll(primary, () => { attempts++; throw new IOException("清理一"); },
            () => { attempts++; throw new IOException("清理二"); });
        Assert.Equal(2, attempts);
        var errors = Assert.IsType<List<Exception>>(primary.Data["Bing.Offices.Streaming.CleanupErrors"]);
        Assert.Equal(2, Assert.IsType<AggregateException>(Assert.Single(errors)).InnerExceptions.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_StagingAndOutputFailure_ShouldCleanAndLeaveCallerOpen(bool packagingFailure)
    {
        var folder = NewFolder();
        try
        {
            using var output = new AsyncOnlyOutput(fail: !packagingFailure);
            var options = new ExcelStreamingExportOptions { TempDirectory = folder, MaxTemporaryBytes = packagingFailure ? 1 : 1024 * 1024 };
            await Assert.ThrowsAnyAsync<Exception>(() => new NpoiExcelStreamingExporter().ExportAsync(Request(), output, options));
            Assert.True(output.CanWrite); Assert.Empty(Directory.GetFileSystemEntries(folder));
        }
        finally { Directory.Delete(folder); }
    }

    [Fact]
    public async Task ExportAsync_AndFileExports_ShouldReturnCompleteXlsx_WithAsyncCopy()
    {
        var folder = NewFolder();
        try
        {
            IExcelStreamingExporter exporter = new NpoiExcelStreamingExporter();
            using var output = new AsyncOnlyOutput(false);
            await exporter.ExportAsync(Request(), output, new() { TempDirectory = folder });
            Assert.True(output.AsyncWrites > 0); Assert.True(output.CanWrite);
            Check(output.ToArray());
            var path = Path.Combine(folder, "out.xlsx");
            exporter.ExportToFile(Request(), path, new() { TempDirectory = folder });
            Check(File.ReadAllBytes(path));
            await exporter.ExportToFileAsync(Request(), path, new() { TempDirectory = folder });
            Check(File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(folder));
            File.Delete(path);
        }
        finally { Directory.Delete(folder); }
        static void Check(byte[] bytes)
        {
            using var input = new MemoryStream(bytes);
            var values = new List<int>();
            var result = new NpoiExcelStreamingImporter().Import(ExcelStreamingImport.Sheet<NpoiExcelStreamingImporterTest.NumberRow>(0),
                input, row => { values.Add(row.Item.Count); return ExcelRowReadDecision.Continue; });
            Assert.True(result.IsSuccess); Assert.Equal(new[] { 1, 2 }, values);
        }
    }

    [Fact]
    public void CancellationBeforeWrite_ShouldNotEnumerateOrCreateTemporaryFiles()
    {
        var folder = NewFolder();
        try
        {
            using var output = new MemoryStream();
            Assert.ThrowsAny<OperationCanceledException>(() => new NpoiExcelStreamingExporter().Export(Request(), output,
                new() { TempDirectory = folder }, new CancellationToken(true)));
            Assert.Empty(output.ToArray()); Assert.Empty(Directory.GetFileSystemEntries(folder));
        }
        finally { Directory.Delete(folder); }
    }

    [Fact]
    public void Registration_ShouldPreserveCustomServices_AndKeepOptInSeparate()
    {
        var ordinary = new ServiceCollection().AddBingOfficesNpoi();
        Assert.DoesNotContain(ordinary, descriptor => descriptor.ServiceType == typeof(IExcelStreamingImporter));
        Assert.DoesNotContain(ordinary, descriptor => descriptor.ServiceType == typeof(IExcelStreamingExporter));
        IExcelStreamingImporter importer = new NpoiExcelStreamingImporter();
        IExcelStreamingExporter exporter = new NpoiExcelStreamingExporter();
        var services = new ServiceCollection().AddSingleton(importer).AddSingleton(exporter);
        services.AddBingOfficesNpoiStreaming().AddBingOfficesNpoiStreaming();
        using var provider = services.BuildServiceProvider();
        Assert.Same(importer, provider.GetRequiredService<IExcelStreamingImporter>());
        Assert.Same(exporter, provider.GetRequiredService<IExcelStreamingExporter>());
        Assert.Single(provider.GetServices<IExcelStreamingImporter>());
        using var defaults = new ServiceCollection().AddBingOfficesNpoiStreaming().BuildServiceProvider();
        Assert.IsType<NpoiExcelStreamingImporter>(defaults.GetRequiredService<IExcelStreamingImporter>());
    }

    private static ExcelWorkbookExportRequest Request() => ExcelExport.Workbook(book => book.AddSheet("Data",
        new[] { new NpoiExcelStreamingImporterTest.NumberRow { Count = 1 }, new NpoiExcelStreamingImporterTest.NumberRow { Count = 2 } }));
    private static string NewFolder()
    { var path = Path.Combine(Path.GetTempPath(), "bing-lifecycle-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
    private sealed class AsyncOnlyOutput : MemoryStream
    {
        private readonly bool _fail;
        internal int AsyncWrites;
        internal AsyncOnlyOutput(bool fail) => _fail = fail;
        public override void Write(byte[] buffer, int offset, int count) => throw new InvalidOperationException("禁止同步写出");
        public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
        {
            token.ThrowIfCancellationRequested(); AsyncWrites++;
            if (_fail) throw new IOException("输出失败");
            base.Write(buffer, offset, count); return Task.CompletedTask;
        }
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
            => new(WriteAsync(buffer.ToArray(), 0, buffer.Length, token));
    }
}
