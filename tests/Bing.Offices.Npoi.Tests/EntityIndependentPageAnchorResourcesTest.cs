using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;
using PageDocument = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证独立逐页锚点的异常、取消和文件资源合同。
/// </summary>
public sealed class EntityIndependentPageAnchorResourcesTest
{
    /// <summary>
    /// 验证工厂异常保全模板和输出流且仅观察一次。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task PageFooter_FactoryFailure_ShouldPreserveStreamsAndObserveOnce(bool xls,
        bool asynchronous)
    {
        var failure = new InvalidOperationException("签字区工厂失败");
        var layout = CreateLayout(_ => throw failure);
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("明细");
        using var template = new MemoryStream();
        workbook.Write(template, true);
        template.Position = 0;
        var templateBytes = template.ToArray();
        using var output = new MemoryStream();
        var original = new byte[] { 7, 4, 2, 9, 1 };
        output.Write(original);
        var observer = new IndependentPageAnchorRecordingObserver();
        var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });
        BingOfficesExportException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesExportException>(() => exporter.ExportForTemplateAsync(
                CreateDocument(), layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        else
            error = Assert.Throws<BingOfficesExportException>(() => exporter.ExportForTemplate(
                CreateDocument(), layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Write, error.Stage);
        Assert.Same(failure, error.InnerException);
        Assert.Same(error, Assert.Single(observer.Exceptions));
        Assert.Equal(original, output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证预取消和签字区执行中取消保全旧文件并清理暂存。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="before">是否在开始前取消。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task PageFooter_Cancellation_ShouldPreserveTargetAndCleanup(bool asynchronous, bool before)
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-footer-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "invoice.xlsx");
        var original = new byte[] { 7, 4, 2, 9, 1 };
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        var calls = 0;
        var layout = CreateLayout(_ =>
        {
            calls++;
            cancellation.Cancel();
            cancellation.Token.ThrowIfCancellationRequested();
            return 0;
        });
        var observer = new IndependentPageAnchorRecordingObserver();
        var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });
        if (before)
            cancellation.Cancel();
        try
        {
            OperationCanceledException error;
            if (asynchronous)
                error = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                    exporter.ExportEntityToFileAsync(CreateDocument(), layout, path, cancellation.Token));
            else
                error = Assert.ThrowsAny<OperationCanceledException>(() =>
                    exporter.ExportEntityToFile(CreateDocument(), layout, path, cancellation.Token));
            Assert.Equal(cancellation.Token, error.CancellationToken);
            Assert.Equal(before ? 0 : 1, calls);
            Assert.Empty(observer.Exceptions);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    /// <summary>
    /// 验证提交失败保全旧文件并仅观察一次。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageFooter_CommitFailure_ShouldPreserveTargetAndCleanup(bool asynchronous)
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-footer-commit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "invoice.xlsx");
        var original = new byte[] { 7, 4, 2, 9, 1 };
        File.WriteAllBytes(path, original);
        try
        {
            using var locked = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var observer = new IndependentPageAnchorRecordingObserver();
            var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });
            BingOfficesFileCommitException error;
            if (asynchronous)
                error = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                    exporter.ExportEntityToFileAsync(CreateDocument(), CreateLayout(items => items.Count), path));
            else
                error = Assert.Throws<BingOfficesFileCommitException>(() =>
                    exporter.ExportEntityToFile(CreateDocument(), CreateLayout(items => items.Count), path));
            Assert.Equal(BingOfficesStage.Commit, error.Stage);
            Assert.Equal(BingOfficesErrorCode.FileCommitFailed, error.Code);
            Assert.Same(error, Assert.Single(observer.Exceptions));
            var actual = new byte[original.Length];
            Assert.Equal(original.Length, locked.Length);
            Assert.Equal(actual.Length, locked.Read(actual, 0, actual.Length));
            Assert.Equal(original, actual);
            Assert.Equal(new[] { path }, Directory.GetFiles(directory));
        }
        finally
        {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    /// <summary>
    /// 创建使用工厂的逐页签字区布局。
    /// </summary>
    /// <param name="factory">签字区值工厂。</param>
    /// <returns>测试布局。</returns>
    private static ExcelEntityLayout<PageDocument> CreateLayout(Func<IReadOnlyList<PageLine>, int> factory) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.Cell("B1", factory))));

    /// <summary>
    /// 创建跨页采购单。
    /// </summary>
    /// <returns>测试单据。</returns>
    private static PageDocument CreateDocument() => new()
    {
        Lines = Enumerable.Range(1, 3).Select(index => new PageLine
            { Group = "G", Code = "L" + index, Amount = index }).ToList()
    };
}
