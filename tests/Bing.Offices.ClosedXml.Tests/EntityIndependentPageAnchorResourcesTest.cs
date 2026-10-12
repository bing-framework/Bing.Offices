using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.ClosedXml.Tests.Entities;
using Xunit;
using PageDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证逐页锚点导出的流、取消和文件资源职责。
/// </summary>
public sealed class EntityIndependentPageAnchorResourcesTest
{
    /// <summary>
    /// 验证页尾工厂异常只观察一次并保留调用方流资源。
    /// </summary>
    [Fact]
    public void PageFooter_FactoryFailure_ShouldPreserveStreamsAndObserveOnce()
    {
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 4, 2, 8 });
        var observer = new EntityPageAnchorFailureObserver();
        var expected = new InvalidOperationException("页尾值计算失败");
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            exporter.ExportForTemplate(CreateDocument(), CreateFactoryFailureLayout(expected),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Same(expected, exception.InnerException);
        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(new byte[] { 4, 2, 8 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证页尾工厂异常的异步入口只观察一次并保留调用方流资源。
    /// </summary>
    [Fact]
    public async Task PageFooter_FactoryFailureAsync_ShouldPreserveStreamsAndObserveOnce()
    {
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 4, 2, 8 });
        var observer = new EntityPageAnchorFailureObserver();
        var expected = new InvalidOperationException("异步页尾值计算失败");
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });

        var exception = await Assert.ThrowsAsync<BingOfficesExportException>(() =>
            exporter.ExportForTemplateAsync(CreateDocument(), CreateFactoryFailureLayout(expected),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Same(expected, exception.InnerException);
        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(new byte[] { 4, 2, 8 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证同步导出预取消时不改写目标并保留模板流。
    /// </summary>
    [Fact]
    public void PageFooter_PreCanceled_ShouldPreserveStreams()
    {
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 3, 1, 4 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(), CreateLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output,
                cancellation.Token));
        Assert.Equal(new byte[] { 3, 1, 4 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证异步导出预取消时不改写目标并保留模板流。
    /// </summary>
    [Fact]
    public async Task PageFooter_PreCanceledAsync_ShouldPreserveStreams()
    {
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 3, 1, 4 });
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(CreateDocument(), CreateLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output,
                cancellation.Token));
        Assert.Equal(new byte[] { 3, 1, 4 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证同步页尾工厂执行期间取消时不提交部分结果。
    /// </summary>
    [Fact]
    public void PageFooter_MidFactoryCancellation_ShouldPreserveStreams()
    {
        using var cancellation = new CancellationTokenSource();
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 5, 5, 5 });

        Assert.Throws<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(),
                CreateCancellationLayout(cancellation), new ExcelEntityTemplateOptions(template, leaveOpen: true),
                output, cancellation.Token));
        Assert.Equal(new byte[] { 5, 5, 5 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证异步页尾工厂执行期间取消时不提交部分结果。
    /// </summary>
    [Fact]
    public async Task PageFooter_MidFactoryCancellationAsync_ShouldPreserveStreams()
    {
        using var cancellation = new CancellationTokenSource();
        using var template = CreateTemplate();
        using var output = CreateTarget(new byte[] { 5, 5, 5 });

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(CreateDocument(),
                CreateCancellationLayout(cancellation), new ExcelEntityTemplateOptions(template, leaveOpen: true),
                output, cancellation.Token));
        Assert.Equal(new byte[] { 5, 5, 5 }, output.ToArray());
        Assert.True(template.CanRead);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证同步提交遇到锁定旧目标时保留原文件、清理暂存并观察一次。
    /// </summary>
    [Fact]
    public void PageFooter_LockedTargetCommitFailure_ShouldPreserveFileAndObserveOnce()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, original);
        var observer = new EntityPageAnchorFailureObserver();
        try
        {
            using var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var exception = Assert.Throws<BingOfficesFileCommitException>(() =>
                new ClosedXmlExcelExporter(exceptionObservers: new[] { observer })
                    .ExportEntityToFile(CreateDocument(), CreateLayout(), path));
            Assert.Equal(1, observer.Count);
            Assert.Same(exception, observer.LastException);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证异步提交遇到锁定旧目标时保留原文件、清理暂存并观察一次。
    /// </summary>
    [Fact]
    public async Task PageFooter_LockedTargetCommitFailureAsync_ShouldPreserveFileAndObserveOnce()
    {
        var directory = CreateTemporaryDirectory();
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, original);
        var observer = new EntityPageAnchorFailureObserver();
        try
        {
            using var locked = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                new ClosedXmlExcelExporter(exceptionObservers: new[] { observer })
                    .ExportEntityToFileAsync(CreateDocument(), CreateLayout(), path));
            Assert.Equal(1, observer.Count);
            Assert.Same(exception, observer.LastException);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 创建页尾工厂异常布局。
    /// </summary>
    private static ExcelEntityLayout<PageDocument> CreateFactoryFailureLayout(Exception expected) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN",
                    footer => footer.Cell<string>("B1", _ => throw expected))));

    /// <summary>
    /// 创建在页尾工厂中取消的布局。
    /// </summary>
    private static ExcelEntityLayout<PageDocument> CreateCancellationLayout(
        CancellationTokenSource cancellation) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN",
                    footer => footer.Cell("B1", _ =>
                    {
                        cancellation.Cancel();
                        return "取消";
                    }))));

    /// <summary>
    /// 创建基础页尾布局。
    /// </summary>
    private static ExcelEntityLayout<PageDocument> CreateLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.Cell("B1", "签字"))));

    /// <summary>
    /// 创建用于测试的采购明细。
    /// </summary>
    private static PageDocument CreateDocument() => new()
    {
        Lines = Enumerable.Range(1, 3).Select(index => new PageLine
        {
            Group = "G",
            Code = "L" + index,
            Amount = index
        }).ToList()
    };

    /// <summary>
    /// 创建空白明细模板。
    /// </summary>
    private static MemoryStream CreateTemplate()
    {
        using var workbook = new ClosedXML.Excel.XLWorkbook();
        workbook.Worksheets.Add("明细");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建带有既有字节的目标流。
    /// </summary>
    private static MemoryStream CreateTarget(byte[] original)
    {
        var stream = new MemoryStream();
        stream.Write(original, 0, original.Length);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建本测试专用的临时目录。
    /// </summary>
    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-page-anchor-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }
}
