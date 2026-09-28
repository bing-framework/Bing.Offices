using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// 报表级工作簿导出策略的职责测试。
/// </summary>
public sealed class ExcelWorkbookExportStrategyTest
{
    /// <summary>
    /// 验证完整工作簿模式按格式输出可读数据。
    /// </summary>
    /// <param name="provider">完整工作簿 Provider。</param>
    /// <param name="format">工作簿格式。</param>
    /// <param name="async">是否使用异步入口。</param>
    [Theory]
    [InlineData("NPOI", ExcelFormat.Xls, false)]
    [InlineData("NPOI", ExcelFormat.Xls, true)]
    [InlineData("NPOI", ExcelFormat.Xlsx, false)]
    [InlineData("NPOI", ExcelFormat.Xlsx, true)]
    [InlineData("ClosedXML", ExcelFormat.Xlsx, false)]
    [InlineData("ClosedXML", ExcelFormat.Xlsx, true)]
    public async Task CompleteWorkbook_ShouldWriteReadableWorkbook(string provider, ExcelFormat format, bool async)
    {
        var strategy = new ExcelWorkbookExportStrategy(provider == "NPOI"
            ? new NpoiExcelExporter() : new ClosedXmlExcelExporter());
        using var output = new MemoryStream();

        if (async)
            await strategy.ExportAsync(Request(format), output, ExcelWorkbookExportMode.CompleteWorkbook);
        else
            strategy.Export(Request(format), output, ExcelWorkbookExportMode.CompleteWorkbook);

        Assert.True(output.CanWrite);
        AssertWorkbook(output.ToArray());
    }

    /// <summary>
    /// 验证前向流式模式按同步和异步路径生成工作簿。
    /// </summary>
    /// <param name="async">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ForwardStreaming_ShouldWriteReadableWorkbook(bool async)
    {
        var strategy = new ExcelWorkbookExportStrategy(streamingExporter: new SpreadCheetahStreamingExcelExporter());
        using var output = new MemoryStream();
        var options = new ExcelStreamingExportOptions { BatchSize = 2 };

        if (async)
            await strategy.ExportAsync(Request(ExcelFormat.Xlsx), output,
                ExcelWorkbookExportMode.ForwardStreaming, options);
        else
            strategy.Export(Request(ExcelFormat.Xlsx), output,
                ExcelWorkbookExportMode.ForwardStreaming, options);

        Assert.True(output.CanWrite);
        AssertWorkbook(output.ToArray());
    }

    /// <summary>
    /// 验证完整工作簿模式保留模板导出能力。
    /// </summary>
    /// <param name="provider">完整工作簿 Provider。</param>
    [Theory]
    [InlineData("NPOI")]
    [InlineData("ClosedXML")]
    public void CompleteWorkbook_ShouldAcceptTemplate(string provider)
    {
        IExcelExporter exporter = provider == "NPOI" ? new NpoiExcelExporter() : new ClosedXmlExcelExporter();
        var strategy = new ExcelWorkbookExportStrategy(exporter);
        using var templateOutput = new MemoryStream();
        exporter.Export(Request(ExcelFormat.Xlsx), templateOutput);
        using var template = new MemoryStream(templateOutput.ToArray());
        var request = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .AddSheet("Rows", new[] { new Row { Id = 7, Name = "采购" } }));
        using var output = new MemoryStream();

        strategy.Export(request, output, ExcelWorkbookExportMode.CompleteWorkbook);

        Assert.True(template.CanRead);
        AssertWorkbook(output.ToArray());
    }

    /// <summary>
    /// 验证不支持的模式、格式与模板在写入前被拒绝。
    /// </summary>
    [Fact]
    public void UnsupportedSelection_ShouldLeaveDestinationEmpty()
    {
        var strategy = new ExcelWorkbookExportStrategy(new NpoiExcelExporter(),
            new SpreadCheetahStreamingExcelExporter());
        using var output = new MemoryStream();
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var templateRequest = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .AddSheet("Rows", new[] { new Row { Id = 7, Name = "采购" } }));

        AssertPreflight(() => strategy.Export(templateRequest, output, ExcelWorkbookExportMode.ForwardStreaming));
        AssertPreflight(() => strategy.Export(Request(ExcelFormat.Xls), output,
            ExcelWorkbookExportMode.ForwardStreaming));
        Assert.Throws<ArgumentOutOfRangeException>(() => strategy.Export(Request(ExcelFormat.Xlsx), output,
            (ExcelWorkbookExportMode)100));
        Assert.Throws<ArgumentException>(() => strategy.Export(Request(ExcelFormat.Xlsx), output,
            ExcelWorkbookExportMode.CompleteWorkbook, new ExcelStreamingExportOptions()));
        Assert.Empty(output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证未配置的模式及格式不匹配不会切换导出器。
    /// </summary>
    [Fact]
    public void MissingOrIncompatibleProvider_ShouldFailBeforeOutput()
    {
        using var output = new MemoryStream();
        var streamingOnly = new ExcelWorkbookExportStrategy(streamingExporter: new SpreadCheetahStreamingExcelExporter());
        AssertPreflight(() => streamingOnly.Export(Request(ExcelFormat.Xlsx), output,
            ExcelWorkbookExportMode.CompleteWorkbook));

        var closedXmlOnly = new ExcelWorkbookExportStrategy(new ClosedXmlExcelExporter());
        AssertPreflight(() => closedXmlOnly.Export(Request(ExcelFormat.Xls), output,
            ExcelWorkbookExportMode.CompleteWorkbook));
        AssertPreflight(() => closedXmlOnly.Export(Request(ExcelFormat.Xlsx), output,
            ExcelWorkbookExportMode.ForwardStreaming));
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证文件入口沿用所选 Provider 的原子提交路径。
    /// </summary>
    /// <param name="mode">工作簿导出模式。</param>
    /// <param name="async">是否使用异步入口。</param>
    [Theory]
    [InlineData(ExcelWorkbookExportMode.CompleteWorkbook, false)]
    [InlineData(ExcelWorkbookExportMode.CompleteWorkbook, true)]
    [InlineData(ExcelWorkbookExportMode.ForwardStreaming, false)]
    [InlineData(ExcelWorkbookExportMode.ForwardStreaming, true)]
    public async Task FileExport_ShouldCommitSelectedWorkbook(ExcelWorkbookExportMode mode, bool async)
    {
        var strategy = new ExcelWorkbookExportStrategy(new NpoiExcelExporter(),
            new SpreadCheetahStreamingExcelExporter());
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-strategy-{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            if (async)
                await strategy.ExportToFileAsync(Request(ExcelFormat.Xlsx), path, mode);
            else
                strategy.ExportToFile(Request(ExcelFormat.Xlsx), path, mode);

            AssertWorkbook(File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>
    /// 验证预取消不会修改已有文件。
    /// </summary>
    /// <param name="mode">工作簿导出模式。</param>
    [Theory]
    [InlineData(ExcelWorkbookExportMode.CompleteWorkbook)]
    [InlineData(ExcelWorkbookExportMode.ForwardStreaming)]
    public async Task PreCancelledFileExport_ShouldPreserveTarget(ExcelWorkbookExportMode mode)
    {
        var strategy = new ExcelWorkbookExportStrategy(new NpoiExcelExporter(),
            new SpreadCheetahStreamingExcelExporter());
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-strategy-{Guid.NewGuid():N}.xlsx");
        var original = new byte[] { 1, 2, 3 };
        using var source = new CancellationTokenSource();
        source.Cancel();
        try
        {
            File.WriteAllBytes(path, original);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                strategy.ExportToFileAsync(Request(ExcelFormat.Xlsx), path, mode,
                    cancellationToken: source.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>
    /// 创建包含一行数据的导出请求。
    /// </summary>
    /// <param name="format">目标工作簿格式。</param>
    /// <returns>工作簿请求。</returns>
    private static ExcelWorkbookExportRequest Request(ExcelFormat format) =>
        ExcelExport.Workbook(book => book.Format(format)
            .AddSheet("Rows", new[] { new Row { Id = 7, Name = "采购" } }));

    /// <summary>
    /// 验证工作簿结构和完整数据。
    /// </summary>
    /// <param name="content">生成的工作簿内容。</param>
    private static void AssertWorkbook(byte[] content)
    {
        using var input = new MemoryStream(content);
        using var workbook = WorkbookFactory.Create(input);
        Assert.Equal(1, workbook.NumberOfSheets);
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal("Rows", sheet.SheetName);
        Assert.Equal("Id", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Name", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal(7, sheet.GetRow(1).GetCell(0).NumericCellValue);
        Assert.Equal("采购", sheet.GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 验证异常属于导出预检阶段。
    /// </summary>
    /// <param name="action">应失败的导出操作。</param>
    private static void AssertPreflight(Action action)
    {
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(action);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 导出测试数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置编号。
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }
}
