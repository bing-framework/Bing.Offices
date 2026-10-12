using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证各 Provider 对模板透视保留和更新模式的当前拒绝边界。
/// </summary>
public sealed class TemplatePivotProviderBoundaryTest
{
    /// <summary>
    /// 验证同步 Provider 在预检阶段拒绝模板透视模式且不触碰调用方流或数据源。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="update">是否通过更新入口隐含启用保留模式。</param>
    [Theory]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", true)]
    [InlineData("MiniExcel", false)]
    [InlineData("MiniExcel", true)]
    [InlineData("SpreadCheetah", false)]
    [InlineData("SpreadCheetah", true)]
    public void Sync_ShouldRejectTemplatePivotModeBeforeWritingOrEnumerating(string provider, bool update)
    {
        using var template = new TemplatePivotTrackingStream(CreateTemplateBytes());
        template.Position = 11;
        var originalTemplateBytes = template.ToArray();
        var enumerated = false;
        var request = CreateRequest(template, update, () => enumerated = true);
        using var destination = new TemplatePivotTrackingStream(new byte[] { 9, 8, 7, 6, 5 });
        destination.Position = 2;
        var originalBytes = destination.ToArray();
        var originalPosition = destination.Position;

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            Export(provider, request, destination));

        AssertUnsupported(exception, provider);
        Assert.Equal(originalBytes, destination.ToArray());
        Assert.Equal(originalPosition, destination.Position);
        Assert.True(destination.CanWrite);
        Assert.False(destination.WasDisposed);
        Assert.Equal(originalTemplateBytes, template.ToArray());
        Assert.Equal(11, template.Position);
        Assert.False(template.WasDisposed);
        Assert.False(enumerated);
        AssertPivotFeatureBoundary(provider);
    }

    /// <summary>
    /// 验证异步 Provider 在预检阶段拒绝模板透视模式且不触碰调用方流或数据源。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    /// <param name="update">是否通过更新入口隐含启用保留模式。</param>
    [Theory]
    [InlineData("NPOI", true)]
    [InlineData("ClosedXML", true)]
    [InlineData("MiniExcel", false)]
    [InlineData("MiniExcel", true)]
    [InlineData("SpreadCheetah", false)]
    [InlineData("SpreadCheetah", true)]
    public async Task Async_ShouldRejectTemplatePivotModeBeforeWritingOrEnumerating(string provider, bool update)
    {
        using var template = new TemplatePivotTrackingStream(CreateTemplateBytes());
        template.Position = 11;
        var originalTemplateBytes = template.ToArray();
        var enumerated = false;
        var request = CreateRequest(template, update, () => enumerated = true);
        using var destination = new TemplatePivotTrackingStream(new byte[] { 9, 8, 7, 6, 5 });
        destination.Position = 2;
        var originalBytes = destination.ToArray();
        var originalPosition = destination.Position;

        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            ExportAsync(provider, request, destination));

        AssertUnsupported(exception, provider);
        Assert.Equal(originalBytes, destination.ToArray());
        Assert.Equal(originalPosition, destination.Position);
        Assert.True(destination.CanWrite);
        Assert.False(destination.WasDisposed);
        Assert.Equal(originalTemplateBytes, template.ToArray());
        Assert.Equal(11, template.Position);
        Assert.False(template.WasDisposed);
        Assert.False(enumerated);
        AssertPivotFeatureBoundary(provider);
    }

    /// <summary>
    /// 创建带有效 XLSX 工作表的只读模板样本。
    /// </summary>
    /// <returns>由 ClosedXML Provider 公开导出得到的 XLSX 字节。</returns>
    private static byte[] CreateTemplateBytes()
    {
        using var output = new MemoryStream();
        var request = ExcelExport.Workbook(builder => builder.AddSheet("TemplateData",
            new[] { new TemplatePivotBoundaryRow { Region = "North", Amount = 10 } }));
        new ClosedXmlExcelExporter().Export(request, output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建启用模板透视能力并包含惰性追加 Sheet 的公开请求。
    /// </summary>
    /// <param name="template">调用方拥有的模板流。</param>
    /// <param name="update">是否添加更新定义。</param>
    /// <param name="onEnumeration">追加 Sheet 首次枚举时执行的回调。</param>
    /// <returns>经过公开 Builder 构建的请求。</returns>
    private static ExcelWorkbookExportRequest CreateRequest(Stream template, bool update,
        Action onEnumeration) => ExcelExport.Workbook(builder =>
    {
        builder.UseTemplate(template, leaveOpen: true)
            .AddSheet("Appended", EnumerateRow(onEnumeration));
        if (update)
            builder.UpdatePivotTable(Definition());
        else
            builder.PreserveTemplatePivotTables();
    });

    /// <summary>
    /// 创建一次枚举时通知测试的延迟数据源。
    /// </summary>
    /// <param name="onEnumeration">开始枚举时调用的回调。</param>
    /// <returns>仅包含一个实体的延迟序列。</returns>
    private static IEnumerable<TemplatePivotBoundaryRow> EnumerateRow(Action onEnumeration)
    {
        onEnumeration();
        yield return new TemplatePivotBoundaryRow { Region = "Extra", Amount = 1 };
    }

    /// <summary>
    /// 创建指向模板中已有工作表的合法更新定义。
    /// </summary>
    /// <returns>有效的旧字段透视定义。</returns>
    private static ExcelPivotTableDefinition Definition() => new()
    {
        Name = "SalesByRegion",
        SourceSheet = "TemplateData",
        SourceRange = new ExcelRangeDefinition
        {
            StartRow = 0, StartColumn = 0, EndRow = 1, EndColumn = 1
        },
        TargetSheet = "Summary",
        RowField = "Region",
        ValueField = "Amount"
    };

    /// <summary>
    /// 按 Provider 调用同步公开导出接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">待验证的工作簿请求。</param>
    /// <param name="destination">带哨兵内容的调用方流。</param>
    private static void Export(string provider, ExcelWorkbookExportRequest request, Stream destination)
    {
        if (provider == "SpreadCheetah")
            new SpreadCheetahStreamingExcelExporter().ExportBatches(request, destination);
        else
            CreateExporter(provider).Export(request, destination);
    }

    /// <summary>
    /// 按 Provider 调用异步公开导出接口。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <param name="request">待验证的工作簿请求。</param>
    /// <param name="destination">带哨兵内容的调用方流。</param>
    /// <returns>异步导出操作任务。</returns>
    private static Task ExportAsync(string provider, ExcelWorkbookExportRequest request, Stream destination)
    {
        if (provider == "SpreadCheetah")
            return new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination);
        return CreateExporter(provider).ExportAsync(request, destination);
    }

    /// <summary>
    /// 创建指定 Provider 的完整工作簿导出器。
    /// </summary>
    /// <param name="provider">Provider 名称。</param>
    /// <returns>完整工作簿导出器。</returns>
    private static IExcelExporter CreateExporter(string provider) => provider switch
    {
        "NPOI" => new NpoiExcelExporter(),
        "ClosedXML" => new ClosedXmlExcelExporter(),
        "MiniExcel" => new MiniExcelExcelExporter(),
        _ => throw new ArgumentOutOfRangeException(nameof(provider), provider, "未知 Provider。")
    };

    /// <summary>
    /// 验证拒绝结果的结构化错误元数据。
    /// </summary>
    /// <param name="exception">Provider 抛出的错误。</param>
    /// <param name="provider">预期 Provider 名称。</param>
    private static void AssertUnsupported(BingOfficesUnsupportedFeatureException exception, string provider)
    {
        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(provider, exception.Provider);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("不支持", exception.Message);
    }

    /// <summary>
    /// 验证各 Provider 的新建透视和模板透视能力声明与实际边界一致。
    /// </summary>
    /// <param name="provider">待验证的 Provider 名称。</param>
    private static void AssertPivotFeatureBoundary(string provider)
    {
        IExcelProviderFeatureDescriptor descriptor = provider switch
        {
            "NPOI" => new NpoiExcelExporter(),
            "ClosedXML" => new ClosedXmlExcelExporter(),
            "MiniExcel" => new MiniExcelExcelExporter(),
            _ => new SpreadCheetahStreamingExcelExporter()
        };

        if (provider is "NPOI" or "ClosedXML")
        {
            Assert.True(descriptor.Supports(ExcelProviderFeatures.PivotTables));
            Assert.True(descriptor.Supports(ExcelProviderFeatures.TemplatePivotTables));
        }
        else
        {
            Assert.False(descriptor.Supports(ExcelProviderFeatures.PivotTables));
            Assert.False(descriptor.Supports(ExcelProviderFeatures.TemplatePivotTables));
        }
    }
}
