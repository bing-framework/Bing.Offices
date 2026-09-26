using System;
using System.IO;
using System.Text;
using Bing.Offices.AsposeCells;
using Bing.Offices.Exceptions;
using Bing.Offices.Rendering;
using Xunit;

namespace Bing.Offices.AsposeCells.Tests;

/// <summary>
/// Aspose.Cells 可选商业 Provider 的许可证与能力边界合同。
/// </summary>
public sealed class AsposeCellsProviderContractTest
{
    /// <summary>
    /// 无许可证时在预检阶段拒绝，且不写入调用方目标流。
    /// </summary>
    [Fact]
    public void RenderWithoutLicense_ShouldRejectBeforeWritingDestination()
    {
        var engine = new AsposeCellsEngine();
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("not-a-workbook"));
        using var destination = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
        {
            engine.Render(source, destination, new ExcelRenderRequest
            {
                InputFormat = Bing.Offices.ExcelFormat.Xlsx,
                ImageFormat = ExcelRenderFormat.Pdf
            });
        });

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, destination.Length);
        Assert.DoesNotContain("password", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 能力声明区分工作簿编辑、渲染、格式和加密方向。
    /// </summary>
    [Fact]
    public void FeatureDescriptor_ShouldDeclareCommercialBoundaries()
    {
        var engine = new AsposeCellsEngine();

        Assert.Equal("Aspose.Cells", engine.ProviderName);
        Assert.True(engine.Supports(Bing.Offices.Providers.ExcelProviderFeatures.PdfRendering
            | Bing.Offices.Providers.ExcelProviderFeatures.PageImageRendering));
        Assert.Contains(Bing.Offices.ExcelFormat.Xlsm, engine.ReadFormats);
        Assert.Contains(Bing.Offices.ExcelFormat.Ods, engine.WriteFormats);
        Assert.False(engine.SupportsCompleteWorkbookImport);
    }
}
