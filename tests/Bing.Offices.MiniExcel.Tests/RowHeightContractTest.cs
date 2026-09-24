using System.IO;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.MiniExcel.Tests;

/// <summary>
/// 验证 MiniExcel 对尚未支持的行高能力明确失败。
/// </summary>
public sealed class RowHeightContractTest
{
    /// <summary>
    /// 验证 MiniExcel 遇到行高要求时在写出前报告不支持能力。
    /// </summary>
    [Fact]
    public void Export_WithRowHeight_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new Row { Name = "A" } }, sheet => sheet
                .RowHeight(new ExcelRowHeightOptions { BodyHeight = 18 })));
        using var stream = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new MiniExcelExcelExporter().Export(request, stream));

        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, stream.Length);
    }

    /// <summary>
    /// 表示行高契约测试使用的数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置行名称。
        /// </summary>
        public string Name { get; set; }
    }
}
