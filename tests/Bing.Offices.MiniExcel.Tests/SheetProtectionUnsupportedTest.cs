using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Bing.Offices.MiniExcel.Tests;

/// <summary>
/// MiniExcel 工作表保护支持边界测试。
/// </summary>
public sealed class SheetProtectionUnsupportedTest
{
    /// <summary>
    /// 验证同步和异步入口在写入前拒绝保护请求。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_ShouldRejectProtectionBeforeWriting(bool asynchronous)
    {
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new Row() }, sheet => sheet.Protection(
                new ExcelSheetProtectionDefinition(Array.Empty<ExcelRangeDefinition>()))));
        var exporter = new MiniExcelExcelExporter();
        var original = new byte[] { 4, 5, 6 };
        using var output = new MemoryStream(original);
        output.Position = 1;

        if (asynchronous)
            await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(
                () => exporter.ExportAsync(request, output));
        else
            Assert.Throws<BingOfficesUnsupportedFeatureException>(
                () => exporter.Export(request, output));
        Assert.Equal(original, output.ToArray());
        Assert.Equal(1, output.Position);
    }

    /// <summary>
    /// 测试行。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }
}
