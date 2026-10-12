using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.MiniExcel.Tests;

/// <summary>
/// 验证 MiniExcel 的级联字典能力边界。
/// </summary>
public sealed class CascadeListUnsupportedTest
{
    /// <summary>
    /// 验证同步和异步导出均在写入前拒绝级联模板。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_CascadeList_RejectsWithoutWriting(bool asynchronous)
    {
        var request = ExcelExport.Workbook(book => book
            .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
            {
                ["A"] = new[] { "B" }
            })
            .AddSheet("Data", new[] { new Row() }));
        using var output = new MemoryStream(new byte[] { 4, 5, 6 });
        output.Position = 1;
        var exporter = new MiniExcelExcelExporter();
        if (asynchronous)
            await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(
                () => exporter.ExportAsync(request, output));
        else
            Assert.Throws<BingOfficesUnsupportedFeatureException>(
                () => exporter.Export(request, output));
        Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray());
        Assert.Equal(1, output.Position);
    }

    /// <summary>
    /// 提供测试行。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }
}
