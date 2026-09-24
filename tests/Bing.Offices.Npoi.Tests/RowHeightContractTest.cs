using System.IO;
using Bing.Offices.Exports;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Npoi.Tests;

/// <summary>
/// 验证 NPOI 对公共行高契约的真实 XLSX 写入。
/// </summary>
public sealed class RowHeightContractTest
{
    /// <summary>
    /// 验证 NPOI 将表头和数据行高度写入真实 XLSX。
    /// </summary>
    [Fact]
    public void Export_ShouldApplyHeaderAndBodyRowHeights()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new Row { Name = "A" } }, sheet => sheet
                .RowHeight(new ExcelRowHeightOptions { HeaderHeight = 22, BodyHeight = 18 })));
        using var stream = new MemoryStream();

        new NpoiExcelExporter().Export(request, stream);

        stream.Position = 0;
        using var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheet("Data");
        Assert.Equal(22, sheet.GetRow(0).HeightInPoints);
        Assert.Equal(18, sheet.GetRow(1).HeightInPoints);
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
