using Bing.Offices;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// FontExtensions 的职责级测试。
/// </summary>
public sealed class NpoiFontExtensionsTest
{
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetFontHeightInPoints_ShouldSetHeightAndReturnSameFont(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var font = workbook.CreateFont();

        var returned = FontExtensions.SetFontHeightInPoints(font, 12);

        Assert.Same(font, returned);
        Assert.Equal((short)12, font.FontHeightInPoints);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetColor_ShouldSetColorAndReturnSameFont(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var font = workbook.CreateFont();

        var returned = FontExtensions.SetColor(font, 10);

        Assert.Same(font, returned);
        Assert.Equal((short)10, font.Color);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetBoldWeight_ShouldUse700AsBoldThresholdAndReturnSameFont(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var font = workbook.CreateFont();

        Assert.Same(font, FontExtensions.SetBoldWeight(font, 699));
        Assert.False(font.IsBold);
        Assert.Same(font, FontExtensions.SetBoldWeight(font, 700));
        Assert.True(font.IsBold);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void DefaultFont_ShouldSetSongFontAndNinePointHeight(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var font = workbook.CreateFont();
        font.IsBold = true;

        var returned = FontExtensions.DefaultFont(font);

        Assert.Same(font, returned);
        Assert.Equal("宋体", font.FontName);
        Assert.Equal((short)9, font.FontHeightInPoints);
        Assert.True(font.IsBold);
    }

    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
}
