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
    /// <summary>
    /// 验证设置字体磅值会设置高度并返回相同字体。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
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

    /// <summary>
    /// 验证设置颜色会设置颜色并返回相同字体。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
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

    /// <summary>
    /// 验证设置粗体权重会按阈值处理并返回相同字体。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
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

    /// <summary>
    /// 验证默认字体会设置宋体和九磅高度。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
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

    /// <summary>
    /// 创建测试工作簿。
    /// </summary>
    /// <param name="format">目标文件格式。</param>
    /// <returns>生成的 IWorkbook 结果。</returns>
    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
}
