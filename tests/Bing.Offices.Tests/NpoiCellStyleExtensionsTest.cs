using System;
using Bing.Offices;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// CellStyleExtensions 的职责级测试。
/// </summary>
public sealed class NpoiCellStyleExtensionsTest
{
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Setters_ShouldSetPropertiesAndReturnSameStyle(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var style = workbook.CreateCellStyle();

        Assert.Same(style, CellStyleExtensions.SetDataFormat(style, 14));
        Assert.Equal((short)14, style.DataFormat);

        Assert.Same(style, CellStyleExtensions.SetHorizontalAlignment(style, HorizontalAlignment.Right));
        Assert.Equal(HorizontalAlignment.Right, style.Alignment);

        Assert.Same(style, CellStyleExtensions.SetVerticalAlignment(style, VerticalAlignment.Bottom));
        Assert.Equal(VerticalAlignment.Bottom, style.VerticalAlignment);

        Assert.Same(style, CellStyleExtensions.SetFillForegroundColor(style, 13));
        Assert.Equal((short)13, style.FillForegroundColor);

        Assert.Same(style, CellStyleExtensions.SetFillBackgroundColor(style, 12));
        Assert.Equal((short)12, style.FillBackgroundColor);

        Assert.Same(style, CellStyleExtensions.SetFillPattern(style, FillPattern.SolidForeground));
        Assert.Equal(FillPattern.SolidForeground, style.FillPattern);

        Assert.Same(style, CellStyleExtensions.SetBorderTop(style, BorderStyle.Medium));
        Assert.Equal(BorderStyle.Medium, style.BorderTop);
        Assert.Same(style, CellStyleExtensions.SetBorderRight(style, BorderStyle.Dashed));
        Assert.Equal(BorderStyle.Dashed, style.BorderRight);
        Assert.Same(style, CellStyleExtensions.SetBorderBottom(style, BorderStyle.Double));
        Assert.Equal(BorderStyle.Double, style.BorderBottom);
        Assert.Same(style, CellStyleExtensions.SetBorderLeft(style, BorderStyle.Thin));
        Assert.Equal(BorderStyle.Thin, style.BorderLeft);

        Assert.Same(style, CellStyleExtensions.SetBorder(style, BorderStyle.Hair));
        Assert.Equal(BorderStyle.Hair, style.BorderTop);
        Assert.Equal(BorderStyle.Hair, style.BorderRight);
        Assert.Equal(BorderStyle.Hair, style.BorderBottom);
        Assert.Equal(BorderStyle.Hair, style.BorderLeft);

        Assert.Same(style, CellStyleExtensions.SetBorderTopColor(style, 10));
        Assert.Equal((short)10, style.TopBorderColor);
        Assert.Same(style, CellStyleExtensions.SetBorderRightColor(style, 11));
        Assert.Equal((short)11, style.RightBorderColor);
        Assert.Same(style, CellStyleExtensions.SetBorderBottomColor(style, 12));
        Assert.Equal((short)12, style.BottomBorderColor);
        Assert.Same(style, CellStyleExtensions.SetBorderLeftColor(style, 13));
        Assert.Equal((short)13, style.LeftBorderColor);

        Assert.Same(style, CellStyleExtensions.SetBorderColor(style, 14));
        Assert.Equal((short)14, style.TopBorderColor);
        Assert.Equal((short)14, style.RightBorderColor);
        Assert.Equal((short)14, style.BottomBorderColor);
        Assert.Equal((short)14, style.LeftBorderColor);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetFont_ShouldCreateConfigureBindAndReturnSameStyle(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var style = workbook.CreateCellStyle();

        var returned = CellStyleExtensions.SetFont(style, workbook, font =>
        {
            font.FontName = "Arial";
            font.FontHeightInPoints = 11;
            font.IsItalic = true;
        });

        Assert.Same(style, returned);
        var font = workbook.GetFontAt(style.FontIndex);
        Assert.Equal("Arial", font.FontName);
        Assert.Equal((short)11, font.FontHeightInPoints);
        Assert.True(font.IsItalic);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetFont_NullAction_ShouldBindDefaultCreatedFont(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var style = workbook.CreateCellStyle();

        var returned = CellStyleExtensions.SetFont(style, workbook, null);

        Assert.Same(style, returned);
        Assert.NotNull(workbook.GetFontAt(style.FontIndex));
    }

    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
}
