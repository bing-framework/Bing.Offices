using System;
using System.Globalization;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Dates;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Excel 与 CSV 共享日期解析合同测试。
/// </summary>
public sealed class ExcelDateParserTest
{
    /// <summary>
    /// 测试 - 默认日期输入仅接受 ISO yyyy-MM-dd，并返回无时区 DateTime。
    /// </summary>
    [Fact]
    public void TryParse_DefaultFormat_ShouldAcceptIsoAndRejectCultureDependentText()
    {
        // Arrange
        var cell = TextCell("2026-09-04");

        // Act
        var valid = ExcelDateParser.TryParse(cell, cell.Text, typeof(DateTime),
            CultureInfo.GetCultureInfo("zh-CN"), null, out var parsed);
        var invalid = ExcelDateParser.TryParse(TextCell("2026/09/04"), "2026/09/04",
            typeof(DateTime), CultureInfo.GetCultureInfo("en-US"), null, out _);

        // Assert
        Assert.True(valid);
        Assert.Equal(new DateTime(2026, 9, 4), parsed);
        Assert.Equal(DateTimeKind.Unspecified, ((DateTime)parsed).Kind);
        Assert.False(invalid);
    }

    /// <summary>
    /// 测试 - 默认 ISO 日期不应因当前区域性变化而改变解析结果。
    /// </summary>
    [Fact]
    public void TryParse_DefaultFormat_ShouldIgnoreCurrentCultureForAcceptedShape()
    {
        // Arrange
        var originalCulture = CultureInfo.CurrentCulture;
        var originalUICulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ar-SA");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("ar-SA");

            // Act
            var result = ExcelDateParser.TryParse(TextCell("2026-09-04"), "2026-09-04",
                typeof(DateTime?), CultureInfo.GetCultureInfo("fr-FR"), null, out var parsed);

            // Assert
            Assert.True(result);
            Assert.Equal(new DateTime(2026, 9, 4), parsed);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
            CultureInfo.CurrentUICulture = originalUICulture;
        }
    }

    /// <summary>
    /// 测试 - 显式输入格式应允许非默认日期文本，但不应接受非法闰日。
    /// </summary>
    [Fact]
    public void TryParse_ExplicitFormat_ShouldAcceptConfiguredShapeAndRejectInvalidDate()
    {
        // Arrange
        var attribute = new ExcelDateAttribute("yyyy/MM/dd");

        // Act
        var valid = ExcelDateParser.TryParse(TextCell("2026/09/04"), "2026/09/04",
            typeof(DateTime), CultureInfo.InvariantCulture, attribute, out var parsed);
        var invalid = ExcelDateParser.TryParse(TextCell("2026/02/29"), "2026/02/29",
            typeof(DateTime), CultureInfo.InvariantCulture, attribute, out _);

        // Assert
        Assert.True(valid);
        Assert.Equal(new DateTime(2026, 9, 4), parsed);
        Assert.False(invalid);
    }

    /// <summary>
    /// 测试 - DateTimeOffset 默认要求显式 offset，固定 offset 策略才允许无 offset 文本。
    /// </summary>
    [Fact]
    public void TryParse_DateTimeOffset_ShouldRequireExplicitOrConfiguredOffset()
    {
        // Arrange
        var text = "2026-09-04T12:30:00+08:00";
        var defaultAttribute = new ExcelDateAttribute();
        var fixedAttribute = new ExcelDateAttribute
        {
            OffsetPolicy = ExcelDateOffsetPolicy.UseFixedOffset,
            OffsetMinutes = 480
        };

        // Act
        var explicitResult = ExcelDateParser.TryParse(TextCell(text), text, typeof(DateTimeOffset),
            CultureInfo.InvariantCulture, defaultAttribute, out var explicitValue);
        var rejectedResult = ExcelDateParser.TryParse(TextCell("2026-09-04"), "2026-09-04",
            typeof(DateTimeOffset), CultureInfo.InvariantCulture, defaultAttribute, out _);
        var fixedResult = ExcelDateParser.TryParse(TextCell("2026-09-04"), "2026-09-04",
            typeof(DateTimeOffset), CultureInfo.InvariantCulture, fixedAttribute, out var fixedValue);

        // Assert
        Assert.True(explicitResult);
        Assert.Equal(TimeSpan.FromHours(8), ((DateTimeOffset)explicitValue).Offset);
        Assert.False(rejectedResult);
        Assert.True(fixedResult);
        Assert.Equal(TimeSpan.FromHours(8), ((DateTimeOffset)fixedValue).Offset);
    }

    /// <summary>
    /// 测试 - 原生日期值应优先于显示文本，且 serial 应区分 1900 与 1904 日期系统。
    /// </summary>
    [Fact]
    public void TryParse_NativeAndSerialValues_ShouldUseRawValueAndDateSystem()
    {
        // Arrange
        var nativeDate = new DateTime(2026, 9, 4, 12, 30, 0, DateTimeKind.Local);
        var nativeCell = new ExcelCellValue(nativeDate, "not-a-date-text", ExcelCellKind.DateTime);
        var serial1900 = new ExcelCellValue(61d, "61", ExcelCellKind.Number);
        var serial1904 = new ExcelCellValue(0d, "0", ExcelCellKind.Number, isDate1904: true);

        // Act
        var nativeResult = ExcelDateParser.TryParse(nativeCell, nativeCell.Text, typeof(DateTime),
            CultureInfo.InvariantCulture, null, out var nativeValue);
        var serial1900Result = ExcelDateParser.TryParse(serial1900, serial1900.Text, typeof(DateTime),
            CultureInfo.InvariantCulture, null, out var value1900);
        var serial1904Result = ExcelDateParser.TryParse(serial1904, serial1904.Text, typeof(DateTime),
            CultureInfo.InvariantCulture, null, out var value1904);

        // Assert
        Assert.True(nativeResult);
        Assert.Equal(new DateTime(2026, 9, 4, 12, 30, 0), nativeValue);
        Assert.Equal(DateTimeKind.Unspecified, ((DateTime)nativeValue).Kind);
        Assert.True(serial1900Result);
        Assert.Equal(new DateTime(1900, 3, 1), value1900);
        Assert.True(serial1904Result);
        Assert.Equal(new DateTime(1904, 1, 1), value1904);
    }

    /// <summary>
    /// 测试 - Workbook 时间校验应使用固定基准日期，避免服务器当前日期影响比较结果。
    /// </summary>
    [Fact]
    public void TryParseValidation_Time_ShouldUseStableDateAndSerialValue()
    {
        // Arrange
        var textCell = TextCell("12:30:00");
        var serialCell = new ExcelCellValue(0.5d, "0.5", ExcelCellKind.Number);

        // Act
        var textResult = ExcelDateParser.TryParseValidation(textCell, textCell.Text, true, false,
            out var textValue);
        var serialResult = ExcelDateParser.TryParseValidation(serialCell, serialCell.Text, true, false,
            out var serialValue);

        // Assert
        Assert.True(textResult);
        Assert.True(serialResult);
        Assert.Equal(DateTime.MinValue.Date, textValue.Date);
        Assert.Equal(DateTime.MinValue.Date, serialValue.Date);
        Assert.Equal(new TimeSpan(12, 30, 0), textValue.TimeOfDay);
        Assert.Equal(new TimeSpan(12, 0, 0), serialValue.TimeOfDay);
        Assert.Equal(DateTimeKind.Unspecified, textValue.Kind);
        Assert.Equal(DateTimeKind.Unspecified, serialValue.Kind);
    }

    private static ExcelCellValue TextCell(string value) =>
        new(value, value, ExcelCellKind.Text);
}
