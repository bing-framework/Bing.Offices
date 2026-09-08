using System;
using System.Globalization;
using Bing.Offices;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// CellExtensions 的职责级测试。
/// </summary>
public sealed class NpoiCellExtensionsTest
{
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void GetStringValue_ShouldReadAllSupportedCellKinds(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var row = sheet.CreateRow(0);
        var stringCell = row.CreateCell(0);
        stringCell.SetCellValue("  text  ");
        var booleanCell = row.CreateCell(1);
        booleanCell.SetCellValue(true);
        var errorCell = row.CreateCell(2);
        errorCell.SetCellErrorValue(FormulaError.DIV0.Code);
        var numericCell = row.CreateCell(3);
        numericCell.SetCellValue(12.5d);
        var dateCell = row.CreateCell(4);
        var date = new DateTime(2026, 9, 7, 14, 15, 16);
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd hh:mm:ss");
        dateCell.SetCellValue(date);
        dateCell.CellStyle = dateStyle;
        var blankCell = row.CreateCell(5, CellType.Blank);
        var formulaCell = row.CreateCell(6);
        formulaCell.SetCellFormula("1+1");
        workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(formulaCell);

        Assert.Equal("text", CellExtensions.GetStringValue(stringCell));
        Assert.Equal("True", CellExtensions.GetStringValue(booleanCell));
        Assert.Equal(Convert.ToString(FormulaError.DIV0.Code, CultureInfo.InvariantCulture),
            CellExtensions.GetStringValue(errorCell));
        Assert.Equal("12.5", CellExtensions.GetStringValue(numericCell));
        Assert.Equal(date.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture),
            CellExtensions.GetStringValue(dateCell));
        Assert.Equal(string.Empty, CellExtensions.GetStringValue(blankCell));
        Assert.Equal("2", CellExtensions.GetStringValue(formulaCell));
        Assert.Equal(string.Empty, CellExtensions.GetStringValue(null));
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetValue_ShouldWriteNullSpecialTypesAndPictures(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var row = sheet.CreateRow(0);

        var nullCell = row.CreateCell(0);
        CellExtensions.SetValue(nullCell, null);
        Assert.Equal(CellType.Blank, nullCell.CellType);

        var dbNullCell = row.CreateCell(1);
        CellExtensions.SetValue(dbNullCell, DBNull.Value);
        Assert.Equal(CellType.String, dbNullCell.CellType);
        Assert.Equal(string.Empty, dbNullCell.StringCellValue);

        var date = new DateTime(2026, 9, 7, 14, 15, 16);
        var dateCell = row.CreateCell(2);
        CellExtensions.SetValue(dateCell, date);
        Assert.Equal(CellType.Numeric, dateCell.CellType);
        Assert.True(DateUtil.IsCellDateFormatted(dateCell));
        Assert.Equal(date, dateCell.DateCellValue);
        Assert.Equal("yyyy-mm-dd hh:mm:ss", dateCell.CellStyle.GetDataFormatString());

        var enumCell = row.CreateCell(3);
        CellExtensions.SetValue(enumCell, CellStatus.Enabled);
        Assert.Equal(CellType.String, enumCell.CellType);
        Assert.Equal(nameof(CellStatus.Enabled), enumCell.StringCellValue);

        var guid = Guid.Parse("de305d54-75b4-431b-adb2-eb6b9e546014");
        var guidCell = row.CreateCell(4);
        CellExtensions.SetValue(guidCell, guid);
        Assert.Equal(guid.ToString("D"), guidCell.StringCellValue);

        var versionCell = row.CreateCell(5);
        var version = new Version(1, 2, 3, 4);
        CellExtensions.SetValue(versionCell, version);
        Assert.Equal(version.ToString(), versionCell.StringCellValue);

        var offsetCell = row.CreateCell(6);
        var offset = new DateTimeOffset(2026, 9, 7, 14, 15, 16, TimeSpan.FromHours(8));
        CellExtensions.SetValue(offsetCell, offset);
        Assert.Equal(offset.ToString("O", CultureInfo.InvariantCulture), offsetCell.StringCellValue);

        var pictureCell = row.CreateCell(7);
        CellExtensions.SetValue(pictureCell, Png);
        Assert.Single(workbook.GetAllPictures());
        Assert.NotNull(sheet.DrawingPatriarch);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetValue_ShouldPreserveExcelSafeNumericPrecisionAndApplyScale(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);
        const long maxSafeInteger = 9007199254740991;
        const long unsafeLong = 9007199254740992;
        const ulong unsafeUnsignedLong = 9007199254740992;
        var unsafeDecimal = 9007199254740991.1m;

        var safeLongCell = row.CreateCell(0);
        CellExtensions.SetValue(safeLongCell, maxSafeInteger);
        Assert.Equal(CellType.Numeric, safeLongCell.CellType);
        Assert.Equal((double)maxSafeInteger, safeLongCell.NumericCellValue);

        var unsafeLongCell = row.CreateCell(1);
        CellExtensions.SetValue(unsafeLongCell, unsafeLong);
        Assert.Equal(CellType.String, unsafeLongCell.CellType);
        Assert.Equal(unsafeLong.ToString(CultureInfo.InvariantCulture), unsafeLongCell.StringCellValue);

        var unsafeUnsignedLongCell = row.CreateCell(2);
        CellExtensions.SetValue(unsafeUnsignedLongCell, unsafeUnsignedLong);
        Assert.Equal(CellType.String, unsafeUnsignedLongCell.CellType);
        Assert.Equal(unsafeUnsignedLong.ToString(CultureInfo.InvariantCulture), unsafeUnsignedLongCell.StringCellValue);

        var safeDecimalCell = row.CreateCell(3);
        CellExtensions.SetValue(safeDecimalCell, 12.5m);
        Assert.Equal(CellType.Numeric, safeDecimalCell.CellType);
        Assert.Equal(12.5d, safeDecimalCell.NumericCellValue, 10);

        var unsafeDecimalCell = row.CreateCell(4);
        CellExtensions.SetValue(unsafeDecimalCell, unsafeDecimal);
        Assert.Equal(CellType.String, unsafeDecimalCell.CellType);
        Assert.Equal(unsafeDecimal.ToString(CultureInfo.InvariantCulture), unsafeDecimalCell.StringCellValue);

        var scaledCell = row.CreateCell(5);
        CellExtensions.SetValue(scaledCell, 12.3456m, 2);
        Assert.Equal(CellType.Numeric, scaledCell.CellType);
        Assert.Equal(12.35d, scaledCell.NumericCellValue, 10);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetCellValue_Overloads_ShouldWriteFormattedDatesNumbersTextAndBlanks(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var row = workbook.CreateSheet("Data").CreateRow(0);
        var date = new DateTime(2026, 9, 7, 14, 15, 16);
        var offset = new DateTimeOffset(2026, 9, 7, 14, 15, 16, TimeSpan.FromHours(8));

        var directCell = row.CreateCell(0);
        CellExtensions.SetCellValue(directCell, "direct overload");
        Assert.Equal(CellType.String, directCell.CellType);
        Assert.Equal("direct overload", directCell.StringCellValue);

        var dateCell = row.CreateCell(1);
        CellExtensions.SetCellValue(dateCell, date, "dd/mm/yyyy hh:mm");
        Assert.Equal(CellType.Numeric, dateCell.CellType);
        Assert.Equal(date, dateCell.DateCellValue);
        Assert.Equal("dd/mm/yyyy hh:mm", dateCell.CellStyle.GetDataFormatString());

        var numberCell = row.CreateCell(2);
        CellExtensions.SetCellValue(numberCell, 12.5m, "0.000");
        Assert.Equal(CellType.Numeric, numberCell.CellType);
        Assert.Equal(12.5d, numberCell.NumericCellValue, 10);
        Assert.Equal("0.000", numberCell.CellStyle.GetDataFormatString());

        var formattedTextCell = row.CreateCell(3);
        CellExtensions.SetCellValue(formattedTextCell, offset, "yyyy");
        Assert.Equal(CellType.String, formattedTextCell.CellType);
        Assert.Equal("2026", formattedTextCell.StringCellValue);

        var nullCell = row.CreateCell(4);
        CellExtensions.SetCellValue(nullCell, null, "0.00");
        Assert.Equal(CellType.Blank, nullCell.CellType);

        var dbNullCell = row.CreateCell(5);
        CellExtensions.SetCellValue(dbNullCell, DBNull.Value, "0.00");
        Assert.Equal(CellType.Blank, dbNullCell.CellType);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void StyleHelpers_ShouldCacheDerivedStylesWithoutMutatingBaseStyle(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var baseStyle = workbook.CreateCellStyle();
        baseStyle.VerticalAlignment = VerticalAlignment.Top;
        var first = sheet.CreateRow(0).CreateCell(0);
        var second = sheet.CreateRow(1).CreateCell(0);
        first.CellStyle = baseStyle;
        second.CellStyle = baseStyle;
        var stylesBefore = workbook.NumCellStyles;

        var firstWrap = CellExtensions.GetStyleWithWrapText(first);
        var stylesAfterFirstWrap = workbook.NumCellStyles;
        var secondWrap = CellExtensions.GetStyleWithWrapText(second);

        Assert.True(firstWrap.WrapText);
        Assert.False(baseStyle.WrapText);
        Assert.Equal(firstWrap.Index, secondWrap.Index);
        Assert.Equal(stylesAfterFirstWrap, workbook.NumCellStyles);
        Assert.True(workbook.NumCellStyles > stylesBefore);

        var firstVertical = CellExtensions.GetStyleWithVerticalAlignment(first, VerticalAlignment.Bottom);
        var stylesAfterFirstVertical = workbook.NumCellStyles;
        var secondVertical = CellExtensions.GetStyleWithVerticalAlignment(second, VerticalAlignment.Bottom);
        Assert.Equal(VerticalAlignment.Bottom, firstVertical.VerticalAlignment);
        Assert.Equal(firstVertical.Index, secondVertical.Index);
        Assert.Equal(stylesAfterFirstVertical, workbook.NumCellStyles);
        Assert.True(workbook.NumCellStyles > stylesAfterFirstWrap);

        first.CellStyle = firstWrap;
        Assert.Equal(firstWrap.Index,
            CellExtensions.GetStyleWithWrapText(first).Index);
        first.CellStyle = firstVertical;
        Assert.Equal(firstVertical.Index,
            CellExtensions.GetStyleWithVerticalAlignment(first, VerticalAlignment.Bottom).Index);

        var firstDate = sheet.CreateRow(2).CreateCell(0);
        var secondDate = sheet.CreateRow(3).CreateCell(0);
        var dateStylesBefore = workbook.NumCellStyles;
        CellExtensions.SetValue(firstDate, new DateTime(2026, 9, 7));
        var dateStylesAfterFirst = workbook.NumCellStyles;
        CellExtensions.SetValue(secondDate, new DateTime(2026, 9, 7));
        Assert.Equal(firstDate.CellStyle.Index, secondDate.CellStyle.Index);
        Assert.Equal(dateStylesAfterFirst, workbook.NumCellStyles);
        Assert.True(workbook.NumCellStyles > dateStylesBefore);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ConditionalFormatting_ShouldReturnRulesForCoveredCellOnly(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var covered = sheet.CreateRow(0).CreateCell(0);
        var outside = sheet.GetRow(0).CreateCell(1);
        var formatting = sheet.SheetConditionalFormatting;
        var equalRule = formatting.CreateConditionalFormattingRule(ComparisonOperator.Equal, "1");
        var greaterRule = formatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan, "0");

        CellExtensions.AddConditionalFormattingRules(covered, new[] { equalRule, greaterRule });

        var coveredRules = CellExtensions.GetConditionalFormattingRules(covered);
        var outsideRules = CellExtensions.GetConditionalFormattingRules(outside);

        Assert.Equal(2, coveredRules.Length);
        Assert.Equal(equalRule.Formula1, coveredRules[0].Formula1);
        Assert.Equal(greaterRule.Formula1, coveredRules[1].Formula1);
        Assert.Empty(outsideRules);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Merge_ShouldMergeSameSheetAndSupportExpandAndNonExpandModes(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var sheet = workbook.CreateSheet("Data");
        var from = sheet.CreateRow(0).CreateCell(0);
        var to = sheet.CreateRow(2).CreateCell(2);

        CellExtensions.Merge(from, to);

        var firstRegion = sheet.GetMergedRegion(0);
        Assert.Equal(0, firstRegion.FirstRow);
        Assert.Equal(2, firstRegion.LastRow);
        Assert.Equal(0, firstRegion.FirstColumn);
        Assert.Equal(2, firstRegion.LastColumn);

        var expandSheet = workbook.CreateSheet("Expand");
        expandSheet.CreateRow(1).CreateCell(1);
        expandSheet.CreateRow(3).CreateCell(3);
        expandSheet.AddMergedRegion(new CellRangeAddress(0, 2, 2, 4));
        CellExtensions.Merge(expandSheet.GetRow(1).GetCell(1), expandSheet.GetRow(3).GetCell(3), true);

        var expandedRegion = expandSheet.GetMergedRegion(0);
        Assert.Equal(0, expandedRegion.FirstRow);
        Assert.Equal(3, expandedRegion.LastRow);
        Assert.Equal(1, expandedRegion.FirstColumn);
        Assert.Equal(4, expandedRegion.LastColumn);

        var nonExpandSheet = workbook.CreateSheet("NonExpand");
        nonExpandSheet.CreateRow(1).CreateCell(1);
        nonExpandSheet.CreateRow(3).CreateCell(3);
        nonExpandSheet.AddMergedRegion(new CellRangeAddress(0, 2, 2, 4));
        CellExtensions.Merge(nonExpandSheet.GetRow(1).GetCell(1), nonExpandSheet.GetRow(3).GetCell(3), false);

        var nonExpandedRegion = nonExpandSheet.GetMergedRegion(0);
        Assert.Equal(1, nonExpandedRegion.FirstRow);
        Assert.Equal(3, nonExpandedRegion.LastRow);
        Assert.Equal(1, nonExpandedRegion.FirstColumn);
        Assert.Equal(3, nonExpandedRegion.LastColumn);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Merge_CrossSheetCells_ShouldThrowInvalidOperationException(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        var first = workbook.CreateSheet("First").CreateRow(0).CreateCell(0);
        var second = workbook.CreateSheet("Second").CreateRow(0).CreateCell(0);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            CellExtensions.Merge(first, second));

        Assert.Equal("单元格不在同一个工作表上", exception.Message);
    }

    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void SetValueAndFormattedSetCellValue_NullCell_ShouldThrowArgumentNullException(ExcelFormat format)
    {
        using var workbook = CreateWorkbook(format);
        Assert.Equal(format, WorkbookExtensions.GetExcelFormat(workbook));

        var setValueException = Assert.Throws<ArgumentNullException>(() =>
            CellExtensions.SetValue(null, "value"));
        var formattedException = Assert.Throws<ArgumentNullException>(() =>
            CellExtensions.SetCellValue(null, "value", "0.00"));

        Assert.Equal("cell", setValueException.ParamName);
        Assert.Equal("cell", formattedException.ParamName);
    }

    private static IWorkbook CreateWorkbook(ExcelFormat format) =>
        format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();

    private static readonly byte[] Png =
    {
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 82,
        0, 0, 0, 1, 0, 0, 0, 1, 8, 6, 0, 0, 0, 31, 21, 196, 137,
        0, 0, 0, 13, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 240,
        31, 0, 5, 0, 1, 255, 137, 153, 61, 29, 0, 0, 0, 0, 73, 69,
        78, 68, 174, 66, 96, 130
    };

    private enum CellStatus
    {
        Disabled,
        Enabled
    }
}
