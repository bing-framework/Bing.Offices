using System.IO;
using System;
using System.Linq;
using System.IO.Compression;
using System.Text;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Bing.Offices.Providers;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 公共报表模型的直接结构测试。
/// </summary>
public sealed class NpoiReportContractTest
{
    /// <summary>
    /// 验证 XLSX 表格、筛选、冻结、名称范围、条件格式和打印布局。
    /// </summary>
    [Fact]
    public void Export_CommonReportDefinitions_ShouldWriteWorkbookStructure()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new ReportRow { Name = "A", Amount = 1 }, new ReportRow { Name = "B", Amount = 2 } },
            sheet => sheet
                .Table(new ExcelTableDefinition
                {
                    Name = "RowsTable",
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 },
                    ShowTotals = true
                })
                .AutoFilter(new ExcelAutoFilterDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 }
                })
                .FreezePane(new ExcelFreezePaneDefinition { Rows = 1 })
                .ConditionalFormat(new ExcelConditionalFormatDefinition
                {
                    Type = ExcelConditionalFormatType.CellValue,
                    Range = new ExcelRangeDefinition { StartRow = 1, StartColumn = 1, EndRow = 2, EndColumn = 1 },
                    Operator = ExcelConditionalComparisonOperator.GreaterThan,
                    Formula1 = "1"
                })
                .NamedRange(new ExcelNamedRangeDefinition { Name = "RowsAmount", Address = "$B$2:$B$3" })
                .PrintLayout(new ExcelPrintLayoutOptions
                {
                    Orientation = ExcelPrintOrientation.Landscape,
                    PaperSize = ExcelPrintPaperSize.A4,
                    PrintArea = "$A$1:$B$3",
                    RepeatRows = "$1:$1"
                })));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var sheet = Assert.IsType<XSSFSheet>(workbook.GetSheet("Rows"));

        var table = Assert.Single(sheet.GetTables());
        Assert.Equal("RowsTable", table.Name);
        Assert.Equal("A1:B4", table.CellReferences.FormatAsString());
        Assert.True(table.IsHasTotalsRow);
        Assert.Equal(1u, table.GetCTTable().totalsRowCount);
        Assert.Equal(2d, sheet.GetRow(2).GetCell(1).NumericCellValue);
        Assert.NotNull(sheet.GetRow(3));
        Assert.Null(sheet.GetRow(3).GetCell(1));
        Assert.Null(sheet.GetCTWorksheet().autoFilter);
        Assert.NotNull(table.GetCTTable().autoFilter);
        Assert.Equal("A1:B3", table.GetCTTable().autoFilter.@ref);
        Assert.NotNull(sheet.PaneInformation);
        Assert.Contains(workbook.GetAllNames(), item => item.NameName == "RowsAmount");
        Assert.True(sheet.SheetConditionalFormatting.NumConditionalFormattings > 0);
        Assert.Equal((short)PaperSize.A4, sheet.PrintSetup.PaperSize);
        Assert.True(sheet.PrintSetup.Landscape);
    }

    /// <summary>
    /// 验证 XLSX 高级条件格式使用确定的最小值、最大值和 33/67 阈值。
    /// </summary>
    [Fact]
    public void Export_AdvancedConditionalFormats_ShouldWriteExpectedXlsxRules()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new ReportRow { Name = "A", Amount = 1 }, new ReportRow { Name = "B", Amount = 2 } },
            sheet =>
            {
                var range = new ExcelRangeDefinition { StartRow = 1, StartColumn = 1, EndRow = 2, EndColumn = 1 };
                sheet.ConditionalFormat(new ExcelConditionalFormatDefinition
                    { Type = ExcelConditionalFormatType.ColorScale, Range = range,
                        ForegroundColor = "#FF0000", BackgroundColor = "#00FF00" });
                sheet.ConditionalFormat(new ExcelConditionalFormatDefinition
                    { Type = ExcelConditionalFormatType.DataBar, Range = range, ForegroundColor = "#0000FF" });
                sheet.ConditionalFormat(new ExcelConditionalFormatDefinition
                    { Type = ExcelConditionalFormatType.IconSet, Range = range });
            }));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        var xml = ReadWorksheetXml(output);

        Assert.Contains("<colorScale>", xml);
        Assert.Contains("type=\"min\"", xml);
        Assert.Contains("type=\"max\"", xml);
        Assert.Contains("<dataBar", xml);
        Assert.Contains("<iconSet>", xml);
        Assert.Contains("type=\"percent\" val=\"33\"", xml);
        Assert.Contains("type=\"percent\" val=\"67\"", xml);
    }

    /// <summary>
    /// 验证 XLS/HSSF 在读取数据前拒绝高级条件格式。
    /// </summary>
    /// <param name="type">待验证的条件格式类型。</param>
    [Theory]
    [InlineData(ExcelConditionalFormatType.ColorScale)]
    [InlineData(ExcelConditionalFormatType.DataBar)]
    [InlineData(ExcelConditionalFormatType.IconSet)]
    public void Export_XlsAdvancedConditionalFormat_ShouldFailBeforeEnumeration(
        ExcelConditionalFormatType type)
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls).AddSheet("Rows", data,
            sheet => sheet.ConditionalFormat(new ExcelConditionalFormatDefinition
            {
                Type = type,
                Range = new ExcelRangeDefinition { StartRow = 1, StartColumn = 1, EndRow = 2, EndColumn = 1 },
                ForegroundColor = type == ExcelConditionalFormatType.IconSet ? null : "#FF0000",
                BackgroundColor = type == ExcelConditionalFormatType.ColorScale ? "#00FF00" : null
            })));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证 XLS 的普通条件格式在同步和异步出口均使用 HSSF 调色板颜色。
    /// </summary>
    /// <param name="type">待验证的条件格式类型。</param>
    /// <param name="hasForeground">是否设置前景色。</param>
    /// <param name="hasBackground">是否设置背景色。</param>
    [Theory]
    [InlineData(ExcelConditionalFormatType.CellValue, false, false)]
    [InlineData(ExcelConditionalFormatType.CellValue, true, false)]
    [InlineData(ExcelConditionalFormatType.CellValue, false, true)]
    [InlineData(ExcelConditionalFormatType.CellValue, true, true)]
    [InlineData(ExcelConditionalFormatType.Formula, false, false)]
    [InlineData(ExcelConditionalFormatType.Formula, true, false)]
    [InlineData(ExcelConditionalFormatType.Formula, false, true)]
    [InlineData(ExcelConditionalFormatType.Formula, true, true)]
    public async System.Threading.Tasks.Task Export_XlsBasicConditionalFormat_ShouldRoundTripPaletteColors(
        ExcelConditionalFormatType type, bool hasForeground, bool hasBackground)
    {
        foreach (var useAsync in new[] { false, true })
        {
            var definition = new ExcelConditionalFormatDefinition
            {
                Type = type,
                Range = new ExcelRangeDefinition { StartRow = 1, StartColumn = 1, EndRow = 1, EndColumn = 1 },
                Operator = ExcelConditionalComparisonOperator.GreaterThan,
                Formula1 = type == ExcelConditionalFormatType.Formula ? "MOD(ROW(),2)=0" : "0",
                ForegroundColor = hasForeground ? "#FF0000" : null,
                BackgroundColor = hasBackground ? "#00FF00" : null
            };
            var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls).AddSheet("Rows",
                new[] { new ReportRow { Name = "A", Amount = 1 } },
                sheet => sheet.ConditionalFormat(definition)));

            using var output = new MemoryStream();
            var exporter = new NpoiExcelExporter();
            if (useAsync)
                await exporter.ExportAsync(request, output);
            else
                exporter.Export(request, output);

            output.Position = 0;
            using var workbook = WorkbookFactory.Create(output);
            var formatting = workbook.GetSheet("Rows").SheetConditionalFormatting;
            Assert.Equal(1, formatting.NumConditionalFormattings);
            var rule = formatting.GetConditionalFormattingAt(0).GetRule(0);
            if (hasForeground)
                Assert.Equal(new byte[] { 255, 0, 0 }, rule.FontFormatting.FontColor.RGB);
            if (hasBackground)
            {
                Assert.Equal(FillPattern.SolidForeground, rule.PatternFormatting.FillPattern);
                Assert.Equal(new byte[] { 0, 255, 0 }, rule.PatternFormatting.FillForegroundColorColor.RGB);
            }
        }
    }

    /// <summary>
    /// 验证 XLS/HSSF 在读取数据前拒绝表格定义。
    /// </summary>
    [Fact]
    public void Export_XlsTable_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls).AddSheet("Rows", data,
            sheet => sheet.Table(new ExcelTableDefinition
            {
                Name = "RowsTable",
                Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 }
            })));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证汇总行扩展超出 XLSX 末行时在枚举前拒绝。
    /// </summary>
    [Fact]
    public void Export_TableTotalsBeyondLastRow_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", data,
            sheet => sheet.Table(new ExcelTableDefinition
            {
                Name = "RowsTable",
                Range = new ExcelRangeDefinition
                {
                    StartRow = 1048575, StartColumn = 0, EndRow = 1048575, EndColumn = 1
                },
                ShowTotals = true
            })));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证打印缩放冲突在枚举前拒绝。
    /// </summary>
    [Fact]
    public void Export_PrintScaleConflict_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", data,
            sheet => sheet.PrintLayout(new ExcelPrintLayoutOptions { ScalePercent = 100, FitToWidth = 1 })));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证 XLS/XLSX 保存重开后 Letter、缩放和重复标题范围保真。
    /// </summary>
    /// <param name="format">待验证的工作簿格式。</param>
    /// <param name="scalePercent">待验证的打印缩放百分比。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, 10)]
    [InlineData(ExcelFormat.Xls, 400)]
    [InlineData(ExcelFormat.Xlsx, 10)]
    [InlineData(ExcelFormat.Xlsx, 400)]
    public void Export_PrintLayout_ShouldRoundTrip(ExcelFormat format, int scalePercent)
    {
        var request = ExcelExport.Workbook(workbook => workbook.Format(format).AddSheet("Rows",
            new[] { new ReportRow { Name = "A", Amount = 1 } },
            sheet => sheet.PrintLayout(new ExcelPrintLayoutOptions
            {
                PaperSize = ExcelPrintPaperSize.Letter,
                ScalePercent = (short)scalePercent,
                PrintArea = "$A$1:$B$2",
                RepeatRows = "$1:$1",
                RepeatColumns = "$A:$A"
            })));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var sheet = workbook.GetSheet("Rows");

        Assert.Equal((short)PaperSize.US_Letter_Small, sheet.PrintSetup.PaperSize);
        Assert.Equal((short)scalePercent, sheet.PrintSetup.Scale);
        Assert.Contains("$A$1:$B$2", workbook.GetPrintArea(workbook.GetSheetIndex(sheet)));
        var printTitles = workbook.GetAllNames()
            .Where(name => name.NameName.EndsWith("Print_Titles", StringComparison.Ordinal))
            .Select(name => name.RefersToFormula)
            .ToArray();
        Assert.Contains(printTitles, formula => formula.Contains("$1:$1", StringComparison.Ordinal)
            || formula.Contains("$A$1:$IV$1", StringComparison.Ordinal));
        Assert.Contains(printTitles, formula => formula.Contains("$A:$A", StringComparison.Ordinal)
            || formula.EndsWith("A:A", StringComparison.Ordinal));
    }

    /// <summary>
    /// 验证底层格式不接受的打印缩放在枚举前拒绝。
    /// </summary>
    [Fact]
    public void Export_PrintScaleBelowProviderMinimum_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", data,
            sheet => sheet.PrintLayout(new ExcelPrintLayoutOptions { ScalePercent = 9 })));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证重复工作簿名称在枚举前拒绝。
    /// </summary>
    [Fact]
    public void Export_DuplicateWorkbookName_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.AddSheet("First", data, sheet => sheet
                .NamedRange(new ExcelNamedRangeDefinition { Name = "GlobalName", Address = "$A$1" }));
            workbook.AddSheet("Second", new[] { new ReportRow() }, sheet => sheet
                .NamedRange(new ExcelNamedRangeDefinition { Name = "GlobalName", Address = "$A$1" }));
        });

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证全局名称与各 Sheet 的同名局部名称可同时保存重开。
    /// </summary>
    /// <param name="format">待验证的工作簿格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_SameNamedWorkbookAndSheetRanges_ShouldPreserveScopes(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.Format(format);
            workbook.AddSheet("First", new[] { new ReportRow { Name = "A", Amount = 1 } }, sheet => sheet
                .NamedRange(new ExcelNamedRangeDefinition { Name = "SharedName", Address = "$A$1" })
                .NamedRange(new ExcelNamedRangeDefinition
                    { Name = "SharedName", SheetName = "First", Address = "$A$1" }));
            workbook.AddSheet("Second", new[] { new ReportRow { Name = "B", Amount = 2 } }, sheet => sheet
                .NamedRange(new ExcelNamedRangeDefinition
                    { Name = "SharedName", SheetName = "Second", Address = "$A$1" }));
        });

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);

        var scopes = workbook.GetAllNames().Where(name => name.NameName == "SharedName")
            .Select(name => name.SheetIndex).OrderBy(index => index).ToArray();
        Assert.Equal(new[] { -1, 0, 1 }, scopes);
    }

    /// <summary>
    /// 验证 XLSX 区域边界在枚举前检查。
    /// </summary>
    [Fact]
    public void Export_OutOfBoundsRange_ShouldFailBeforeEnumeration()
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", data,
            sheet => sheet.AutoFilter(new ExcelAutoFilterDefinition
            {
                Range = new ExcelRangeDefinition
                    { StartRow = 0, StartColumn = 0, EndRow = 0, EndColumn = 16384 }
            })));

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().Export(request, new MemoryStream()));
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 读取导出工作簿第一个工作表的 XML。
    /// </summary>
    /// <param name="output">包含导出工作簿的内存流。</param>
    /// <returns>第一个工作表的完整 XML 文本。</returns>
    private static string ReadWorksheetXml(MemoryStream output)
    {
        output.Position = 0;
        using var archive = new ZipArchive(output, ZipArchiveMode.Read, true);
        using var reader = new StreamReader(archive.GetEntry("xl/worksheets/sheet1.xml").Open(), Encoding.UTF8);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// 用于报表结构验证的数据行。
    /// </summary>
    private sealed class ReportRow
    {
        /// <summary>
        /// 获取或设置测试行名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置报表测试数值。
        /// </summary>
        public int Amount { get; set; }
    }

    /// <summary>
    /// 用于检测预检前数据枚举的测试序列。
    /// </summary>
    /// <typeparam name="T">测试序列的元素类型。</typeparam>
    private sealed class ThrowOnEnumeration<T> : System.Collections.Generic.IEnumerable<T>
    {
        /// <summary>
        /// 获取尝试枚举数据的次数。
        /// </summary>
        public int EnumerationCount { get; private set; }

        /// <inheritdoc />
        public System.Collections.Generic.IEnumerator<T> GetEnumerator()
        {
            EnumerationCount++;
            throw new InvalidOperationException("数据不应在预检失败前被枚举。");
        }

        /// <inheritdoc />
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
