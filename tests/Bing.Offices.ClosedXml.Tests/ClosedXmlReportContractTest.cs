using System.IO;
using System;
using System.IO.Compression;
using System.Text;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Bing.Offices.ClosedXml.Exports;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML 公共报表模型的直接结构测试。
/// </summary>
public sealed class ClosedXmlReportContractTest
{
    /// <summary>
    /// 验证表格、筛选、冻结、名称范围、条件格式和打印布局均写入真实 XLSX。
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
                    StyleName = "TableStyleMedium2",
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
                    Formula1 = "1",
                    ForegroundColor = "#FF0000"
                })
                .NamedRange(new ExcelNamedRangeDefinition { Name = "RowsAmount", Address = "$B$2:$B$3" })
                .PrintLayout(new ExcelPrintLayoutOptions
                {
                    Orientation = ExcelPrintOrientation.Landscape,
                    PaperSize = ExcelPrintPaperSize.A4,
                    PrintArea = "$A$1:$B$3",
                    RepeatRows = "1:1"
                })));

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("Rows");

        Assert.True(sheet.Tables.Contains("RowsTable"));
        var table = sheet.Tables.Table("RowsTable");
        Assert.True(table.ShowAutoFilter);
        Assert.True(table.ShowTotalsRow);
        Assert.Equal(2d, sheet.Cell("B3").GetDouble());
        Assert.Contains(workbook.DefinedNames.ValidNamedRanges(), item => item.Name == "RowsAmount");
        Assert.Equal(XLPageOrientation.Landscape, sheet.PageSetup.PageOrientation);
        Assert.Equal(XLPaperSize.A4Paper, sheet.PageSetup.PaperSize);
        Assert.NotEmpty(sheet.ConditionalFormats);
    }

    /// <summary>
    /// 验证三类高级条件格式写入确定的 OOXML 规则。
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
        new ClosedXmlExcelExporter().Export(request, output);
        var xml = ReadWorksheetXml(output);

        Assert.Contains("colorScale>", xml);
        Assert.Contains("type=\"min\"", xml);
        Assert.Contains("type=\"max\"", xml);
        Assert.Contains("dataBar", xml);
        Assert.Contains("iconSet=\"3TrafficLights1\"", xml);
        Assert.Contains("type=\"percent\" val=\"33\"", xml);
        Assert.Contains("type=\"percent\" val=\"67\"", xml);
    }

    /// <summary>
    /// 验证样式、筛选和冻结冲突在数据枚举前拒绝。
    /// </summary>
    /// <param name="scenario">预检失败场景名称。</param>
    [Theory]
    [InlineData("style")]
    [InlineData("filter")]
    [InlineData("freeze")]
    [InlineData("range")]
    [InlineData("print")]
    public void Export_InvalidReportDefinition_ShouldFailBeforeEnumeration(string scenario)
    {
        var data = new ThrowOnEnumeration<ReportRow>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", data, sheet =>
        {
            if (scenario == "style")
                sheet.Table(new ExcelTableDefinition
                {
                    Name = "RowsTable",
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 },
                    StyleName = "NotARealTableStyle"
                });
            else if (scenario == "filter")
            {
                sheet.Table(new ExcelTableDefinition
                {
                    Name = "RowsTable",
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 2, EndColumn = 1 }
                });
                sheet.AutoFilter(new ExcelAutoFilterDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 0, StartColumn = 0, EndRow = 3, EndColumn = 1 }
                });
            }
            else if (scenario == "freeze")
                sheet.FreezePane(new ExcelFreezePaneDefinition { Rows = 1, TopRow = 2 });
            else if (scenario == "range")
                sheet.ConditionalFormat(new ExcelConditionalFormatDefinition
                {
                    Type = ExcelConditionalFormatType.Formula,
                    Formula1 = "TRUE",
                    Range = new ExcelRangeDefinition
                        { StartRow = 0, StartColumn = 0, EndRow = 1048576, EndColumn = 0 }
                });
            else
                sheet.PrintLayout(new ExcelPrintLayoutOptions { ScalePercent = 100, FitToWidth = 1 });
        }));

        var exception = Record.Exception(() =>
            new ClosedXmlExcelExporter().Export(request, new MemoryStream()));

        Assert.True(exception is BingOfficesConfigurationException
            || exception is BingOfficesUnsupportedFeatureException);
        Assert.Equal(0, data.EnumerationCount);
    }

    /// <summary>
    /// 验证保存重开后 Letter、缩放和带 `$` 的重复标题范围保真。
    /// </summary>
    /// <param name="scalePercent">待验证的打印缩放百分比。</param>
    [Theory]
    [InlineData(10)]
    [InlineData(400)]
    public void Export_PrintLayout_ShouldRoundTrip(int scalePercent)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
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
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var setup = workbook.Worksheet("Rows").PageSetup;

        Assert.Equal(XLPaperSize.LetterPaper, setup.PaperSize);
        Assert.Equal(scalePercent, setup.Scale);
        Assert.Single(setup.PrintAreas);
        Assert.Equal(1, setup.FirstRowToRepeatAtTop);
        Assert.Equal(1, setup.LastRowToRepeatAtTop);
        Assert.Equal(1, setup.FirstColumnToRepeatAtLeft);
        Assert.Equal(1, setup.LastColumnToRepeatAtLeft);
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
            new ClosedXmlExcelExporter().Export(request, new MemoryStream()));

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
            new ClosedXmlExcelExporter().Export(request, new MemoryStream()));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
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
