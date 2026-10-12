using System.IO;
using System.Linq;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.ClosedXml;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Mappings;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML LookupList 导出的真实工作簿结构测试。
/// </summary>
public sealed class LookupListExportStructureTest
{
    /// <summary>
    /// 验证 XLSX 生成隐藏字典 Sheet、工作簿名称区域和引用名称的原生下拉校验。
    /// </summary>
    [Fact]
    public void Export_LookupList_ShouldWriteHiddenNamedRangeAndListValidation()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .LookupList("OrderStatus", new[] { "Open", "Closed", "Open" })
            .AddSheet("Orders", new[]
            {
                new LookupRow { Status = "Open" },
                new LookupRow { Status = "Closed" }
            }, sheet => sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 1,
                    EndRow = 2,
                    StartColumn = 0,
                    EndColumn = 0
                },
                Type = ExcelDataValidationType.List,
                ListSourceName = "OrderStatus"
            })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;

        using var workbook = new XLWorkbook(output);
        Assert.Equal(2, workbook.Worksheets.Count);
        var helper = Assert.Single(workbook.Worksheets.Where(sheet =>
            sheet.Name.StartsWith("__BingLookup_", System.StringComparison.Ordinal)));
        Assert.Equal(XLWorksheetVisibility.Hidden, helper.Visibility);
        Assert.Equal("Open", helper.Cell("A1").GetString());
        Assert.Equal("Closed", helper.Cell("A2").GetString());
        Assert.True(helper.Cell("A3").IsEmpty());

        var definedName = Assert.Single(workbook.DefinedNames.ValidNamedRanges()
            .Where(item => item.Name == "OrderStatus"));
        var range = Assert.Single(definedName.Ranges);
        Assert.Equal(helper.Name, range.Worksheet.Name);
        Assert.Equal(1, range.RangeAddress.FirstAddress.RowNumber);
        Assert.Equal(1, range.RangeAddress.FirstAddress.ColumnNumber);
        Assert.Equal(2, range.RangeAddress.LastAddress.RowNumber);
        Assert.Equal(1, range.RangeAddress.LastAddress.ColumnNumber);

        var data = workbook.Worksheet("Orders");
        var validation = Assert.Single(data.DataValidations.GetAllInRange(
            data.Range("A2:A3").RangeAddress));
        Assert.Equal(XLAllowedValues.List, validation.AllowedValues);
        Assert.Equal("OrderStatus", validation.Value.TrimStart('='));
    }

    /// <summary>
    /// 验证异步导出能写出可读的字典工作簿结构并保持目标流打开。
    /// </summary>
    [Fact]
    public async Task ExportAsync_LookupList_ShouldWriteReadableWorkbook()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .LookupList("OrderStatus", new[] { "Open", "Closed" })
            .AddSheet("Orders", new[]
            {
                new LookupRow { Status = "Open" },
                new LookupRow { Status = "Closed" }
            }, sheet => sheet.DataValidation(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 0
                },
                Type = ExcelDataValidationType.List,
                ListSourceName = "OrderStatus"
            })));
        using var output = new MemoryStream();

        await new ClosedXmlExcelExporter().ExportAsync(request, output);

        Assert.True(output.CanWrite);
        Assert.True(output.Length > 0);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal(2, workbook.Worksheets.Count);
        var helper = workbook.Worksheets.Last();
        Assert.Equal(XLWorksheetVisibility.Hidden, helper.Visibility);
        Assert.Equal("Open", helper.Cell("A1").GetString());
        Assert.Equal("Closed", helper.Cell("A2").GetString());
        var definedName = Assert.Single(workbook.DefinedNames.ValidNamedRanges()
            .Where(item => item.Name == "OrderStatus"));
        var range = Assert.Single(definedName.Ranges);
        Assert.Equal(helper.Name, range.Worksheet.Name);
        Assert.Equal(1, range.RangeAddress.FirstAddress.RowNumber);
        Assert.Equal(2, range.RangeAddress.LastAddress.RowNumber);
        var data = workbook.Worksheet("Orders");
        var validation = Assert.Single(data.DataValidations.GetAllInRange(data.Range("A2:A3").RangeAddress));
        Assert.Equal("OrderStatus", validation.Value.TrimStart('='));
    }

    /// <summary>
    /// 验证名称列表与内联列表共存，字典公式形内容保持文本。
    /// </summary>
    [Fact]
    public void Export_MixedLists_ShouldKeepNamedFormulaAndLiteralText()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .LookupList("OrderStatus", new[] { "=A1", "A,B", "第一行\n第二行" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }, sheet => sheet
                .DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                    Type = ExcelDataValidationType.List,
                    ListSourceName = "OrderStatus"
                })
                .DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1 },
                    Type = ExcelDataValidationType.List,
                    Values = new[] { "Yes", "No" }
                })));

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var helper = workbook.Worksheets.Last();
        Assert.Equal(XLDataType.Text, helper.Cell("A1").DataType);
        Assert.Equal("=A1", helper.Cell("A1").GetString());
        Assert.Equal("A,B", helper.Cell("A2").GetString());
        Assert.Equal("第一行\n第二行", helper.Cell("A3").GetString());

        var sheet = workbook.Worksheet("Orders");
        var named = Assert.Single(sheet.DataValidations.GetAllInRange(sheet.Cell("A2").AsRange().RangeAddress));
        var inline = Assert.Single(sheet.DataValidations.GetAllInRange(sheet.Cell("B2").AsRange().RangeAddress));
        Assert.Equal("OrderStatus", named.Value.TrimStart('='));
        Assert.Contains("Yes", inline.Value);
        Assert.Contains("No", inline.Value);
    }

    /// <summary>
    /// 验证名称下拉可引用后声明的工作表。
    /// </summary>
    [Fact]
    public void Export_NamedList_ShouldResolveLaterSheet()
    {
        var request = ExcelExport.Workbook(book => book
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }, sheet =>
                sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                    Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
                }))
            .AddSheet("Dictionary", new[] { new LookupRow { Status = "Open" } }, sheet =>
                sheet.NamedRange(new ExcelNamedRangeDefinition { Name = "Allowed", Address = "$A$2" })));

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal("Allowed", Assert.Single(workbook.Worksheet("Orders").DataValidations).Value.TrimStart('='));
        Assert.Equal("Open", workbook.Worksheet("Dictionary").Cell("A2").GetString());
    }

    /// <summary>
    /// 验证二维名称来源在写出前被拒绝。
    /// </summary>
    [Fact]
    public void Export_TwoDimensionalNamedList_ShouldLeaveOutputEmpty()
    {
        var request = ExcelExport.Workbook(book => book
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }, sheet =>
                sheet.NamedRange(new ExcelNamedRangeDefinition
                {
                    Name = "Allowed", Address = "$A$1:$B$2"
                }).DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                    Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
                })));

        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() => new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>
    /// 验证字典辅助工作表避开模板已有名称并保留模板内容。
    /// </summary>
    [Fact]
    public void Export_LookupList_ShouldSkipExistingHelperSheet()
    {
        using var templateBook = new XLWorkbook();
        templateBook.Worksheets.Add("Orders");
        templateBook.Worksheets.Add("__BingLookup_001").Cell("A1").Value = "Preserved";
        using var template = new MemoryStream();
        templateBook.SaveAs(template);
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var actual = new XLWorkbook(output);
        Assert.Equal(3, actual.Worksheets.Count);
        Assert.Equal("Preserved", actual.Worksheet("__BingLookup_001").Cell("A1").GetString());
        Assert.Equal("Open", actual.Worksheet("__BingLookup_002").Cell("A1").GetString());
        Assert.Equal(XLWorksheetVisibility.Hidden, actual.Worksheet("__BingLookup_002").Visibility);
    }

    /// <summary>
    /// 验证模板同名定义在目标流写出前被拒绝。
    /// </summary>
    [Fact]
    public void Export_LookupList_TemplateNameCollision_ShouldLeaveOutputEmpty()
    {
        using var templateBook = new XLWorkbook();
        var sheet = templateBook.Worksheets.Add("Orders");
        sheet.Cell("A1").Value = "Open";
        templateBook.DefinedNames.Add("Allowed", sheet.Range("A1"));
        using var template = new MemoryStream();
        templateBook.SaveAs(template);
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));

        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() => new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证预取消的字典导出保留旧目标并清理临时文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步文件入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Export_LookupList_PreCanceled_ShouldPreserveTarget(bool asynchronous)
    {
        var request = ExcelExport.Workbook(book => book.LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));
        var directory = Path.Combine(Path.GetTempPath(), "bing-lookup-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "orders.xlsx");
        File.WriteAllText(target, "Original", Encoding.UTF8);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        try
        {
            var exporter = new ClosedXmlExcelExporter();
            if (asynchronous)
                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    exporter.ExportToFileAsync(request, target, cancellation.Token));
            else
                Assert.Throws<OperationCanceledException>(() =>
                    exporter.ExportToFile(request, target, cancellation.Token));
            Assert.Equal("Original", File.ReadAllText(target, Encoding.UTF8));
            Assert.Equal(new[] { target }, Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证字典请求在明细写入中途取消时保留旧文件并清理临时文件。
    /// </summary>
    [Fact]
    public async Task ExportAsync_LookupList_MidFlightCancellation_ShouldPreserveTargetAndCleanTemporaryFile()
    {
        var request = ExcelExport.Workbook(book => book.LookupList("Allowed", new[] { "Open", "Closed" })
            .AddSheet("Orders", new[]
            {
                new LookupRow { Status = "Open" },
                new LookupRow { Status = "Closed" }
            }, sheet => sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(LookupRow.Status), ConverterName = "lookup-cancel" }
                }
            }).DataValidation(new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 0 },
                Type = ExcelDataValidationType.List,
                ListSourceName = "Allowed"
            })));
        var directory = Path.Combine(Path.GetTempPath(), "bing-lookup-midflight-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var target = Path.Combine(directory, "orders.xlsx");
        var original = new byte[] { 9, 8, 7, 6 };
        File.WriteAllBytes(target, original);
        using var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstLookupExportConverter(cancellation);

        try
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                new ClosedXmlExcelExporter(valueConverters: new[] { converter })
                    .ExportToFileAsync(request, target, cancellation.Token));

            Assert.Equal(1, converter.ConvertToCalls);
            Assert.Equal(original, File.ReadAllBytes(target));
            Assert.Equal(new[] { target }, Directory.GetFiles(directory));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证字典工作簿提交失败后保留目标目录并清理临时文件。
    /// </summary>
    [Fact]
    public void Export_LookupList_CommitFailure_ShouldCleanTemporaryFile()
    {
        var request = ExcelExport.Workbook(book => book.LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));
        var directory = Path.Combine(Path.GetTempPath(), "bing-lookup-commit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var exception = Assert.Throws<BingOfficesFileCommitException>(() =>
                new ClosedXmlExcelExporter().ExportToFile(request, directory));
            Assert.Equal(BingOfficesStage.Commit, exception.Stage);
            Assert.True(Directory.Exists(directory));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(directory),
                Path.GetFileName(directory) + ".*.tmp"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 验证超过内联列表限制的文本由字典单元格完整保存。
    /// </summary>
    [Fact]
    public void Export_LookupList_LongText_ShouldPreserveFullValue()
    {
        var value = new string('字', 300);
        var request = ExcelExport.Workbook(book => book.LookupList("LongValues", new[] { value })
            .AddSheet("Orders", new[] { new LookupRow { Status = value } }, sheet =>
                sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0
                    },
                    Type = ExcelDataValidationType.List,
                    ListSourceName = "LongValues"
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal(value, workbook.Worksheet(2).Cell("A1").GetString());
        Assert.Equal("LongValues", Assert.Single(workbook.Worksheet("Orders").DataValidations)
            .Value.TrimStart('='));
    }

    /// <summary>
    /// 名称列表导出测试行。
    /// </summary>
    private sealed class LookupRow
    {
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }
    }

    /// <summary>
    /// 首次转换后触发取消的测试转换器。
    /// </summary>
    private sealed class CancelAfterFirstLookupExportConverter : INamedExcelValueConverter
    {
        /// <summary>
        /// 保存用于触发取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstLookupExportConverter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">用于触发取消的令牌源。</param>
        public CancelAfterFirstLookupExportConverter(CancellationTokenSource cancellation)
        {
            _cancellation = cancellation;
        }

        /// <inheritdoc />
        public string Name => "lookup-cancel";

        /// <summary>
        /// 获取已执行的导出转换调用次数。
        /// </summary>
        public int ConvertToCalls { get; private set; }

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            ConvertToCalls++;
            if (ConvertToCalls == 1)
                _cancellation.Cancel();
            return true;
        }
    }
}
