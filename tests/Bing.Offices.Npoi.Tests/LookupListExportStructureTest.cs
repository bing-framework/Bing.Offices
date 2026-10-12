using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Mappings;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI LookupList 导出的真实工作簿结构测试。
/// </summary>
public sealed class LookupListExportStructureTest
{
    /// <summary>
    /// 验证 XLS 与 XLSX 均生成隐藏字典 Sheet、绝对名称区域和名称列表下拉，且字典采用构建时文本快照。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_LookupList_ShouldWriteHiddenNamedRangeAndSnapshot(ExcelFormat format)
    {
        var values = new List<string> { " Open ", "Closed", "Open" };
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
            .LookupList("OrderStatus", values)
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
        values[0] = "Changed after request creation";
        values.Add("Added later");

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);

        Assert.Equal(2, workbook.NumberOfSheets);
        var data = workbook.GetSheet("Orders");
        var helper = workbook.GetSheetAt(1);
        Assert.StartsWith("__BingLookup_", helper.SheetName, StringComparison.Ordinal);
        Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(workbook.GetSheetIndex(helper)));

        Assert.Equal("Open", helper.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Closed", helper.GetRow(1).GetCell(0).StringCellValue);
        Assert.Null(helper.GetRow(2));

        var definedName = Assert.Single(workbook.GetAllNames());
        Assert.Equal("OrderStatus", definedName.NameName);
        Assert.EndsWith("!$A$1:$A$2", definedName.RefersToFormula, StringComparison.Ordinal);

        var validation = Assert.Single(data.GetDataValidations());
        Assert.Equal(ValidationType.LIST, validation.ValidationConstraint.GetValidationType());
        Assert.Equal("OrderStatus", validation.ValidationConstraint.Formula1);
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(1, region.FirstRow);
        Assert.Equal(2, region.LastRow);
        Assert.Equal(0, region.FirstColumn);
        Assert.Equal(0, region.LastColumn);
    }

    /// <summary>
    /// 验证异步导出能写出可读的字典工作簿结构并保持目标流打开。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportAsync_LookupList_ShouldWriteReadableWorkbook(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
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

        await new NpoiExcelExporter().ExportAsync(request, output);

        Assert.True(output.CanWrite);
        Assert.True(output.Length > 0);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        Assert.Equal(2, workbook.NumberOfSheets);
        var helper = workbook.GetSheetAt(1);
        Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(1));
        Assert.Equal("Open", helper.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Closed", helper.GetRow(1).GetCell(0).StringCellValue);
        Assert.EndsWith("!$A$1:$A$2", Assert.Single(workbook.GetAllNames()).RefersToFormula,
            StringComparison.Ordinal);
        Assert.Equal("OrderStatus", Assert.Single(workbook.GetSheet("Orders").GetDataValidations())
            .ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证 LookupList 拒绝空白字典项。
    /// </summary>
    [Fact]
    public void LookupList_BlankValue_ShouldBeRejected()
    {
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(workbook =>
            workbook.LookupList("OrderStatus", new[] { "Open", "  " })));
    }

    /// <summary>
    /// 验证名称列表与内联列表共存，字典公式形内容保持文本。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_MixedLists_ShouldKeepNamedFormulaAndLiteralText(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
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
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var helper = workbook.GetSheetAt(1);
        Assert.Equal(CellType.String, helper.GetRow(0).GetCell(0).CellType);
        Assert.Equal("=A1", helper.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("A,B", helper.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("第一行\n第二行", helper.GetRow(2).GetCell(0).StringCellValue);

        var validations = workbook.GetSheet("Orders").GetDataValidations();
        Assert.Equal(2, validations.Count);
        Assert.Equal("OrderStatus", validations[0].ValidationConstraint.Formula1);
        Assert.Equal(0, validations[0].Regions.CellRangeAddresses[0].FirstColumn);
        Assert.Equal(1, validations[1].Regions.CellRangeAddresses[0].FirstColumn);
        Assert.Equal(new[] { "Yes", "No" }, validations[1].ValidationConstraint.ExplicitListValues);
    }

    /// <summary>
    /// 验证同区域混用列表时名称公式不被内联列表修正覆盖。
    /// </summary>
    /// <param name="inlineFirst">是否先添加内联列表。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Export_OverlappingLists_ShouldPreserveNamedFormula(bool inlineFirst)
    {
        var named = new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
            Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
        };
        var inline = new ExcelDataValidationDefinition
        {
            Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
            Type = ExcelDataValidationType.List, Values = new[] { "Yes", "No" }
        };
        var request = ExcelExport.Workbook(book => book.LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }, sheet =>
            {
                sheet.DataValidation(inlineFirst ? inline : named);
                sheet.DataValidation(inlineFirst ? named : inline);
            }));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var validations = workbook.GetSheet("Orders").GetDataValidations();
        Assert.Equal(2, validations.Count);
        Assert.Contains(validations, item => item.ValidationConstraint.Formula1 == "Allowed");
        Assert.Contains(validations, item => item.ValidationConstraint.ExplicitListValues is { Length: 2 } values
            && values[0] == "Yes" && values[1] == "No");
    }

    /// <summary>
    /// 验证名称下拉可引用后声明的工作表。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_NamedList_ShouldResolveLaterSheet(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(book => book.Format(format)
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }, sheet =>
                sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 },
                    Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
                }))
            .AddSheet("Dictionary", new[] { new LookupRow { Status = "Open" } }, sheet =>
                sheet.NamedRange(new ExcelNamedRangeDefinition { Name = "Allowed", Address = "$A$2" })));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        Assert.Equal("Allowed", Assert.Single(workbook.GetSheet("Orders").GetDataValidations())
            .ValidationConstraint.Formula1);
        Assert.Equal("Open", workbook.GetSheet("Dictionary").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 验证二维名称来源在写出前被拒绝。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_TwoDimensionalNamedList_ShouldLeaveOutputEmpty(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(book => book.Format(format)
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
        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().Export(request, output));
        Assert.Equal(0, output.Length);
    }

    /// <summary>
    /// 验证字典辅助工作表避开模板已有名称并保留模板内容。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_LookupList_ShouldSkipExistingHelperSheet(ExcelFormat format)
    {
        using IWorkbook templateBook = format == ExcelFormat.Xls
            ? new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        templateBook.CreateSheet("Orders");
        templateBook.CreateSheet("__BingLookup_001").CreateRow(0).CreateCell(0).SetCellValue("Preserved");
        using var template = new MemoryStream();
        templateBook.Write(template, leaveOpen: true);
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.Format(format)
            .UseTemplate(template, leaveOpen: true)
            .LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var actual = WorkbookFactory.Create(output);
        Assert.Equal(3, actual.NumberOfSheets);
        Assert.Equal("Preserved", actual.GetSheet("__BingLookup_001").GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Open", actual.GetSheet("__BingLookup_002").GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(SheetVisibility.Hidden,
            actual.GetSheetVisibility(actual.GetSheetIndex("__BingLookup_002")));
    }

    /// <summary>
    /// 验证模板同名定义在目标流写出前被拒绝。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_LookupList_TemplateNameCollision_ShouldLeaveOutputEmpty(ExcelFormat format)
    {
        using IWorkbook templateBook = format == ExcelFormat.Xls
            ? new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        var sheet = templateBook.CreateSheet("Orders");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Open");
        var name = templateBook.CreateName();
        name.NameName = "Allowed";
        name.RefersToFormula = "Orders!$A$1";
        using var template = new MemoryStream();
        templateBook.Write(template, leaveOpen: true);
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.Format(format)
            .UseTemplate(template, leaveOpen: true)
            .LookupList("Allowed", new[] { "Open" })
            .AddSheet("Orders", new[] { new LookupRow { Status = "Open" } }));

        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().Export(request, output));
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
            var exporter = new NpoiExcelExporter();
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
                new NpoiExcelExporter(valueConverters: new[] { converter })
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
                new NpoiExcelExporter().ExportToFile(request, directory));
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
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_LookupList_LongText_ShouldPreserveFullValue(ExcelFormat format)
    {
        var value = new string('字', 300);
        var request = ExcelExport.Workbook(book => book.Format(format)
            .LookupList("LongValues", new[] { value })
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
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        Assert.Equal(value, workbook.GetSheetAt(1).GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("LongValues", Assert.Single(workbook.GetSheet("Orders").GetDataValidations())
            .ValidationConstraint.Formula1);
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
