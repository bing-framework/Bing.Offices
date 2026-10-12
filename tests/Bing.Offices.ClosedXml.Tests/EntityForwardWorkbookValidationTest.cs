using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Entities;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML Forward 工作簿原生校验测试。
/// </summary>
public sealed class EntityForwardWorkbookValidationTest
{
    /// <summary>
    /// 验证 Forward 在转换固定字段前执行工作簿原生规则。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldValidateFixedCellBeforeConversion()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateFixedWorkbook(), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportEntityBatches(source, CreateFixedLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Null(summary.Result.Entity.Status);
        Assert.Single(summary.Result.Entity.Lines);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
    }

    /// <summary>
    /// 验证 Forward 在转换前执行明细行的工作簿整数规则。
    /// </summary>
    /// <param name="quantity">输入数量。</param>
    /// <param name="expectedSuccess">是否应导入成功。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(5, true, false)]
    [InlineData(20, false, false)]
    [InlineData(5, true, true)]
    [InlineData(20, false, true)]
    public async Task ImportEntityBatches_Forward_ShouldApplyNativeValidation(
        int quantity, bool expectedSuccess, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(quantity), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.Equal(expectedSuccess, summary.IsSuccess);
        Assert.True(source.CanRead);
        if (expectedSuccess)
        {
            Assert.Single(summary.Result.Entity.Lines);
            Assert.Equal(5, summary.Result.Entity.Lines[0].Quantity);
            Assert.Empty(summary.Result.Errors);
        }
        else
        {
            var error = Assert.Single(summary.Result.Errors);
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
            Assert.Equal(2, error.RowIndex);
            Assert.Equal(2, error.ColumnIndex);
            Assert.Empty(summary.Result.Entity.Lines);
        }
    }

    /// <summary>
    /// 验证低保留失败工作簿按保留行重映射字面量列表校验。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_FailureWorkbook_ShouldRemapLiteralValidation(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
                new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = failure
                }),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var sheet = workbook.Worksheet("明细");
        var validation = Assert.Single(sheet.DataValidations.GetAllInRange(
            sheet.Cell("B2").AsRange().RangeAddress));
        Assert.Equal(XLAllowedValues.List, validation.AllowedValues);
        Assert.Equal("\"1,2,3\"", validation.Value.TrimStart('='));
    }

    /// <summary>
    /// 验证低保留失败工作簿按保留行重映射数值边界校验。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_FailureWorkbook_ShouldRemapNumericValidation()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateNumericFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
                new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = failure
                }),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var validation = Assert.Single(workbook.Worksheet("明细").DataValidations.GetAllInRange(
            workbook.Worksheet("明细").Cell("B2").AsRange().RangeAddress));
        Assert.Equal(XLAllowedValues.WholeNumber, validation.AllowedValues);
        Assert.Equal("1", validation.MinValue);
        Assert.Equal("10", validation.MaxValue);
    }

    /// <summary>
    /// 验证低保留失败工作簿只复制完整落在失败行中的定义名称。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_FailureWorkbook_ShouldRemapSafeDefinedName(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateNamedFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
                ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
                new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = failure
                }),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var retained = Assert.Single(workbook.DefinedNames.ValidNamedRanges()
            .Where(name => name.Name == "FailedQuantity"));
        Assert.EndsWith("!$B$2", retained.RefersTo, StringComparison.Ordinal);
        var local = Assert.Single(workbook.Worksheet("明细").DefinedNames.ValidNamedRanges()
            .Where(name => name.Name == "LocalFailedQuantity"));
        Assert.EndsWith("!$B$2", local.RefersTo, StringComparison.Ordinal);
        Assert.DoesNotContain(workbook.DefinedNames.ValidNamedRanges(), name =>
            name.Name == "SuccessfulQuantity");
    }

    /// <summary>
    /// 创建带整数原生校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(int quantity)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = quantity;
        sheet.Range("B2").CreateDataValidation().WholeNumber.Between(1, 10);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建首行有效、次行转换失败且带列表校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = 2;
        sheet.Cell("A3").Value = "B";
        sheet.Cell("B3").Value = "wrong";
        var validation = sheet.Range("B2:B3").CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.List;
        validation.Value = "\"1,2,3\"";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建首行有效、次行转换失败且带数值边界校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateNumericFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = 2;
        sheet.Cell("A3").Value = "B";
        sheet.Cell("B3").Value = "wrong";
        sheet.Range("B2:B3").CreateDataValidation().WholeNumber.Between(1, 10);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含失败行定义名称的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateNamedFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = 2;
        sheet.Cell("A3").Value = "B";
        sheet.Cell("B3").Value = "wrong";
        workbook.DefinedNames.Add("FailedQuantity", sheet.Range("B3"));
        workbook.DefinedNames.Add("SuccessfulQuantity", sheet.Range("B2"));
        sheet.DefinedNames.Add("LocalFailedQuantity", sheet.Range("B3"));
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带固定单元格列表校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateFixedWorkbook()
    {
        using var workbook = new XLWorkbook();
        var metadata = workbook.AddWorksheet("元数据");
        metadata.Cell("A1").Value = "Bad";
        metadata.Range("A1").CreateDataValidation().List("\"Good\"", true);
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = 1;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建 Forward 列表布局。
    /// </summary>
    private static ExcelEntityLayout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>
        CreateLayout() => ExcelEntity.Layout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含固定字段和明细列表的 Forward 布局。
    /// </summary>
    private static ExcelEntityLayout<FixedDocument> CreateFixedLayout() =>
        ExcelEntity.Layout<FixedDocument>(builder => builder
            .Cell("元数据", "A1", document => document.Status)
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 保存固定字段 Forward 测试单据。
    /// </summary>
    private sealed class FixedDocument
    {
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<EntityLayoutBatchImportTest.BatchLine> Lines { get; } = new();
    }
}
