using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Entities;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI Forward 工作簿原生校验测试。
/// </summary>
public sealed class EntityForwardWorkbookValidationTest
{
    /// <summary>
    /// 验证 Forward 在转换固定字段前执行工作簿原生规则。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldValidateFixedCellBeforeConversion()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
        using var workbook = new XSSFWorkbook(failure);
        var sheet = workbook.GetSheet("明细");
        var validation = Assert.Single(sheet.GetDataValidations());
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(1, region.FirstRow);
        Assert.Equal(1, region.LastRow);
        Assert.Equal(1, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
        Assert.Equal("\"1,2,3\"", validation.ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证低保留失败工作簿按保留行重映射数值边界校验。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_FailureWorkbook_ShouldRemapNumericValidation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
        using var workbook = new XSSFWorkbook(failure);
        var validation = Assert.Single(workbook.GetSheet("明细").GetDataValidations());
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(1, region.FirstRow);
        Assert.Equal(1, region.LastRow);
        Assert.Equal(1, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
        Assert.Equal("1", validation.ValidationConstraint.Formula1);
        Assert.Equal("10", validation.ValidationConstraint.Formula2);
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
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
        using var workbook = new XSSFWorkbook(failure);
        var retained = Assert.Single(workbook.GetAllNames().Where(name =>
            name.NameName == "FailedQuantity"));
        Assert.EndsWith("!$B$2", retained.RefersToFormula, StringComparison.Ordinal);
        var local = Assert.Single(workbook.GetAllNames().Where(name =>
            name.NameName == "LocalFailedQuantity"));
        Assert.Equal(1, local.SheetIndex);
        Assert.EndsWith("!$B$2", local.RefersToFormula, StringComparison.Ordinal);
        Assert.DoesNotContain(workbook.GetAllNames(), name => name.NameName == "SuccessfulQuantity");
    }

    /// <summary>
    /// 验证 HSSF Forward 在投影后执行工作簿原生整数校验。
    /// </summary>
    /// <param name="quantity">输入数量。</param>
    /// <param name="expectedSuccess">是否应导入成功。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(5, true, false)]
    [InlineData(20, false, false)]
    [InlineData(5, true, true)]
    [InlineData(20, false, true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldApplyNativeValidation(
        int quantity, bool expectedSuccess, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbook(quantity), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHssfLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { });

        Assert.Equal(expectedSuccess, summary.IsSuccess);
        Assert.True(source.CanRead);
        if (expectedSuccess)
        {
            Assert.Single(summary.Result.Entity.Lines);
            Assert.Equal(quantity, summary.Result.Entity.Lines[0].Quantity);
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
    /// 验证 HSSF Forward 在转换后执行安全的显式列表校验。
    /// </summary>
    /// <param name="code">输入编码。</param>
    /// <param name="expectedSuccess">是否应导入成功。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData("A", true, false)]
    [InlineData("C", false, false)]
    [InlineData("A", true, true)]
    [InlineData("C", false, true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldApplyExplicitListValidation(
        string code, bool expectedSuccess, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithExplicitList(code),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHssfLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { });

        Assert.Equal(expectedSuccess, summary.IsSuccess);
        Assert.True(source.CanRead);
        if (expectedSuccess)
        {
            Assert.Single(summary.Result.Entity.Lines);
            Assert.Equal(code, summary.Result.Entity.Lines[0].Code);
            Assert.Empty(summary.Result.Errors);
        }
        else
        {
            var error = Assert.Single(summary.Result.Errors);
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
            Assert.Empty(summary.Result.Entity.Lines);
        }
    }

    /// <summary>
    /// 验证 HSSF Forward 在转换后执行小数边界校验。
    /// </summary>
    /// <param name="quantity">输入数量。</param>
    /// <param name="expectedSuccess">是否应导入成功。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(5, true, false)]
    [InlineData(20, false, false)]
    [InlineData(5, true, true)]
    [InlineData(20, false, true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldApplyDecimalValidation(
        int quantity, bool expectedSuccess, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithDecimalValidation(quantity),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHssfLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { });

        Assert.Equal(expectedSuccess, summary.IsSuccess);
        Assert.True(source.CanRead);
        if (expectedSuccess)
        {
            Assert.Single(summary.Result.Entity.Lines);
            Assert.Equal(quantity, summary.Result.Entity.Lines[0].Quantity);
            Assert.Empty(summary.Result.Errors);
        }
        else
        {
            var error = Assert.Single(summary.Result.Errors);
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
            Assert.Empty(summary.Result.Entity.Lines);
        }
    }

    /// <summary>
    /// 验证 HSSF Forward 在转换后执行文本长度校验。
    /// </summary>
    /// <param name="code">输入编码。</param>
    /// <param name="expectedSuccess">是否应导入成功。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData("AB", true, false)]
    [InlineData("TOO-LONG", false, false)]
    [InlineData("AB", true, true)]
    [InlineData("TOO-LONG", false, true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldApplyTextLengthValidation(
        string code, bool expectedSuccess, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithTextLengthValidation(code),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateHssfLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { });

        Assert.Equal(expectedSuccess, summary.IsSuccess);
        Assert.True(source.CanRead);
        if (expectedSuccess)
        {
            Assert.Single(summary.Result.Entity.Lines);
            Assert.Equal(code, summary.Result.Entity.Lines[0].Code);
            Assert.Empty(summary.Result.Errors);
        }
        else
        {
            var error = Assert.Single(summary.Result.Errors);
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
            Assert.Empty(summary.Result.Entity.Lines);
        }
    }

    /// <summary>
    /// 验证 HSSF Forward 可以读取跨 Sheet 一维命名列表校验。
    /// </summary>
    /// <param name="lookupFirst">命名列表来源是否位于目标 Sheet 之前。</param>
    /// <param name="localScope">命名列表是否使用目标 Sheet 局部作用域。</param>
    /// <param name="specialValues">是否使用包含特殊字符的来源值。</param>
    /// <param name="expectedCode">预期导入的编码。</param>
    [Theory]
    [InlineData(false, false, false, "Open")]
    [InlineData(true, false, false, "Open")]
    [InlineData(false, true, false, "Open")]
    [InlineData(true, true, false, "Open")]
    [InlineData(false, false, true, "A,1")]
    [InlineData(true, true, true, "A,1")]
    public void ImportEntityBatches_Forward_Hssf_ShouldApplyCrossSheetNamedList(
        bool lookupFirst, bool localScope, bool specialValues, string expectedCode)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(
            CreateHssfWorkbookWithCrossSheetNamedList(lookupFirst, localScope, specialValues),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal(expectedCode, summary.Result.Entity.Lines[0].Code);
        Assert.Empty(summary.Result.Errors);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 投影命名列表来源时保留隐藏工作表状态。
    /// </summary>
    /// <param name="visibility">命名列表来源工作表的隐藏状态。</param>
    [Theory]
    [InlineData(SheetVisibility.Hidden)]
    [InlineData(SheetVisibility.VeryHidden)]
    public void Convert_Forward_Hssf_ShouldPreserveNamedListSourceVisibility(
        SheetVisibility visibility)
    {
        using var source = new MemoryStream(
            CreateHssfWorkbookWithCrossSheetNamedList(false, false, false, visibility),
            writable: false);
        using var projected = NpoiHssfForwardXlsxConverter.Convert(source,
            new[] { "明细" }, new ExcelResourceLimits(), CancellationToken.None,
            requireNativeDataValidation: true);
        using var workbook = new XSSFWorkbook(projected);

        var lookupIndex = workbook.GetSheetIndex("Lookup");
        Assert.True(lookupIndex >= 0);
        Assert.Equal(visibility, workbook.GetSheetVisibility(lookupIndex));
    }

    /// <summary>
    /// 验证 HSSF Forward 对无法安全投影的原生校验返回结构化不支持错误。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ShouldRejectUnsupportedNativeValidation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithCustomValidation(),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { }));

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.Contains("原生数据校验", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 拒绝无法安全编码为内联列表公式的值。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ShouldRejectUnsafeExplicitList()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithUnsafeExplicitList(),
            writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.StopOnFirstFailure,
                ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail),
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(source, CreateHssfLayout(), options, _ => { }));

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.Contains("原生数据校验", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建带整数原生校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(int quantity)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(quantity);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(0, "1", "10");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建首行有效、次行转换失败且带列表校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateFailureWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(2);
        sheet.CreateRow(2).CreateCell(0).SetCellValue("B");
        sheet.GetRow(2).CreateCell(1).SetCellValue("wrong");
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateExplicitListConstraint(new[] { "1", "2", "3" });
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 2, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建首行有效、次行转换失败且带数值边界校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateNumericFailureWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(2);
        sheet.CreateRow(2).CreateCell(0).SetCellValue("B");
        sheet.GetRow(2).CreateCell(1).SetCellValue("wrong");
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(0, "1", "10");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 2, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含失败行定义名称的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateNamedFailureWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(2);
        sheet.CreateRow(2).CreateCell(0).SetCellValue("B");
        sheet.GetRow(2).CreateCell(1).SetCellValue("wrong");
        var retained = workbook.CreateName();
        retained.NameName = "FailedQuantity";
        retained.RefersToFormula = "'明细'!$B$3";
        var omitted = workbook.CreateName();
        omitted.NameName = "SuccessfulQuantity";
        omitted.RefersToFormula = "'明细'!$B$2";
        var local = workbook.CreateName();
        local.NameName = "LocalFailedQuantity";
        local.SheetIndex = workbook.GetSheetIndex(sheet);
        local.RefersToFormula = "'明细'!$B$3";
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带固定单元格列表校验的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateFixedWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var metadata = workbook.CreateSheet("元数据");
        metadata.CreateRow(0).CreateCell(0).SetCellValue("Bad");
        var helper = metadata.GetDataValidationHelper();
        metadata.AddValidationData(helper.CreateValidation(
            helper.CreateExplicitListConstraint(new[] { "Good" }),
            new CellRangeAddressList(0, 0, 0, 0)));
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(1);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建用于验证 HSSF Forward 格式边界的 XLS 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbook(int quantity)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(quantity);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(0, "1", "10");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带安全显式列表校验的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithExplicitList(string code)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(code);
        sheet.GetRow(1).CreateCell(1).SetCellValue(5);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateExplicitListConstraint(new[] { "A", "B" });
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带小数边界校验的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithDecimalValidation(int quantity)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(quantity);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateDecimalConstraint(1, "1", "10");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带文本长度校验的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithTextLengthValidation(string code)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(code);
        sheet.GetRow(1).CreateCell(1).SetCellValue(5);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateTextLengthConstraint(1, "1", "3");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带跨 Sheet 命名列表校验的 HSSF 工作簿。
    /// </summary>
    /// <param name="lookupFirst">命名列表来源是否位于目标 Sheet 之前。</param>
    /// <param name="localScope">命名列表是否使用目标 Sheet 局部作用域。</param>
    /// <param name="specialValues">是否使用包含特殊字符的来源值。</param>
    /// <param name="lookupVisibility">命名列表来源工作表的隐藏状态。</param>
    private static byte[] CreateHssfWorkbookWithCrossSheetNamedList(bool lookupFirst,
        bool localScope = false, bool specialValues = false,
        SheetVisibility? lookupVisibility = null)
    {
        using var workbook = new HSSFWorkbook();
        var lookup = lookupFirst ? workbook.CreateSheet("Lookup") : null;
        var sheet = workbook.CreateSheet("明细");
        lookup ??= workbook.CreateSheet("Lookup");
        if (lookupVisibility.HasValue)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(lookup), lookupVisibility.Value);
        var allowedValues = specialValues
            ? new[] { "A,1", "中\n文" }
            : new[] { "Open", "Closed" };
        lookup.CreateRow(0).CreateCell(0).SetCellValue("Allowed");
        lookup.CreateRow(1).CreateCell(0).SetCellValue(allowedValues[0]);
        lookup.CreateRow(2).CreateCell(0).SetCellValue(allowedValues[1]);
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue(allowedValues[0]);
        sheet.GetRow(1).CreateCell(1).SetCellValue(5);

        var name = workbook.CreateName();
        name.NameName = "Allowed";
        if (localScope)
            name.SheetIndex = workbook.GetSheetIndex(sheet);
        name.RefersToFormula = "'Lookup'!$A$2:$A$3";
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateFormulaListConstraint("Allowed");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含逗号列表项的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithUnsafeExplicitList()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A,B");
        sheet.GetRow(1).CreateCell(1).SetCellValue(5);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateExplicitListConstraint(new[] { "A,B", "C" });
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 0, 0)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带无法投影自定义校验的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithCustomValidation()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(5);
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateCustomConstraint("B2>0");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(1, 1, 1, 1)));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建 Forward 列表布局。
    /// </summary>
    private static ExcelEntityLayout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>
        CreateLayout() => ExcelEntity.Layout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建不包含合并区域的 HSSF Forward 列表布局。
    /// </summary>
    private static ExcelEntityLayout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>
        CreateHssfLayout() => ExcelEntity.Layout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>(builder =>
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
