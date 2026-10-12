using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI Forward 低保留失败工作簿。
/// </summary>
public sealed class EntityHssfLowRetentionFailureWorkbookTest
{
    /// <summary>
    /// 验证同步和异步 HSSF Forward 只保留表头与失败行。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldWriteSparseWorkbook(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.False(summary.HasCompleteResult);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        Assert.IsType<HSSFWorkbook>(workbook);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Equal(1, sheet.LastRowNum);
        Assert.Equal(2, sheet.PhysicalNumberOfRows);
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Quantity", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("B", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("wrong", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Null(sheet.GetRow(2));
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));

        source.Position = 0;
        Assert.True(source.ReadByte() >= 0);
    }

    /// <summary>
    /// 验证固定单元格失败时使用稀疏输出并保留固定字段错误上下文。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithFixedCellFailure_ShouldRetainContext()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateFixedCellFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateFixedCellFailureLayout(), options,
            _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.Result.Errors.Count);
        Assert.Contains(summary.Result.Errors, error => error.Code == ExcelImportErrorCode.ValueConversion
            && error.SheetName == "单据" && error.RowIndex == 1 && error.ColumnIndex == 1
            && error.PropertyName == nameof(FailureDocument.Number));

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        Assert.IsType<HSSFWorkbook>(workbook);
        var documentSheet = workbook.GetSheet("单据");
        Assert.NotNull(documentSheet);
        Assert.Equal("wrong", documentSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("采购单", documentSheet.GetRow(0).GetCell(1).StringCellValue);

        var detailSheet = workbook.GetSheet("明细");
        Assert.NotNull(detailSheet);
        Assert.Equal("Code", detailSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("B", detailSheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("wrong", detailSheet.GetRow(1).GetCell(1).StringCellValue);

        var errorSheet = workbook.GetSheet("_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.Equal("单据", errorSheet.GetRow(1).GetCell(2).StringCellValue);
        Assert.Equal(1d, errorSheet.GetRow(1).GetCell(3).NumericCellValue);
        Assert.Equal(1d, errorSheet.GetRow(1).GetCell(4).NumericCellValue);
        Assert.Equal(nameof(FailureDocument.Number), errorSheet.GetRow(1).GetCell(5).StringCellValue);
        Assert.Equal("wrong", errorSheet.GetRow(1).GetCell(7).StringCellValue);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建只输出错误行的实体导入选项。
    /// </summary>
    /// <param name="destination">失败工作簿目标流。</param>
    /// <returns>实体批次导入选项。</returns>
    private static ExcelEntityImportOptions CreateImportOptions(Stream destination,
        long? maxEstimatedTargetObjects = null) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination,
                MaxEstimatedTargetObjects = maxEstimatedTargetObjects
            });

    /// <summary>
    /// 创建实体列表布局。
    /// </summary>
    /// <returns>包含明细列表区域的实体布局。</returns>
    private static ExcelEntityLayout<FailureDocument> CreateLayout() =>
        ExcelEntity.Layout<FailureDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含固定单元格和明细列表的实体布局。
    /// </summary>
    /// <returns>包含单据头和明细区域的实体布局。</returns>
    private static ExcelEntityLayout<FailureDocument> CreateFixedCellFailureLayout() =>
        ExcelEntity.Layout<FailureDocument>(builder => builder
            .Cell("单据", "A1", document => document.Number)
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 验证 HSSF 低保留失败工作簿保留列宽和隐藏列状态。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldPreserveColumnMetadata()
    {
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");

        Assert.NotNull(sheet);
        Assert.Equal(18 * 256, sheet.GetColumnWidth(0));
        Assert.True(sheet.IsColumnHidden(1));
        Assert.Equal(30f, sheet.GetRow(0).HeightInPoints);
        Assert.Equal(24f, sheet.GetRow(1).HeightInPoints);
        Assert.True(sheet.GetRow(1).ZeroHeight);
        Assert.Null(sheet.GetRow(2));
        Assert.Equal(2, sheet.PhysicalNumberOfRows);
    }

    /// <summary>
    /// 验证 HSSF 低保留失败工作簿保留冻结窗格坐标。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldPreserveFreezePane()
    {
        using var source = new MemoryStream(CreateWorkbook(includeFreezePane: true), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var pane = workbook.GetSheet("明细").PaneInformation;
        Assert.NotNull(pane);
        Assert.True(pane.IsFreezePane());
        Assert.Equal(1, pane.VerticalSplitPosition);
        Assert.Equal(1, pane.HorizontalSplitPosition);
        Assert.Equal(1, pane.VerticalSplitLeftColumn);
        Assert.Equal(1, pane.HorizontalSplitTopRow);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证低保留失败工作簿保留表头字体和失败行数字格式。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldPreserveCellStyles()
    {
        using var source = new MemoryStream(CreateStyledWorkbook(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");

        Assert.NotNull(sheet);
        var headerCell = sheet.GetRow(0).GetCell(0);
        var headerFont = workbook.GetFontAt(headerCell.CellStyle.FontIndex);
        Assert.True(headerFont.IsBold);
        Assert.True(headerFont.IsItalic);
        Assert.True(headerFont.IsStrikeout);
        Assert.Equal(FontUnderlineType.Single, headerFont.Underline);
        Assert.Equal("Arial", headerFont.FontName);
        Assert.Equal(14, headerFont.FontHeightInPoints);
        Assert.Equal(IndexedColors.Red.Index, headerFont.Color);
        Assert.Equal(HorizontalAlignment.Center, headerCell.CellStyle.Alignment);
        Assert.Equal(VerticalAlignment.Center, headerCell.CellStyle.VerticalAlignment);
        Assert.True(headerCell.CellStyle.WrapText);
        Assert.True(headerCell.CellStyle.ShrinkToFit);
        Assert.Equal(15, headerCell.CellStyle.Rotation);
        Assert.Equal(2, headerCell.CellStyle.Indention);
        Assert.Equal(FillPattern.SolidForeground, headerCell.CellStyle.FillPattern);
        Assert.Equal(IndexedColors.Yellow.Index, headerCell.CellStyle.FillForegroundColor);
        Assert.Equal(IndexedColors.Blue.Index, headerCell.CellStyle.FillBackgroundColor);
        Assert.Equal(BorderStyle.Thin, headerCell.CellStyle.BorderLeft);
        Assert.Equal(IndexedColors.Green.Index, headerCell.CellStyle.LeftBorderColor);
        Assert.Equal(BorderStyle.Medium, headerCell.CellStyle.BorderBottom);
        Assert.Equal(IndexedColors.Blue.Index, headerCell.CellStyle.BottomBorderColor);
        Assert.False(headerCell.CellStyle.IsLocked);
        var quantityHeaderFont = workbook.GetFontAt(sheet.GetRow(0).GetCell(1).CellStyle.FontIndex);
        Assert.Equal("宋体", quantityHeaderFont.FontName);
        Assert.Equal(12, quantityHeaderFont.FontHeightInPoints);
        var failureQuantityStyle = sheet.GetRow(1).GetCell(1).CellStyle;
        Assert.Equal("#,##0.00", failureQuantityStyle.GetDataFormatString());
        Assert.True(failureQuantityStyle.IsHidden);
    }

    /// <summary>
    /// 验证 XSSF 低保留失败工作簿保留表头和失败行单元格样式。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Xssf_ErrorRowsOnly_ShouldPreserveCellStyles()
    {
        using var source = new MemoryStream(CreateXlsxStyledWorkbook(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = (XSSFWorkbook)WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");

        Assert.NotNull(sheet);
        Assert.Equal(22 * 256, sheet.GetColumnWidth(0));
        Assert.Equal(16 * 256, sheet.GetColumnWidth(1));
        Assert.True(sheet.IsColumnHidden(1));
        Assert.False(sheet.IsColumnHidden(2));
        var headerCell = sheet.GetRow(0).GetCell(0);
        var headerFont = workbook.GetFontAt(headerCell.CellStyle.FontIndex);
        Assert.True(headerFont.IsBold);
        Assert.True(headerFont.IsItalic);
        Assert.True(headerFont.IsStrikeout);
        Assert.Equal(FontUnderlineType.Single, headerFont.Underline);
        Assert.Equal("Arial", headerFont.FontName);
        Assert.Equal(14, headerFont.FontHeightInPoints);
        Assert.Equal(IndexedColors.Red.Index, headerFont.Color);
        Assert.Equal(HorizontalAlignment.Center, headerCell.CellStyle.Alignment);
        Assert.Equal(VerticalAlignment.Center, headerCell.CellStyle.VerticalAlignment);
        Assert.True(headerCell.CellStyle.WrapText);
        Assert.True(headerCell.CellStyle.ShrinkToFit);
        Assert.Equal(15, headerCell.CellStyle.Rotation);
        Assert.Equal(2, headerCell.CellStyle.Indention);
        Assert.Equal(FillPattern.SolidForeground, headerCell.CellStyle.FillPattern);
        Assert.Equal(IndexedColors.Yellow.Index, headerCell.CellStyle.FillForegroundColor);
        Assert.Equal(IndexedColors.Blue.Index, headerCell.CellStyle.FillBackgroundColor);
        Assert.Equal(BorderStyle.Thin, headerCell.CellStyle.BorderLeft);
        Assert.Equal(IndexedColors.Green.Index, headerCell.CellStyle.LeftBorderColor);
        Assert.Equal(BorderStyle.Medium, headerCell.CellStyle.BorderBottom);
        Assert.Equal(IndexedColors.Blue.Index, headerCell.CellStyle.BottomBorderColor);
        Assert.False(headerCell.CellStyle.IsLocked);
        var quantityHeaderFont = workbook.GetFontAt(sheet.GetRow(0).GetCell(1).CellStyle.FontIndex);
        Assert.Equal("宋体", quantityHeaderFont.FontName);
        Assert.Equal(12, quantityHeaderFont.FontHeightInPoints);
        var failureQuantityStyle = sheet.GetRow(1).GetCell(1).CellStyle;
        Assert.Equal("#,##0.00", failureQuantityStyle.GetDataFormatString());
        Assert.True(failureQuantityStyle.IsHidden);
    }

    /// <summary>
    /// 验证 HSSF Forward 能在稀疏失败工作簿中保留安全条件格式。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldPreserveConditionalFormatting(
        bool asynchronous)
    {
        using var source = new MemoryStream(CreateWorkbookWithConditionalFormatting(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ValueConversion);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        Assert.IsType<HSSFWorkbook>(workbook);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        var conditionalFormatting = sheet.SheetConditionalFormatting;
        Assert.Equal(1, conditionalFormatting.NumConditionalFormattings);
        var formatting = conditionalFormatting.GetConditionalFormattingAt(0);
        var range = Assert.Single(formatting.GetFormattingRanges());
        Assert.Equal(1, range.FirstRow);
        Assert.Equal(1, range.LastRow);
        Assert.Equal(1, range.FirstColumn);
        Assert.Equal(1, range.LastColumn);
        var rule = formatting.GetRule(0);
        Assert.Equal(ConditionType.CellValueIs, rule.ConditionType);
        Assert.Equal(ComparisonOperator.GreaterThan, rule.ComparisonOperation);
        Assert.Equal("0", rule.Formula1);
        Assert.Equal(FillPattern.SolidForeground, rule.PatternFormatting.FillPattern);
        Assert.Equal(IndexedColors.Yellow.Index, rule.PatternFormatting.FillBackgroundColor);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 能保留同一区域的多条条件格式规则及其顺序。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldPreserveConditionalFormattingRuleOrder()
    {
        using var source = new MemoryStream(CreateWorkbookWithMultipleConditionalFormatting(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = Assert.IsType<HSSFWorkbook>(WorkbookFactory.Create(failure));
        var conditionalFormatting = workbook.GetSheet("明细")!.SheetConditionalFormatting;
        var formatting = conditionalFormatting.GetConditionalFormattingAt(0);
        Assert.Equal(2, formatting.NumberOfRules);
        Assert.Equal(ConditionType.CellValueIs, formatting.GetRule(0).ConditionType);
        Assert.Equal(ComparisonOperator.GreaterThan, formatting.GetRule(0).ComparisonOperation);
        Assert.Equal("0", formatting.GetRule(0).Formula1);
        Assert.Equal(ConditionType.CellValueIs, formatting.GetRule(1).ConditionType);
        Assert.Equal(ComparisonOperator.LessThan, formatting.GetRule(1).ComparisonOperation);
        Assert.Equal("10", formatting.GetRule(1).Formula1);
        Assert.Equal(IndexedColors.Yellow.Index,
            formatting.GetRule(0).PatternFormatting.FillBackgroundColor);
        Assert.Equal(IndexedColors.Red.Index,
            formatting.GetRule(1).PatternFormatting.FillBackgroundColor);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 遇到复杂条件格式时结构化拒绝，避免静默丢失规则。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldRejectComplexConditionalFormatting()
    {
        using var source = new MemoryStream(CreateWorkbookWithComplexConditionalFormatting(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Contains("条件格式", exception.Message);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.Equal(0, failure.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XSSF Forward 在源坐标稳定时保留安全条件格式规则。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Xssf_ErrorRowsOnly_ShouldPreserveConditionalFormatting(
        bool asynchronous)
    {
        using var source = new MemoryStream(CreateXlsxWorkbookWithConditionalFormatting(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = Assert.IsType<XSSFWorkbook>(WorkbookFactory.Create(failure));
        var conditionalFormatting = workbook.GetSheet("明细")!.SheetConditionalFormatting;
        Assert.Equal(1, conditionalFormatting.NumConditionalFormattings);
        var formatting = conditionalFormatting.GetConditionalFormattingAt(0);
        var range = Assert.Single(formatting.GetFormattingRanges());
        Assert.Equal(1, range.FirstRow);
        Assert.Equal(1, range.LastRow);
        Assert.Equal(1, range.FirstColumn);
        Assert.Equal(1, range.LastColumn);
        var rule = formatting.GetRule(0);
        Assert.Equal(ConditionType.CellValueIs, rule.ConditionType);
        Assert.Equal(ComparisonOperator.GreaterThan, rule.ComparisonOperation);
        Assert.Equal("0", rule.Formula1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XSSF Forward 在错误行压缩会改变条件格式坐标时结构化拒绝。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Xssf_ErrorRowsOnly_ShouldRejectCompressedConditionalFormatting()
    {
        using var source = new MemoryStream(
            CreateXlsxWorkbookWithConditionalFormatting(includeSuccessfulRow: true),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesErrorCode.UnsupportedFeature, exception.Code);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("源坐标不变", exception.Message);
        Assert.Equal(0, failure.Length);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证简单数值原生校验可以随稀疏失败行写出。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithNativeDataValidation_ShouldWriteSparseValidation(
        bool asynchronous)
    {
        using var source = new MemoryStream(CreateWorkbookWithNativeDataValidation(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ValueConversion);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        Assert.IsType<HSSFWorkbook>(workbook);
        var sheet = workbook.GetSheet("明细");

        Assert.NotNull(sheet);
        Assert.Equal(2, sheet.LastRowNum);
        Assert.Equal(2, sheet.PhysicalNumberOfRows);
        Assert.Null(sheet.GetRow(1));
        Assert.Equal("B", sheet.GetRow(2).GetCell(0).StringCellValue);

        var validation = Assert.Single(sheet.GetDataValidations());
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(2, region.FirstRow);
        Assert.Equal(2, region.LastRow);
        Assert.Equal(1, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
        var constraint = Assert.IsType<DVConstraint>(validation.ValidationConstraint);
        Assert.Equal(1d, constraint.Value1);
        Assert.Equal(10d, constraint.Value2);
    }

    /// <summary>
    /// 验证跨越未保留成功行的原生校验按保留失败行拆分重建。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithCrossRowNativeDataValidation_ShouldRemapSparseRows()
    {
        using var source = new MemoryStream(CreateWorkbookWithNativeDataValidation(true),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");
        Assert.Equal(2, sheet.LastRowNum);
        Assert.Null(sheet.GetRow(1));
        Assert.Equal("B", sheet.GetRow(2).GetCell(0).StringCellValue);
        var validation = Assert.Single(sheet.GetDataValidations());
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(2, region.FirstRow);
        Assert.Equal(2, region.LastRow);
        Assert.Equal(1, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
    }

    /// <summary>
    /// 验证命名列表来源完整保留时可以随稀疏失败行写出。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithNamedListValidation_ShouldWriteSparseValidation()
    {
        using var source = new MemoryStream(CreateWorkbookWithNamedListDataValidation(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Null(sheet.GetRow(1));
        Assert.Equal("B", sheet.GetRow(2).GetCell(0).StringCellValue);

        var namedList = Assert.Single(workbook.GetAllNames()
            .Where(name => name.NameName == "Allowed"));
        Assert.Equal("明细!$A$3", NormalizeNameFormula(namedList.RefersToFormula));

        var validation = Assert.Single(sheet.GetDataValidations());
        var region = Assert.Single(validation.Regions.CellRangeAddresses);
        Assert.Equal(2, region.FirstRow);
        Assert.Equal(2, region.LastRow);
        Assert.Equal(1, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
        Assert.Equal(ValidationType.LIST, validation.ValidationConstraint.GetValidationType());
        Assert.Equal("Allowed", validation.ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证局部作用域命名列表可以随稀疏失败行写出。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithLocalNamedListValidation_ShouldWriteSparseValidation()
    {
        using var source = new MemoryStream(CreateWorkbookWithNamedListDataValidation(true),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");
        var namedList = Assert.Single(workbook.GetAllNames()
            .Where(name => name.NameName == "Allowed"));
        Assert.Equal(workbook.GetSheetIndex(sheet), namedList.SheetIndex);
        Assert.Equal("明细!$A$3", NormalizeNameFormula(namedList.RefersToFormula));
        var validation = Assert.Single(sheet.GetDataValidations());
        Assert.Equal(ValidationType.LIST, validation.ValidationConstraint.GetValidationType());
        Assert.Equal("Allowed", validation.ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证显式列表值可以随稀疏失败行写出。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithExplicitListValidation_ShouldWriteSparseValidation()
    {
        using var source = new MemoryStream(CreateWorkbookWithExplicitListDataValidation(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Null(sheet.GetRow(1));
        var validation = Assert.Single(sheet.GetDataValidations());
        var constraint = Assert.IsType<DVConstraint>(validation.ValidationConstraint);
        Assert.Equal(ValidationType.LIST, constraint.GetValidationType());
        Assert.True(constraint.IsExplicitList);
        Assert.Equal(new[] { "A,1", "B", "中文" }, constraint.ExplicitListValues);
    }

    /// <summary>
    /// 验证跨 Sheet 命名列表来源可以随稀疏失败行写出。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithCrossSheetNamedListValidation_ShouldWriteSparseValidation()
    {
        using var source = new MemoryStream(CreateWorkbookWithCrossSheetNamedListDataValidation(),
            writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        Assert.IsType<HSSFWorkbook>(workbook);
        var lookup = workbook.GetSheet("Lookup");
        Assert.NotNull(lookup);
        Assert.Equal("Open", lookup.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("Closed", lookup.GetRow(2).GetCell(0).StringCellValue);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Equal("B", sheet.GetRow(2).GetCell(0).StringCellValue);
        var namedList = Assert.Single(workbook.GetAllNames()
            .Where(name => name.NameName == "Allowed"));
        Assert.Equal("Lookup!$A$2:$A$3", NormalizeNameFormula(namedList.RefersToFormula));
        var validation = Assert.Single(sheet.GetDataValidations());
        Assert.Equal(ValidationType.LIST, validation.ValidationConstraint.GetValidationType());
        Assert.Equal("Allowed", validation.ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证隐藏的跨 Sheet 命名列表来源仍保留名称、数据和可见性。
    /// </summary>
    /// <param name="visibility">命名列表来源的隐藏状态。</param>
    [Theory]
    [InlineData(SheetVisibility.Hidden)]
    [InlineData(SheetVisibility.VeryHidden)]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithHiddenCrossSheetNamedListValidation_ShouldPreserveSource(
        SheetVisibility visibility)
    {
        using var source = new MemoryStream(
            CreateWorkbookWithCrossSheetNamedListDataValidation(visibility), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var lookup = workbook.GetSheet("Lookup");
        Assert.NotNull(lookup);
        Assert.Equal(visibility, workbook.GetSheetVisibility(workbook.GetSheetIndex(lookup)));
        Assert.Equal("Open", lookup.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("Closed", lookup.GetRow(2).GetCell(0).StringCellValue);
        Assert.Single(workbook.GetSheet("明细").GetDataValidations());
    }

    /// <summary>
    /// 验证命名列表来源行计入低保留失败输出的对象预算。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_WithCrossSheetNamedListValidation_ShouldHonorObjectBudget()
    {
        using var source = new MemoryStream(
            CreateWorkbookWithCrossSheetNamedListDataValidation(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = CreateImportOptions(failure, maxEstimatedTargetObjects: 4);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证异步低保留失败输出同样计入命名列表来源行的对象预算。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_Hssf_ErrorRowsOnly_WithCrossSheetNamedListValidation_ShouldHonorObjectBudget()
    {
        using var source = new MemoryStream(
            CreateWorkbookWithCrossSheetNamedListDataValidation(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = CreateImportOptions(failure, maxEstimatedTargetObjects: 4);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证低保留失败工作簿只复制完全位于保留行中的合并区域。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldCopyOnlyRetainedMerges()
    {
        using var source = new MemoryStream(CreateWorkbookWithMerges(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var sheet = workbook.GetSheet("明细");

        Assert.NotNull(sheet);
        Assert.Equal(2, sheet.NumMergedRegions);
        Assert.Contains(sheet.MergedRegions, range => range.FirstRow == 0
            && range.LastRow == 0 && range.FirstColumn == 0 && range.LastColumn == 1);
        Assert.Contains(sheet.MergedRegions, range => range.FirstRow == 1
            && range.LastRow == 1 && range.FirstColumn == 0 && range.LastColumn == 1);
        Assert.DoesNotContain(sheet.MergedRegions, range => range.FirstColumn == 2
            && range.LastColumn == 3);
    }

    /// <summary>
    /// 验证低保留失败工作簿只复制能够映射到保留行的一维定义名称。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ErrorRowsOnly_ShouldRemapSafeDefinedNames()
    {
        using var source = new MemoryStream(CreateWorkbookWithDefinedNames(), writable: false);
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);

        failure.Position = 0;
        using var workbook = WorkbookFactory.Create(failure);
        var global = Assert.Single(workbook.GetAllNames().Where(name => name.NameName == "FailureCell"));
        Assert.Equal("明细!$A$2", NormalizeNameFormula(global.RefersToFormula));
        Assert.Equal(-1, global.SheetIndex);
        var local = Assert.Single(workbook.GetAllNames().Where(name => name.NameName == "LocalFailureCell"));
        Assert.Equal("明细!$A$2", NormalizeNameFormula(local.RefersToFormula));
        Assert.Equal(workbook.GetSheetIndex("明细"), local.SheetIndex);
        var column = Assert.Single(workbook.GetAllNames().Where(name => name.NameName == "FailureColumn"));
        Assert.Equal("明细!$A$2:$A$3", NormalizeNameFormula(column.RefersToFormula));
        var header = Assert.Single(workbook.GetAllNames().Where(name => name.NameName == "HeaderCell"));
        Assert.Equal("明细!$A$1", NormalizeNameFormula(header.RefersToFormula));
        Assert.DoesNotContain(workbook.GetAllNames(), name => name.NameName == "SuccessfulRows");
    }

    /// <summary>
    /// 统一 HSSF 对简单中文工作表名称的可选引号表示。
    /// </summary>
    /// <param name="formula">定义名称公式。</param>
    /// <returns>便于断言的公式文本。</returns>
    private static string NormalizeNameFormula(string formula) => (formula ?? string.Empty)
        .Replace("'明细'!", "明细!", System.StringComparison.Ordinal)
        .Replace("'Lookup'!", "Lookup!", System.StringComparison.Ordinal);

    /// <summary>
    /// 创建包含一个有效行和一个转换失败行的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbook(bool includeFreezePane = false)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.SetColumnWidth(0, 18 * 256);
        sheet.SetColumnHidden(1, true);
        var header = sheet.CreateRow(0);
        header.HeightInPoints = 30f;
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.HeightInPoints = 24f;
        failure.ZeroHeight = true;
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");
        if (includeFreezePane)
            sheet.CreateFreezePane(1, 1, 1, 1);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含表头粗体和失败行数字格式的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateStyledWorkbook()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        var codeHeader = header.CreateCell(0);
        codeHeader.SetCellValue("Code");
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerFont.IsItalic = true;
        headerFont.IsStrikeout = true;
        headerFont.Underline = FontUnderlineType.Single;
        headerFont.FontName = "Arial";
        headerFont.FontHeightInPoints = 14;
        headerFont.Color = IndexedColors.Red.Index;
        var headerStyle = workbook.CreateCellStyle();
        headerStyle.SetFont(headerFont);
        headerStyle.Alignment = HorizontalAlignment.Center;
        headerStyle.VerticalAlignment = VerticalAlignment.Center;
        headerStyle.WrapText = true;
        headerStyle.ShrinkToFit = true;
        headerStyle.Rotation = 15;
        headerStyle.Indention = 2;
        headerStyle.FillPattern = FillPattern.SolidForeground;
        headerStyle.FillForegroundColor = IndexedColors.Yellow.Index;
        headerStyle.FillBackgroundColor = IndexedColors.Blue.Index;
        headerStyle.BorderLeft = BorderStyle.Thin;
        headerStyle.LeftBorderColor = IndexedColors.Green.Index;
        headerStyle.BorderBottom = BorderStyle.Medium;
        headerStyle.BottomBorderColor = IndexedColors.Blue.Index;
        headerStyle.IsLocked = false;
        codeHeader.CellStyle = headerStyle;
        var quantityHeader = header.CreateCell(1);
        quantityHeader.SetCellValue("Quantity");
        var quantityHeaderFont = workbook.CreateFont();
        quantityHeaderFont.FontName = "宋体";
        quantityHeaderFont.FontHeightInPoints = 12;
        var quantityHeaderStyle = workbook.CreateCellStyle();
        quantityHeaderStyle.SetFont(quantityHeaderFont);
        quantityHeader.CellStyle = quantityHeaderStyle;

        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);

        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        var quantity = failure.CreateCell(1);
        quantity.SetCellValue("wrong");
        var quantityStyle = workbook.CreateCellStyle();
        quantityStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
        quantityStyle.IsHidden = true;
        quantity.CellStyle = quantityStyle;

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含完整展示状态的 XSSF 工作簿。
    /// </summary>
    /// <returns>XLSX 工作簿字节。</returns>
    private static byte[] CreateXlsxStyledWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.SetColumnWidth(0, 22 * 256);
        sheet.SetColumnWidth(1, 16 * 256);
        sheet.SetColumnHidden(1, true);
        sheet.SetColumnHidden(2, true);
        var header = sheet.CreateRow(0);
        var codeHeader = header.CreateCell(0);
        codeHeader.SetCellValue("Code");
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerFont.IsItalic = true;
        headerFont.IsStrikeout = true;
        headerFont.Underline = FontUnderlineType.Single;
        headerFont.FontName = "Arial";
        headerFont.FontHeightInPoints = 14;
        headerFont.Color = IndexedColors.Red.Index;
        var headerStyle = workbook.CreateCellStyle();
        headerStyle.SetFont(headerFont);
        headerStyle.Alignment = HorizontalAlignment.Center;
        headerStyle.VerticalAlignment = VerticalAlignment.Center;
        headerStyle.WrapText = true;
        headerStyle.ShrinkToFit = true;
        headerStyle.Rotation = 15;
        headerStyle.Indention = 2;
        headerStyle.FillPattern = FillPattern.SolidForeground;
        headerStyle.FillForegroundColor = IndexedColors.Yellow.Index;
        headerStyle.FillBackgroundColor = IndexedColors.Blue.Index;
        headerStyle.BorderLeft = BorderStyle.Thin;
        headerStyle.LeftBorderColor = IndexedColors.Green.Index;
        headerStyle.BorderBottom = BorderStyle.Medium;
        headerStyle.BottomBorderColor = IndexedColors.Blue.Index;
        headerStyle.IsLocked = false;
        codeHeader.CellStyle = headerStyle;
        var quantityHeader = header.CreateCell(1);
        quantityHeader.SetCellValue("Quantity");
        var quantityHeaderFont = workbook.CreateFont();
        quantityHeaderFont.FontName = "宋体";
        quantityHeaderFont.FontHeightInPoints = 12;
        var quantityHeaderStyle = workbook.CreateCellStyle();
        quantityHeaderStyle.SetFont(quantityHeaderFont);
        quantityHeader.CellStyle = quantityHeaderStyle;

        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);

        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        var quantity = failure.CreateCell(1);
        quantity.SetCellValue("wrong");
        var quantityStyle = workbook.CreateCellStyle();
        quantityStyle.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0.00");
        quantityStyle.IsHidden = true;
        quantity.CellStyle = quantityStyle;

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含 HSSF 条件格式的工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithConditionalFormatting()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var failure = sheet.CreateRow(1);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var formatting = sheet.SheetConditionalFormatting;
        var rule = formatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan,
            "0");
        var pattern = rule.CreatePatternFormatting();
        pattern.FillPattern = FillPattern.SolidForeground;
        pattern.FillBackgroundColor = IndexedColors.Yellow.Index;
        formatting.AddConditionalFormatting(new[] { new CellRangeAddress(1, 1, 1, 1) }, rule);

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含同一区域两条条件格式规则的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithMultipleConditionalFormatting()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var failure = sheet.CreateRow(1);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var formatting = sheet.SheetConditionalFormatting;
        var greaterRule = formatting.CreateConditionalFormattingRule(
            ComparisonOperator.GreaterThan, "0");
        var greaterPattern = greaterRule.CreatePatternFormatting();
        greaterPattern.FillPattern = FillPattern.SolidForeground;
        greaterPattern.FillBackgroundColor = IndexedColors.Yellow.Index;
        var lessRule = formatting.CreateConditionalFormattingRule(
            ComparisonOperator.LessThan, "10");
        var lessPattern = lessRule.CreatePatternFormatting();
        lessPattern.FillPattern = FillPattern.SolidForeground;
        lessPattern.FillBackgroundColor = IndexedColors.Red.Index;
        formatting.AddConditionalFormatting(new[] { new CellRangeAddress(1, 1, 1, 1) },
            greaterRule, lessRule);

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含复杂公式条件格式的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithComplexConditionalFormatting()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var failure = sheet.CreateRow(1);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var formatting = sheet.SheetConditionalFormatting;
        var rule = formatting.CreateConditionalFormattingRule("SUM(B2)>0");
        formatting.AddConditionalFormatting(new[] { new CellRangeAddress(1, 1, 1, 1) }, rule);

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含 XSSF 条件格式的工作簿。
    /// </summary>
    /// <returns>XLSX 工作簿字节。</returns>
    private static byte[] CreateXlsxWorkbookWithConditionalFormatting(
        bool includeSuccessfulRow = false)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        if (includeSuccessfulRow)
        {
            var valid = sheet.CreateRow(1);
            valid.CreateCell(0).SetCellValue("A");
            valid.CreateCell(1).SetCellValue(1);
        }
        var failureRow = includeSuccessfulRow ? 2 : 1;
        var failure = sheet.CreateRow(failureRow);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var formatting = sheet.SheetConditionalFormatting;
        var rule = formatting.CreateConditionalFormattingRule(ComparisonOperator.GreaterThan,
            "0");
        formatting.AddConditionalFormatting(new[] { new CellRangeAddress(
            includeSuccessfulRow ? 1 : 1, failureRow, 1, 1) }, rule);

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含原生整数校验和转换失败行的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithNativeDataValidation(
        bool includeSuccessfulRowInValidation = false)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateintConstraint(0, "1", "10");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(includeSuccessfulRowInValidation ? 1 : 2, 2, 1, 1)));

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含失败行内命名列表来源和转换失败值的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithNamedListDataValidation(bool localScope = false)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var name = workbook.CreateName();
        name.NameName = "Allowed";
        if (localScope)
            name.SheetIndex = workbook.GetSheetIndex(sheet);
        name.RefersToFormula = "'明细'!$A$3";
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateFormulaListConstraint("Allowed");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(2, 2, 1, 1)));

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建命名列表来源位于未请求工作表的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithCrossSheetNamedListDataValidation(
        SheetVisibility? lookupVisibility = null)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var lookup = workbook.CreateSheet("Lookup");
        if (lookupVisibility.HasValue)
            workbook.SetSheetVisibility(workbook.GetSheetIndex(lookup), lookupVisibility.Value);
        lookup.CreateRow(0).CreateCell(0).SetCellValue("Allowed");
        lookup.CreateRow(1).CreateCell(0).SetCellValue("Open");
        lookup.CreateRow(2).CreateCell(0).SetCellValue("Closed");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var name = workbook.CreateName();
        name.NameName = "Allowed";
        name.RefersToFormula = "'Lookup'!$A$2:$A$3";
        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateFormulaListConstraint("Allowed");
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(2, 2, 1, 1)));

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含显式列表校验和转换失败值的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithExplicitListDataValidation()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");

        var helper = sheet.GetDataValidationHelper();
        var constraint = helper.CreateExplicitListConstraint(new[] { "A,1", "B", "中文" });
        sheet.AddValidationData(helper.CreateValidation(constraint,
            new CellRangeAddressList(2, 2, 1, 1)));

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含表头、失败行及跨成功行合并区域的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithMerges()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));

        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);

        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");
        sheet.AddMergedRegion(new CellRangeAddress(2, 2, 0, 1));

        sheet.AddMergedRegion(new CellRangeAddress(1, 2, 2, 3));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含安全和不可完整保留定义名称的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateWorkbookWithDefinedNames()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = sheet.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = sheet.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");
        var secondFailure = sheet.CreateRow(3);
        secondFailure.CreateCell(0).SetCellValue("C");
        secondFailure.CreateCell(1).SetCellValue("wrong");

        var headerName = workbook.CreateName();
        headerName.NameName = "HeaderCell";
        headerName.RefersToFormula = "'明细'!$A$1";
        var failureName = workbook.CreateName();
        failureName.NameName = "FailureCell";
        failureName.RefersToFormula = "'明细'!$A$3";
        var localFailureName = workbook.CreateName();
        localFailureName.NameName = "LocalFailureCell";
        localFailureName.SheetIndex = workbook.GetSheetIndex(sheet);
        localFailureName.RefersToFormula = "'明细'!$A$3";
        var failureColumnName = workbook.CreateName();
        failureColumnName.NameName = "FailureColumn";
        failureColumnName.RefersToFormula = "'明细'!$A$3:$A$4";
        var successfulRowsName = workbook.CreateName();
        successfulRowsName.NameName = "SuccessfulRows";
        successfulRowsName.RefersToFormula = "'明细'!$A$2:$A$3";

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含固定单元格转换错误和明细行错误的 HSSF 工作簿。
    /// </summary>
    /// <returns>HSSF 工作簿字节。</returns>
    private static byte[] CreateFixedCellFailureWorkbook()
    {
        using var workbook = new HSSFWorkbook();
        var document = workbook.CreateSheet("单据");
        document.CreateRow(0).CreateCell(0).SetCellValue("wrong");
        document.GetRow(0).CreateCell(1).SetCellValue("采购单");
        var detail = workbook.CreateSheet("明细");
        var header = detail.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var valid = detail.CreateRow(1);
        valid.CreateCell(0).SetCellValue("A");
        valid.CreateCell(1).SetCellValue(1);
        var failure = detail.CreateRow(2);
        failure.CreateCell(0).SetCellValue("B");
        failure.CreateCell(1).SetCellValue("wrong");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 保存失败工作簿测试单据。
    /// </summary>
    private sealed class FailureDocument
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<FailureLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存失败工作簿测试明细。
    /// </summary>
    private sealed class FailureLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
