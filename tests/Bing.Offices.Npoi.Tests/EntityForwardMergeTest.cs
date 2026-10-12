using System;
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
/// NPOI Forward 合并区域预检测试。
/// </summary>
public sealed class EntityForwardMergeTest
{
    /// <summary>
    /// 验证 XLSX Forward 同步和异步导入可读取已声明的合并区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldAcceptDeclaredMerge(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var layout = CreateLayout();
        using var source = new MemoryStream(CreateWorkbook(includeMerge: true), writable: false);
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("A", summary.Result.Entity.Lines[0].Code);
    }

    /// <summary>
    /// 验证 XLSX Forward 在源工作表缺少布局合并区域时于计划阶段失败。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectMissingDeclaredMerge()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(includeMerge: false), writable: false);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null,
                    ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("合并区域", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 可以读取投影后的合并区域。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadMerge(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithMerge(), writable: false);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null,
                    ExcelEntityBatchParsingMode.Forward), (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null,
                    ExcelEntityBatchParsingMode.Forward), _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("A", summary.Result.Entity.Lines[0].Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 在投影缺少声明合并区域时返回计划错误。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldRejectMissingMerge(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbookWithMerge(false), writable: false);
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var exception = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (_, _) => Task.CompletedTask))
            : Assert.Throws<BingOfficesConfigurationException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("合并区域", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建包含实体列表和可选合并区域的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(bool includeMerge)
    {
        using var workbook = new XSSFWorkbook();
        var metadata = workbook.CreateSheet("元数据");
        metadata.CreateRow(0).CreateCell(0).SetCellValue("采购单");
        if (includeMerge)
            metadata.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));

        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建带合并区域的 HSSF 工作簿，用于验证格式边界。
    /// </summary>
    private static byte[] CreateHssfWorkbookWithMerge(bool includeMerge = true)
    {
        using var workbook = new HSSFWorkbook();
        var metadata = workbook.CreateSheet("元数据");
        metadata.CreateRow(0).CreateCell(0).SetCellValue("采购单");
        if (includeMerge)
            metadata.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        sheet.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
        sheet.GetRow(1).CreateCell(1).SetCellValue(2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建声明元数据合并区域的 Forward 布局。
    /// </summary>
    private static ExcelEntityLayout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>
        CreateLayout() => ExcelEntity.Layout<EntityLayoutBatchImportTest.GetterOnlyBatchDocument>(builder => builder
            .Merge("元数据", "A1:B1")
            .ListRegion("明细", "A1", document => document.Lines));
}
