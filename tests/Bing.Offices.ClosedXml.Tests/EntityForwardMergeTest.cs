using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML Forward 合并区域预检测试。
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    /// 创建包含实体列表和可选合并区域的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(bool includeMerge)
    {
        using var workbook = new XLWorkbook();
        var metadata = workbook.AddWorksheet("元数据");
        metadata.Cell("A1").Value = "采购单";
        if (includeMerge)
            metadata.Range("A1:B1").Merge();

        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Quantity";
        sheet.Cell("A2").Value = "A";
        sheet.Cell("B2").Value = 2;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
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
