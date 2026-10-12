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
/// 验证 ClosedXML Forward 固定单元格错误的低保留失败工作簿输出。
/// </summary>
public sealed class EntityForwardFixedCellFailureWorkbookTest
{
    /// <summary>
    /// 验证固定单元格错误与列表错误按源行写入失败工作簿。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ErrorRowsOnly_ShouldRetainFixedCellFailure(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
        Assert.Contains(summary.Result.Errors, error =>
            error.SheetName == "单据" && error.RowIndex == 1 && error.ColumnIndex == 1
            && error.PropertyName == nameof(FailureDocument.Number));
        Assert.Contains(summary.Result.Errors, error =>
            error.SheetName == "明细" && error.RowIndex == 2 && error.ColumnIndex == 2
            && error.PropertyName == nameof(FailureLine.Quantity));
        Assert.False(summary.HasCompleteResult);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var document = workbook.Worksheet("单据");
        Assert.Equal("wrong", document.Cell("A1").GetString());
        Assert.Equal("采购单", document.Cell("B1").GetString());

        var details = workbook.Worksheet("明细");
        Assert.Equal("Code", details.Cell("A1").GetString());
        Assert.Equal("Quantity", details.Cell("B1").GetString());
        Assert.Equal("B", details.Cell("A2").GetString());
        Assert.Equal("wrong", details.Cell("B2").GetString());

        var errors = workbook.Worksheet("_ImportErrors");
        Assert.True(errors.RowsUsed().Count() >= 3);
    }

    /// <summary>
    /// 创建只输出错误行的实体导入选项。
    /// </summary>
    /// <param name="destination">失败工作簿目标流。</param>
    /// <returns>实体批次导入选项。</returns>
    private static ExcelEntityImportOptions CreateImportOptions(Stream destination) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            });

    /// <summary>
    /// 创建包含固定单元格和明细列表的实体布局。
    /// </summary>
    /// <returns>实体布局。</returns>
    private static ExcelEntityLayout<FailureDocument> CreateLayout() =>
        ExcelEntity.Layout<FailureDocument>(builder => builder
            .Cell("单据", "A1", document => document.Number)
            .Cell("单据", "B1", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建固定单元格和明细字段均包含转换错误的工作簿。
    /// </summary>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateWorkbook()
    {
        using var workbook = new XLWorkbook();
        var document = workbook.AddWorksheet("单据");
        document.Cell("A1").Value = "wrong";
        document.Cell("B1").Value = "采购单";
        var details = workbook.AddWorksheet("明细");
        details.Cell("A1").Value = "Code";
        details.Cell("B1").Value = "Quantity";
        details.Cell("A2").Value = "B";
        details.Cell("B2").Value = "wrong";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
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
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

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
