using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI Forward 固定单元格错误的低保留失败工作簿输出。
/// </summary>
public sealed class EntityForwardFixedCellFailureWorkbookTest
{
    /// <summary>
    /// 验证固定单元格错误与列表错误按源行写入 XLS 和 XLSX 失败工作簿。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ErrorRowsOnly_ShouldRetainFixedCellFailure(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls), writable: false);
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
        using var workbook = WorkbookFactory.Create(failure);
        var document = workbook.GetSheet("单据");
        Assert.NotNull(document);
        Assert.Equal("wrong", document.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("采购单", document.GetRow(0).GetCell(1).StringCellValue);

        var details = workbook.GetSheet("明细");
        Assert.NotNull(details);
        Assert.Equal("Code", details.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Quantity", details.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("B", details.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("wrong", details.GetRow(1).GetCell(1).StringCellValue);

        var errors = workbook.GetSheet("_ImportErrors");
        Assert.NotNull(errors);
        Assert.True(errors.LastRowNum >= 2);
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
    /// <param name="xls">是否生成 XLS 格式。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var document = workbook.CreateSheet("单据");
        var documentRow = document.CreateRow(0);
        documentRow.CreateCell(0).SetCellValue("wrong");
        documentRow.CreateCell(1).SetCellValue("采购单");
        var details = workbook.CreateSheet("明细");
        var header = details.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        var row = details.CreateRow(1);
        row.CreateCell(0).SetCellValue("B");
        row.CreateCell(1).SetCellValue("wrong");
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
