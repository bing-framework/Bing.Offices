using System;
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
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI Forward 最终 Footer 边界测试。
/// </summary>
public sealed class EntityForwardFooterTest
{
    /// <summary>
    /// 验证 XLS/XLSX 的同步和异步 Forward 导入会跳过 GapRows 与 Footer。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadFinalFooter(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<FooterDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 2, 3 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(new[] { 2, 3 }, batches.Select(batch => batch.FirstRowIndex).ToArray());
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
    }

    /// <summary>
    /// 验证空明细在尾部标记有效时可以成功导入。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldAllowEmptyDetails(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, empty: true));

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Equal(0, summary.RowsRead);
    }

    /// <summary>
    /// 验证最终 Footer 超出显式区域结束地址时在读取前返回配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectFooterOverflow(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(end: "B5"),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("超出结束地址", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证最终 Footer 的 marker 缺失、重复和过早位置均返回结构化配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="mode">异常场景。</param>
    [Theory]
    [InlineData(true, FooterFailure.Missing)]
    [InlineData(true, FooterFailure.Duplicate)]
    [InlineData(true, FooterFailure.Early)]
    [InlineData(false, FooterFailure.Missing)]
    [InlineData(false, FooterFailure.Duplicate)]
    [InlineData(false, FooterFailure.Early)]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidFooter(
        bool xls, FooterFailure mode)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, failure: mode));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains(mode switch
        {
            FooterFailure.Missing => "缺少尾部标记",
            FooterFailure.Duplicate => "尾部标记重复",
            _ => "位于明细起始位置之前"
        }, exception.Message, StringComparison.Ordinal);
    }

    private static ExcelEntityLayout<FooterDocument> CreateLayout(string end = null) =>
        ExcelEntity.Layout<FooterDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region =>
            {
                if (end != null)
                    region.End(end);
                region.Footer("END", footer => footer
                    .GapRows(1)
                    .Cell("B2", items => items.Sum(item => item.Quantity)));
            }));

    private static byte[] CreateWorkbook(bool xls, bool empty = false,
        FooterFailure failure = FooterFailure.None)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        if (!empty && failure != FooterFailure.Early)
        {
            CreateLine(sheet, 1, "A-1", 2);
            CreateLine(sheet, 2, "B-2", 3);
        }

        switch (failure)
        {
            case FooterFailure.Missing:
                break;
            case FooterFailure.Early:
                sheet.CreateRow(1).CreateCell(0).SetCellValue("END");
                break;
            case FooterFailure.Duplicate:
                CreateFooter(sheet, 4);
                sheet.CreateRow(6).CreateCell(0).SetCellValue("END");
                break;
            default:
                CreateFooter(sheet, empty ? 2 : 4);
                break;
        }

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static void CreateFooter(ISheet sheet, int rowIndex)
    {
        sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("END");
        sheet.CreateRow(rowIndex + 1).CreateCell(1).SetCellValue(5);
    }

    private static void CreateLine(ISheet sheet, int rowIndex, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    public enum FooterFailure
    {
        None,
        Missing,
        Duplicate,
        Early
    }

    private sealed class FooterDocument
    {
        public List<FooterLine> Lines { get; } = new();
    }

    private sealed class FooterLine
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
    }
}
