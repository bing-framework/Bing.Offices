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
/// NPOI Forward 分页小计边界测试。
/// </summary>
public sealed class EntityForwardPageSubtotalTest
{
    /// <summary>
    /// 验证 XLS/XLSX 的同步和异步 Forward 导入会跳过分页小计与空行。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageSubtotal(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<PageDocument>>();
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
        Assert.Equal(new[] { "A-1", "A-2", "B-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 1, 2, 3, 4 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(new[] { 2, 3, 7, 8 },
            batches.Select(batch => batch.FirstRowIndex).ToArray());
        Assert.Equal(4, summary.RowsRead);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
    }

    /// <summary>
    /// 验证分页小计缺失、错位和末页无明细时返回 Plan 阶段配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="failure">异常场景。</param>
    [Theory]
    [InlineData(true, PageFailure.Missing)]
    [InlineData(true, PageFailure.Early)]
    [InlineData(true, PageFailure.AfterLast)]
    [InlineData(false, PageFailure.Missing)]
    [InlineData(false, PageFailure.Early)]
    [InlineData(false, PageFailure.AfterLast)]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidPageSubtotal(
        bool xls, PageFailure failure)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, failure));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains(failure switch
        {
            PageFailure.Missing => "缺少分页小计标记",
            PageFailure.Early => "分页小计标记位置",
            _ => "分页小计之后缺少明细"
        }, exception.Message, StringComparison.Ordinal);
    }

    private static ExcelEntityLayout<PageDocument> CreateLayout() =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE", footer => footer.GapRows(1).Cell("B2",
                    items => items.Sum(item => item.Quantity)))
                .Footer("END", footer => footer.GapRows(1))));

    private static byte[] CreateWorkbook(bool xls, PageFailure failure = PageFailure.None)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateLine(sheet, 1, "A-1", 1);
        CreateLine(sheet, 2, "A-2", 2);

        if (failure == PageFailure.Early)
            sheet.CreateRow(3).CreateCell(0).SetCellValue("PAGE");
        else if (failure == PageFailure.AfterLast)
        {
            CreateSubtotal(sheet, 4);
            sheet.CreateRow(6).CreateCell(0).SetCellValue("PAGE");
        }
        else if (failure != PageFailure.Missing)
            CreateSubtotal(sheet, 4);

        if (failure == PageFailure.None)
        {
            CreateSubtotalValue(sheet, 5, 3);
            CreateLine(sheet, 6, "B-1", 3);
            CreateLine(sheet, 7, "B-2", 4);
            sheet.CreateRow(9).CreateCell(0).SetCellValue("END");
        }
        else if (failure == PageFailure.Missing)
        {
            CreateLine(sheet, 3, "B-1", 3);
            CreateLine(sheet, 4, "B-2", 4);
        }

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static void CreateSubtotal(ISheet sheet, int rowIndex)
    {
        sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("PAGE");
    }

    private static void CreateSubtotalValue(ISheet sheet, int rowIndex, int value)
    {
        sheet.CreateRow(rowIndex).CreateCell(1).SetCellValue((double)value);
    }

    private static void CreateLine(ISheet sheet, int rowIndex, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    public enum PageFailure
    {
        None,
        Missing,
        Early,
        AfterLast
    }

    private sealed class PageDocument
    {
        public List<PageLine> Lines { get; } = new();
    }

    private sealed class PageLine
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
    }
}
