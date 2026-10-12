using System;
using System.Collections.Generic;
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
/// ClosedXML Forward 分页小计边界测试。
/// </summary>
public sealed class EntityForwardPageSubtotalTest
{
    /// <summary>
    /// 验证同步和异步 Forward 导入会跳过分页小计与空行。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadPageSubtotal(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
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
        Assert.Equal(new[] { 2, 3, 7, 8 },
            batches.Select(batch => batch.FirstRowIndex).ToArray());
        Assert.Equal(4, summary.RowsRead);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
    }

    /// <summary>
    /// 验证分页小计缺失、错位和末页无明细时返回 Plan 阶段配置错误。
    /// </summary>
    /// <param name="failure">异常场景。</param>
    [Theory]
    [InlineData(PageFailure.Missing)]
    [InlineData(PageFailure.Early)]
    [InlineData(PageFailure.AfterLast)]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidPageSubtotal(PageFailure failure)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(failure));

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

    private static byte[] CreateWorkbook(PageFailure failure = PageFailure.None)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        CreateLine(sheet, 2, "A-1", 1);
        CreateLine(sheet, 3, "A-2", 2);

        if (failure == PageFailure.Early)
            sheet.Cell(4, 1).Value = "PAGE";
        else if (failure == PageFailure.AfterLast)
        {
            CreateSubtotal(sheet, 5);
            sheet.Cell(7, 1).Value = "PAGE";
        }
        else if (failure != PageFailure.Missing)
            CreateSubtotal(sheet, 5);

        if (failure == PageFailure.None)
        {
            CreateSubtotalValue(sheet, 6, 3);
            CreateLine(sheet, 7, "B-1", 3);
            CreateLine(sheet, 8, "B-2", 4);
            sheet.Cell(10, 1).Value = "END";
        }
        else if (failure == PageFailure.Missing)
        {
            CreateLine(sheet, 4, "B-1", 3);
            CreateLine(sheet, 5, "B-2", 4);
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static void CreateSubtotal(IXLWorksheet sheet, int rowNumber)
    {
        sheet.Cell(rowNumber, 1).Value = "PAGE";
    }

    private static void CreateSubtotalValue(IXLWorksheet sheet, int rowNumber, int value)
    {
        sheet.Cell(rowNumber, 2).Value = value;
    }

    private static void CreateLine(IXLWorksheet sheet, int rowNumber, string code, int quantity)
    {
        sheet.Cell(rowNumber, 1).Value = code;
        sheet.Cell(rowNumber, 2).Value = quantity;
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
