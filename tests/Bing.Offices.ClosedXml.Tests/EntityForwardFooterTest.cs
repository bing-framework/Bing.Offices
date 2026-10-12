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
/// ClosedXML Forward 最终 Footer 边界测试。
/// </summary>
public sealed class EntityForwardFooterTest
{
    /// <summary>
    /// 验证同步和异步 Forward 导入会跳过 GapRows 与 Footer。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadFinalFooter(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
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
    }

    /// <summary>
    /// 验证空明细在尾部标记有效时可以成功导入。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldAllowEmptyDetails()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(empty: true));

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
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectFooterOverflow()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout("B5"),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("超出结束地址", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证最终 Footer 的 marker 缺失、重复和过早位置均返回结构化配置错误。
    /// </summary>
    /// <param name="mode">异常场景。</param>
    [Theory]
    [InlineData(FooterFailure.Missing)]
    [InlineData(FooterFailure.Duplicate)]
    [InlineData(FooterFailure.Early)]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidFooter(FooterFailure mode)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(failure: mode));

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

    private static byte[] CreateWorkbook(bool empty = false,
        FooterFailure failure = FooterFailure.None)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        if (!empty && failure != FooterFailure.Early)
        {
            CreateLine(sheet, 2, "A-1", 2);
            CreateLine(sheet, 3, "B-2", 3);
        }

        switch (failure)
        {
            case FooterFailure.Missing:
                break;
            case FooterFailure.Early:
                sheet.Cell(2, 1).Value = "END";
                break;
            case FooterFailure.Duplicate:
                CreateFooter(sheet, 5);
                sheet.Cell(7, 1).Value = "END";
                break;
            default:
                CreateFooter(sheet, empty ? 3 : 5);
                break;
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static void CreateFooter(IXLWorksheet sheet, int rowNumber)
    {
        sheet.Cell(rowNumber, 1).Value = "END";
        sheet.Cell(rowNumber + 1, 2).Value = 5;
    }

    private static void CreateLine(IXLWorksheet sheet, int rowNumber, string code, int quantity)
    {
        sheet.Cell(rowNumber, 1).Value = code;
        sheet.Cell(rowNumber, 2).Value = quantity;
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
