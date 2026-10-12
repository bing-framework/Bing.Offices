using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 ClosedXML Forward 的实体重复页眉边界。
/// </summary>
public sealed class EntityForwardPageHeaderTest
{
    /// <summary>
    /// 验证同步和异步导入会跳过后续页眉并保留原始行号。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldSkipRepeatedHeaders(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<PageHeaderDocument>>();
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
        Assert.Equal(new[] { "A-1", "B-2", "C-3", "D-4" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 3, 4, 7, 8 },
            batches.Select(batch => batch.FirstRowIndex).ToArray());
        Assert.Equal(new[] { 2, 3, 6, 7 },
            summary.Result.Sheets[0].SourceRows.ToArray());
    }

    /// <summary>
    /// 验证重复页眉 marker 缺失时返回原始页眉坐标。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectMissingRepeatedHeaderMarker()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var bytes = CreateWorkbook();
        using (var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false)))
        {
            workbook.Worksheet("明细").Cell(5, 1).Value = "BROKEN";
            using var changed = new MemoryStream();
            workbook.SaveAs(changed);
            bytes = changed.ToArray();
        }

        using var source = new MemoryStream(bytes, writable: false);
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(5, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    /// <summary>
    /// 验证复制页列表标题被改动时返回复制页坐标。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectCopiedHeaderTitleMismatch()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var bytes = CreateWorkbook();
        using (var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false)))
        {
            workbook.Worksheet("明细").Cell(6, 1).Value = "Changed";
            using var changed = new MemoryStream();
            workbook.SaveAs(changed);
            bytes = changed.ToArray();
        }

        using var source = new MemoryStream(bytes, writable: false);
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(6, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
    }

    private static ExcelEntityLayout<PageHeaderDocument> CreateLayout() =>
        ExcelEntity.Layout<PageHeaderDocument>(builder => builder.ListRegion("明细", "A2",
            document => document.Lines, region => region
                .PageBreak(2)
                .PageHeader("A1:B2", "A1", "HEAD")
                .End("B8")));

    private static byte[] CreateWorkbook()
    {
        var document = new PageHeaderDocument
        {
            Lines = new List<PageHeaderLine>
            {
                new() { Code = "A-1", Quantity = 1 },
                new() { Code = "B-2", Quantity = 2 },
                new() { Code = "C-3", Quantity = 3 },
                new() { Code = "D-4", Quantity = 4 }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(document, CreateLayout(), output);
        return output.ToArray();
    }

    private sealed class PageHeaderDocument
    {
        public List<PageHeaderLine> Lines { get; set; } = new();
    }

    private sealed class PageHeaderLine
    {
        public string Code { get; set; }
        public int Quantity { get; set; }
    }
}
