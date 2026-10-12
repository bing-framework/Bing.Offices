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
/// ClosedXML Forward 分组小计边界测试。
/// </summary>
public sealed class EntityForwardGroupSubtotalTest
{
    /// <summary>
    /// 验证同步和异步 Forward 导入会跳过分组小计与空行。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadGroupSubtotal(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<GroupDocument>>();
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
    }

    /// <summary>
    /// 验证分组小计 marker 重复出现在小计占用区域时返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectGroupMarkerInSubtotalBody()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(invalidBody: true));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("分组小计标记位置", exception.Message, StringComparison.Ordinal);
    }

    private static ExcelEntityLayout<GroupDocument> CreateLayout() =>
        ExcelEntity.Layout<GroupDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .GroupSubtotal(item => item.Group, "GROUP", footer => footer.GapRows(1)
                    .Cell("B2", items => items.Sum(item => item.Quantity)))
                .Footer("END", footer => footer.GapRows(1))));

    private static byte[] CreateWorkbook(bool invalidBody = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Group";
        sheet.Cell(1, 2).Value = "Code";
        sheet.Cell(1, 3).Value = "Quantity";
        CreateLine(sheet, 2, "A", "A-1", 1);
        CreateLine(sheet, 3, "A", "A-2", 2);
        CreateGroupSubtotal(sheet, 5);
        if (invalidBody)
            sheet.Cell(6, 1).Value = "GROUP";
        CreateLine(sheet, 7, "B", "B-1", 3);
        CreateLine(sheet, 8, "B", "B-2", 4);
        CreateGroupSubtotal(sheet, 10);
        sheet.Cell(13, 1).Value = "END";

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    private static void CreateGroupSubtotal(IXLWorksheet sheet, int rowNumber)
    {
        sheet.Cell(rowNumber, 1).Value = "GROUP";
        sheet.Cell(rowNumber + 1, 2).Value = 3;
    }

    private static void CreateLine(IXLWorksheet sheet, int rowNumber, string group,
        string code, int quantity)
    {
        sheet.Cell(rowNumber, 1).Value = group;
        sheet.Cell(rowNumber, 2).Value = code;
        sheet.Cell(rowNumber, 3).Value = quantity;
    }

    private sealed class GroupDocument
    {
        public List<GroupLine> Lines { get; } = new();
    }

    private sealed class GroupLine
    {
        public string Group { get; set; }
        public string Code { get; set; }
        public int Quantity { get; set; }
    }
}
