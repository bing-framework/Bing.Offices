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
/// NPOI Forward 分组小计边界测试。
/// </summary>
public sealed class EntityForwardGroupSubtotalTest
{
    /// <summary>
    /// 验证 XLS/XLSX 的同步和异步 Forward 导入会跳过分组小计与空行。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntityBatches_Forward_ShouldReadGroupSubtotal(
        bool xls, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls));
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
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
    }

    /// <summary>
    /// 验证分组小计 marker 重复出现在小计占用区域时返回 Plan 配置错误。
    /// </summary>
    /// <param name="xls">是否使用 XLS/HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_Forward_ShouldRejectGroupMarkerInSubtotalBody(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(xls, invalidBody: true));

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

    private static byte[] CreateWorkbook(bool xls, bool invalidBody = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Group");
        header.CreateCell(1).SetCellValue("Code");
        header.CreateCell(2).SetCellValue("Quantity");
        CreateLine(sheet, 1, "A", "A-1", 1);
        CreateLine(sheet, 2, "A", "A-2", 2);
        CreateGroupSubtotal(sheet, 4);
        if (invalidBody)
            sheet.CreateRow(5).CreateCell(0).SetCellValue("GROUP");
        CreateLine(sheet, 6, "B", "B-1", 3);
        CreateLine(sheet, 7, "B", "B-2", 4);
        CreateGroupSubtotal(sheet, 9);
        var footerRow = 12;
        sheet.CreateRow(footerRow).CreateCell(0).SetCellValue("END");

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static void CreateGroupSubtotal(ISheet sheet, int rowIndex)
    {
        sheet.CreateRow(rowIndex).CreateCell(0).SetCellValue("GROUP");
        sheet.CreateRow(rowIndex + 1).CreateCell(1).SetCellValue(3d);
    }

    private static void CreateLine(ISheet sheet, int rowIndex, string group, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(group);
        row.CreateCell(1).SetCellValue(code);
        row.CreateCell(2).SetCellValue(quantity);
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
