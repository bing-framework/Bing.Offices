using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 分组小计与分页组合测试。
/// </summary>
public sealed class EntityGroupPaginationCombinationTest
{
    /// <summary>
    /// 验证分页容量仅统计明细行，跨页分组小计不会改变回读行号。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExportAndImport_GroupPagination_ShouldPreserveDetailsAndPhysicalRows(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: true);
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("A", "A3", 3m),
            ("B", "B1", 4m), ("B", "B2", 5m), ("C", "C1", 6m));
        var bytes = Export(document, layout, xls);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(new[]
            {
                "Group", "A", "A", "PAGE", "A", "GROUP", "B", "PAGE", "B", "GROUP", "C", "GROUP", "END"
            }, ReadFirstColumn(sheet));
            Assert.Equal(new[] { 3, 7 }, sheet.RowBreaks);
            Assert.Equal(3d, sheet.GetRow(3).GetCell(1).NumericCellValue, 2);
            Assert.Equal(6d, sheet.GetRow(5).GetCell(1).NumericCellValue, 2);
            Assert.Equal(7d, sheet.GetRow(7).GetCell(1).NumericCellValue, 2);
            Assert.Equal(9d, sheet.GetRow(9).GetCell(1).NumericCellValue, 2);
            Assert.Equal(6d, sheet.GetRow(11).GetCell(1).NumericCellValue, 2);
            Assert.Equal(21d, sheet.GetRow(12).GetCell(1).NumericCellValue, 2);
        }

        using var source = new MemoryStream(bytes, writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GroupPageDocument>>();
        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(6, summary.RowsRead);
        Assert.Equal(6, summary.SucceededRows);
        Assert.Equal(new[] { 2, 3, 5, 7, 9, 11 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { "A1", "A2", "A3", "B1", "B2", "C1" }, batches.Select(batch =>
            Assert.IsType<GroupPageLine>(Assert.Single(batch.Items)).Code));
        Assert.Empty(summary.Result.Errors);
    }

    /// <summary>
    /// 验证组结束与分页重合时先写组小计，再写分页小计。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Export_GroupAndPageSubtotalAtSameBoundary_ShouldWriteGroupFirst(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: true);
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m),
            ("B", "B1", 3m), ("B", "B2", 4m), ("C", "C1", 5m));
        var bytes = Export(document, layout, xls);

        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        var sheet = workbook.GetSheet("明细");
        Assert.Equal("GROUP", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal(3d, sheet.GetRow(3).GetCell(1).NumericCellValue, 2);
        Assert.Equal("PAGE", sheet.GetRow(4).GetCell(0).StringCellValue);
        Assert.Equal(3d, sheet.GetRow(4).GetCell(1).NumericCellValue, 2);
        Assert.Equal("GROUP", sheet.GetRow(7).GetCell(0).StringCellValue);
        Assert.Equal(7d, sheet.GetRow(7).GetCell(1).NumericCellValue, 2);
        Assert.Equal("PAGE", sheet.GetRow(8).GetCell(0).StringCellValue);
        Assert.Equal(7d, sheet.GetRow(8).GetCell(1).NumericCellValue, 2);
        Assert.Equal("END", sheet.GetRow(11).GetCell(0).StringCellValue);
        Assert.Equal(15d, sheet.GetRow(11).GetCell(1).NumericCellValue, 2);
        Assert.Equal(new[] { 4, 8 }, sheet.RowBreaks);
    }

    /// <summary>
    /// 验证未配置分页小计时仍按明细容量分页并跳过分组小计。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExportAndImport_GroupPaginationWithoutPageSubtotal_ShouldKeepDetailRows(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: false);
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("A", "A3", 3m),
            ("B", "B1", 4m), ("B", "B2", 5m));
        var bytes = Export(document, layout, xls);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(new[] { "Group", "A", "A", "A", "GROUP", "B", "B", "GROUP", "END" },
                ReadFirstColumn(sheet));
            Assert.Equal(2, sheet.RowBreaks.Length);
        }

        using var source = new MemoryStream(bytes, writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GroupPageDocument>>();
        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { 2, 3, 4, 6, 7 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { "A1", "A2", "A3", "B1", "B2" }, batches.Select(batch =>
            Assert.IsType<GroupPageLine>(Assert.Single(batch.Items)).Code));
    }

    /// <summary>
    /// 验证空明细不生成分页或分组小计，只保留最终尾部。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExportAndImport_EmptyDetails_ShouldWriteOnlyFinalFooter(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: true);
        var bytes = Export(new GroupPageDocument(), layout, xls);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(new[] { "Group", "END" }, ReadFirstColumn(sheet));
            Assert.Empty(sheet.RowBreaks);
            Assert.Equal(0d, sheet.GetRow(1).GetCell(1).NumericCellValue, 2);
        }

        using var source = new MemoryStream(bytes, writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GroupPageDocument>>();
        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(0, summary.RowsRead);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(batches);
    }

    /// <summary>
    /// 验证间隔行和多行小计保持工作表坐标及导入行号。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ExportAndImport_GroupPaginationWithMultiRowFooters_ShouldKeepCoordinates(bool xls)
    {
        var layout = ExcelEntity.Layout<GroupPageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region
                .GroupSubtotal(line => line.Group, "GROUP", footer => footer.GapRows(1)
                    .Cell("B2", items => items.Sum(item => item.Amount)))
                .PageBreak(2)
                .PageSubtotal("PAGE", footer => footer.GapRows(1)
                    .Cell("B2", items => items.Sum(item => item.Amount)))
                .Footer("END", footer => footer.GapRows(1))));
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("B", "B1", 3m));
        var bytes = Export(document, layout, xls);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("GROUP", sheet.GetRow(4).GetCell(0).StringCellValue);
            Assert.Equal(3d, sheet.GetRow(5).GetCell(1).NumericCellValue, 2);
            Assert.Equal("PAGE", sheet.GetRow(7).GetCell(0).StringCellValue);
            Assert.Equal(3d, sheet.GetRow(8).GetCell(1).NumericCellValue, 2);
            Assert.Equal("GROUP", sheet.GetRow(11).GetCell(0).StringCellValue);
            Assert.Equal("END", sheet.GetRow(14).GetCell(0).StringCellValue);
            Assert.Equal(new[] { 8 }, sheet.RowBreaks);
        }

        using var source = new MemoryStream(bytes, writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var batches = new List<ExcelEntityImportBatch<GroupPageDocument>>();
        var result = provider.GetRequiredService<IExcelEntityBatchImporter>().ImportEntityBatches(source,
            layout, new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(result.IsSuccess, string.Join(";", result.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { 2, 3, 10 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { "A1", "A2", "B1" }, result.Result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证组合导入在页小计标记缺失时返回 Plan 阶段配置错误。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Import_GroupPaginationWithMissingPageMarker_ShouldRejectWorkbook(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: true);
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("A", "A3", 3m));
        var bytes = RewriteWorkbook(Export(document, layout, xls), sheet =>
        {
            var marker = Enumerable.Range(0, sheet.LastRowNum + 1)
                .Select(index => sheet.GetRow(index)?.GetCell(0))
                .First(cell => cell?.CellType == CellType.String && cell.StringCellValue == "PAGE");
            marker.SetCellValue("NOT_PAGE");
        });
        using var source = new MemoryStream(bytes, writable: false);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(source, layout));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
    }

    /// <summary>
    /// 验证组页同边界的标记顺序被调换时导入失败。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Import_GroupPaginationWithReversedMarkers_ShouldRejectWorkbook(bool xls)
    {
        var layout = CreateLayout(includePageSubtotal: true);
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("B", "B1", 3m));
        var bytes = RewriteWorkbook(Export(document, layout, xls), sheet =>
        {
            sheet.GetRow(3).GetCell(0).SetCellValue("PAGE");
            sheet.GetRow(4).GetCell(0).SetCellValue("GROUP");
        });
        using var source = new MemoryStream(bytes, writable: false);

        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(source, layout));

        Assert.Equal(BingOfficesStage.Plan, error.Stage);
    }

    /// <summary>
    /// 验证组合小计推移明细后会预检声明区域边界。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Export_GroupPaginationBeyondRegionEnd_ShouldFailBeforeWriting(bool xls)
    {
        var layout = ExcelEntity.Layout<GroupPageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.End("C6")
                .GroupSubtotal(line => line.Group, "GROUP")
                .PageBreak(2).PageSubtotal("PAGE").Footer("END")));
        var document = CreateDocument(("A", "A1", 1m), ("A", "A2", 2m), ("B", "B1", 3m));

        var error = Assert.Throws<BingOfficesConfigurationException>(() => Export(document, layout, xls));

        Assert.Equal(BingOfficesStage.Plan, error.Stage);
    }

    /// <summary>
    /// 验证组合中间小计拒绝跨越插入行的连续求和公式。
    /// </summary>
    [Fact]
    public void Layout_GroupPagination_ShouldRejectContiguousIntermediateFormula()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<GroupPageDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines, region => region
                .GroupSubtotal(line => line.Group, "GROUP")
                .PageBreak(2).PageSubtotal("PAGE", footer => footer
                    .FormulaSumContiguousRowsAbove("B1", "B")))));
    }

    private static ExcelEntityLayout<GroupPageDocument> CreateLayout(bool includePageSubtotal) =>
        ExcelEntity.Layout<GroupPageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region =>
            {
                region.PageBreak(2);
                region.GroupSubtotal(line => line.Group, "GROUP",
                    footer => footer.Cell("B1", items => items.Sum(item => item.Amount)));
                if (includePageSubtotal)
                    region.PageSubtotal("PAGE", footer => footer.Cell("B1",
                        items => items.Sum(item => item.Amount)));
                region.Footer("END", footer => footer.Cell("B1",
                    items => items.Sum(item => item.Amount)));
            }));

    private static GroupPageDocument CreateDocument(params (string Group, string Code, decimal Amount)[] rows) =>
        new()
        {
            Lines = rows.Select(row => new GroupPageLine
            {
                Group = row.Group,
                Code = row.Code,
                Amount = row.Amount
            }).ToList()
        };

    private static byte[] Export(GroupPageDocument document,
        ExcelEntityLayout<GroupPageDocument> layout, bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("明细");
        var template = new MemoryStream();
        workbook.Write(template, true);
        template.Position = 0;
        using (template)
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(document, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            return output.ToArray();
        }
    }

    private static byte[] RewriteWorkbook(byte[] bytes, Action<ISheet> edit)
    {
        using var source = new MemoryStream(bytes, writable: false);
        using var workbook = WorkbookFactory.Create(source);
        edit(workbook.GetSheet("明细"));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private static string[] ReadFirstColumn(ISheet sheet) =>
        Enumerable.Range(0, sheet.LastRowNum + 1)
            .Select(index => sheet.GetRow(index)?.GetCell(0)?.ToString() ?? string.Empty).ToArray();

    /// <summary>
    /// 保存分组分页组合测试根模型。
    /// </summary>
    public sealed class GroupPageDocument
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<GroupPageLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 保存分组分页组合测试明细。
    /// </summary>
    public sealed class GroupPageLine
    {
        /// <summary>
        /// 获取或设置分组键。
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 获取或设置行编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }
}
