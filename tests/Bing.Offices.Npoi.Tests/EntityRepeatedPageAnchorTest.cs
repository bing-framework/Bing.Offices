using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证实体列表分页小计的逐页命名锚点。
/// </summary>
public sealed class EntityRepeatedPageAnchorTest
{
    /// <summary>
    /// 验证每个中间页的名称均指向实际小计标记，且导入跳过小计行。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageSubtotalNamed_MultiplePages_ShouldCreateSequentialNamesAndRoundTrip(bool xls)
    {
        var layout = CreateLayout();
        var source = CreateDocument("A", "A", "B", "B", "C");
        var bytes = Export(source, layout, xls);

        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            AssertNames(workbook, ("PageTotal_001", "$A$4"), ("PageTotal_002", "$A$7"));
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("PAGE", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.Equal("PAGE", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal(new[] { 3, 6 }, sheet.RowBreaks);
            Assert.Equal("END", sheet.GetRow(8).GetCell(0).StringCellValue);
        }

        using var input = new MemoryStream(bytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证同一边界先写分组小计，再将名称定位到分页小计。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageSubtotalNamed_WithGroupSubtotal_ShouldFollowGroupMarker(bool xls)
    {
        var layout = CreateLayout(grouped: true);
        var source = CreateDocument("A", "A", "B", "B", "C");
        var bytes = Export(source, layout, xls);

        using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
        AssertNames(workbook, ("PageTotal_001", "$A$5"), ("PageTotal_002", "$A$9"));
        var sheet = workbook.GetSheet("明细");
        Assert.Equal("GROUP", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal("PAGE", sheet.GetRow(4).GetCell(0).StringCellValue);
        Assert.Equal("GROUP", sheet.GetRow(7).GetCell(0).StringCellValue);
        Assert.Equal("PAGE", sheet.GetRow(8).GetCell(0).StringCellValue);
        Assert.Equal(new[] { 4, 8 }, sheet.RowBreaks);
    }

    /// <summary>
    /// 验证空明细和未满一页的明细都不创建分页名称。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageSubtotalNamed_EmptyOrSinglePage_ShouldNotCreateName(bool xls)
    {
        foreach (var source in new[] { CreateDocument(), CreateDocument("A") })
        {
            var bytes = Export(source, CreateLayout(), xls);
            using var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false));
            Assert.Empty(workbook.GetAllNames().Where(name => name.NameName.StartsWith("PageTotal_",
                StringComparison.OrdinalIgnoreCase)));
            Assert.Empty(workbook.GetSheet("明细").RowBreaks);
        }
    }

    /// <summary>
    /// 验证模板中的同表单格名称逐页重定位并保持原有作用域。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="localScope">名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void PageSubtotalNamed_TemplateNames_ShouldRelocateAndPreserveScope(bool xls,
        bool localScope)
    {
        using var template = CreateTemplate(xls, localScope,
            ("PageTotal_001", "$Z$20"), ("PageTotal_002", "$Z$21"));
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B", "B", "C"),
            CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        AssertNames(workbook, ("PageTotal_001", "$A$4"), ("PageTotal_002", "$A$7"));
        Assert.All(workbook.GetAllNames().Where(name => name.NameName.StartsWith("PageTotal_",
            StringComparison.OrdinalIgnoreCase)), name => Assert.Equal(localScope ? 1 : -1,
            name.SheetIndex));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证模板中过期的同系列名称在计划阶段失败且不改写目标流。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageSubtotalNamed_StaleTemplateName_ShouldPreserveOutput(bool xls)
    {
        using var template = CreateTemplate(xls, false, ("PageTotal_003", "$Z$22"));
        using var output = new MemoryStream();
        var original = new byte[] { 1, 2, 3, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B", "B", "C"),
                CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证前缀过长或与已有实体锚点冲突时无法构建布局。
    /// </summary>
    [Fact]
    public void PageSubtotalNamed_InvalidPrefix_ShouldRejectLayout()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2).PageSubtotalNamed(new string('P', 252), "PAGE"))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<PageDocument>(builder => builder
            .CellNamed("说明", "pagetotal_001", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2).PageSubtotalNamed("PageTotal", "PAGE"))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<PageDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2).PageSubtotalNamed("PageTotal", "PAGE")
                .FooterNamed("PageTotal_002", "END"))));
    }

    /// <summary>
    /// 验证跨 Sheet 占用及多格名称在计划阶段失败且不改写目标流。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="otherSheet">名称是否指向其他工作表。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void PageSubtotalNamed_InvalidTemplateName_ShouldPreserveOutput(bool xls,
        bool otherSheet)
    {
        using var template = CreateInvalidTemplate(xls, otherSheet);
        using var output = new MemoryStream();
        var original = new byte[] { 5, 6, 7 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B"),
                CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证异步模板导出也逐页创建正确的名称。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PageSubtotalNamed_AsyncExport_ShouldCreateSequentialNames(bool xls)
    {
        using var template = CreateTemplate(xls, false);
        using var output = new MemoryStream();
        await new NpoiExcelExporter().ExportForTemplateAsync(CreateDocument("A", "A", "B", "B", "C"),
            CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        AssertNames(workbook, ("PageTotal_001", "$A$4"), ("PageTotal_002", "$A$7"));
        Assert.True(template.CanRead);
    }

    private static ExcelEntityLayout<PageDocument> CreateLayout(bool grouped = false) =>
        ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region =>
            {
                region.PageBreak(2).PageSubtotalNamed("PageTotal", "PAGE",
                    footer => footer.Cell("B1", items => items.Sum(item => item.Amount)));
                if (grouped)
                    region.GroupSubtotal(line => line.Group, "GROUP");
                region.Footer("END");
            }));

    private static PageDocument CreateDocument(params string[] groups) => new()
    {
        Lines = groups.Select((group, index) => new PageLine
        {
            Group = group, Code = "L" + (index + 1), Amount = index + 1
        }).ToList()
    };

    private static byte[] Export(PageDocument source, ExcelEntityLayout<PageDocument> layout, bool xls)
    {
        using var template = CreateTemplate(xls, false);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        return output.ToArray();
    }

    private static MemoryStream CreateTemplate(bool xls, bool localScope,
        params (string Name, string Address)[] names)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("说明");
        var sheet = workbook.CreateSheet("明细");
        foreach (var (nameText, address) in names)
        {
            var name = workbook.CreateName();
            name.NameName = nameText;
            name.SheetIndex = localScope ? workbook.GetSheetIndex(sheet) : -1;
            name.RefersToFormula = "'明细'!" + address;
        }
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateInvalidTemplate(bool xls, bool otherSheet)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("说明");
        workbook.CreateSheet("明细");
        var name = workbook.CreateName();
        name.NameName = "PageTotal_001";
        name.RefersToFormula = otherSheet ? "'说明'!$Z$20" : "'明细'!$Z$20:$Z$21";
        var stream = new MemoryStream();
        workbook.Write(stream, true);
        stream.Position = 0;
        return stream;
    }

    private static void AssertNames(IWorkbook workbook,
        params (string Name, string Address)[] expected)
    {
        var names = workbook.GetAllNames().Where(name => name.NameName.StartsWith("PageTotal_",
            StringComparison.OrdinalIgnoreCase)).ToArray();
        Assert.Equal(expected.Length, names.Length);
        foreach (var (nameText, address) in expected)
        {
            var name = Assert.Single(names.Where(item => item.NameName == nameText));
            Assert.Equal("明细!" + address,
                name.RefersToFormula.Replace("'明细'!", "明细!", StringComparison.Ordinal));
        }
    }

    /// <summary>
    /// 保存分页锚点测试实体。
    /// </summary>
    public sealed class PageDocument
    {
        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取或设置明细。
        /// </summary>
        public List<PageLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 保存分页锚点测试明细。
    /// </summary>
    public sealed class PageLine
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
