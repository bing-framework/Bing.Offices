using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证实体列表分页小计的逐页命名锚点。
/// </summary>
public sealed class EntityRepeatedPageAnchorTest
{
    /// <summary>
    /// 验证每个中间页的名称均指向实际小计标记，且导入跳过小计行。
    /// </summary>
    [Fact]
    public void PageSubtotalNamed_MultiplePages_ShouldCreateSequentialNamesAndRoundTrip()
    {
        var layout = CreateLayout();
        var source = CreateDocument("A", "A", "B", "B", "C");
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);

        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
        {
            AssertNames(workbook, ("PageTotal_001", 4), ("PageTotal_002", 7));
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("PAGE", sheet.Cell("A4").GetString());
            Assert.Equal("PAGE", sheet.Cell("A7").GetString());
            Assert.Equal("END", sheet.Cell("A9").GetString());
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(source.Lines.Select(line => line.Code), result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证同一边界先写分组小计，再将名称定位到分页小计。
    /// </summary>
    [Fact]
    public void PageSubtotalNamed_WithGroupSubtotal_ShouldFollowGroupMarker()
    {
        var layout = CreateLayout(grouped: true);
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateDocument("A", "A", "B", "B", "C"),
            layout, output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        AssertNames(workbook, ("PageTotal_001", 5), ("PageTotal_002", 9));
        var sheet = workbook.Worksheet("明细");
        Assert.Equal("GROUP", sheet.Cell("A4").GetString());
        Assert.Equal("PAGE", sheet.Cell("A5").GetString());
        Assert.Equal("GROUP", sheet.Cell("A8").GetString());
        Assert.Equal("PAGE", sheet.Cell("A9").GetString());
    }

    /// <summary>
    /// 验证空明细和未满一页的明细都不创建分页名称。
    /// </summary>
    [Fact]
    public void PageSubtotalNamed_EmptyOrSinglePage_ShouldNotCreateName()
    {
        foreach (var source in new[] { CreateDocument(), CreateDocument("A") })
        {
            using var output = new MemoryStream();
            new ClosedXmlExcelExporter().ExportEntity(source, CreateLayout(), output);
            using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
            Assert.Empty(workbook.DefinedNames.ValidNamedRanges().Where(name => name.Name.StartsWith(
                "PageTotal_", StringComparison.OrdinalIgnoreCase)));
        }
    }

    /// <summary>
    /// 验证模板中的同表单格名称逐页重定位并保持原有作用域。
    /// </summary>
    /// <param name="localScope">名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageSubtotalNamed_TemplateNames_ShouldRelocateAndPreserveScope(bool localScope)
    {
        using var template = CreateTemplate(localScope,
            ("PageTotal_001", "Z20"), ("PageTotal_002", "Z21"));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B", "B", "C"),
            CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        AssertNames(workbook, ("PageTotal_001", 4), ("PageTotal_002", 7));
        var names = localScope
            ? workbook.Worksheet("明细").DefinedNames.ValidNamedRanges()
            : workbook.DefinedNames.ValidNamedRanges();
        Assert.Equal(2, names.Count(name => name.Name.StartsWith("PageTotal_",
            StringComparison.OrdinalIgnoreCase)));
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证模板中过期的同系列名称在计划阶段失败且不改写目标流。
    /// </summary>
    /// <param name="localScope">名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageSubtotalNamed_StaleTemplateName_ShouldPreserveOutput(bool localScope)
    {
        using var template = CreateTemplate(localScope, ("PageTotal_003", "Z22"));
        using var output = new MemoryStream();
        var original = new byte[] { 1, 2, 3, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B", "B", "C"),
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
    /// <param name="otherSheet">名称是否指向其他工作表。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageSubtotalNamed_InvalidTemplateName_ShouldPreserveOutput(bool otherSheet)
    {
        using var template = CreateInvalidTemplate(otherSheet);
        using var output = new MemoryStream();
        var original = new byte[] { 5, 6, 7 };
        output.Write(original, 0, original.Length);
        output.Position = 0;

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument("A", "A", "B"),
                CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证异步模板导出也逐页创建正确的名称。
    /// </summary>
    [Fact]
    public async Task PageSubtotalNamed_AsyncExport_ShouldCreateSequentialNames()
    {
        using var template = CreateTemplate(false);
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportForTemplateAsync(CreateDocument("A", "A", "B", "B", "C"),
            CreateLayout(), new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        AssertNames(workbook, ("PageTotal_001", 4), ("PageTotal_002", 7));
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

    private static MemoryStream CreateTemplate(bool localScope,
        params (string Name, string Address)[] names)
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("说明");
        var sheet = workbook.Worksheets.Add("明细");
        foreach (var (name, address) in names)
        {
            if (localScope)
                sheet.DefinedNames.Add(name, sheet.Range(address));
            else
                workbook.DefinedNames.Add(name, sheet.Range(address));
        }
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateInvalidTemplate(bool otherSheet)
    {
        using var workbook = new XLWorkbook();
        var other = workbook.Worksheets.Add("说明");
        var detail = workbook.Worksheets.Add("明细");
        workbook.DefinedNames.Add("PageTotal_001", otherSheet
            ? other.Range("Z20") : detail.Range("Z20:Z21"));
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    private static void AssertNames(XLWorkbook workbook,
        params (string Name, int Row)[] expected)
    {
        var names = workbook.DefinedNames.ValidNamedRanges()
            .Concat(workbook.Worksheets.SelectMany(sheet => sheet.DefinedNames.ValidNamedRanges()))
            .Where(name => name.Name.StartsWith("PageTotal_", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        Assert.Equal(expected.Length, names.Length);
        foreach (var (nameText, row) in expected)
        {
            var name = Assert.Single(names.Where(item => item.Name == nameText));
            var range = Assert.Single(name.Ranges);
            Assert.Equal("明细", range.Worksheet.Name);
            Assert.Equal(row, range.RangeAddress.FirstAddress.RowNumber);
            Assert.Equal(1, range.RangeAddress.FirstAddress.ColumnNumber);
            Assert.Equal(range.RangeAddress.FirstAddress, range.RangeAddress.LastAddress);
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
