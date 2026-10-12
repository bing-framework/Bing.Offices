using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Tests.Entities;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using ClosedXML.Excel;
using Xunit;
using PageDocument = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.ClosedXml.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证逐页布局保留标记的导出和导入一致性。
/// </summary>
public sealed class EntityIndependentPageMarkerTest
{
    /// <summary>
    /// 验证普通内容不能在扫描列重复使用页尾标记。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="factory">是否通过工厂生成冲突值。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task PageFooter_DuplicateContent_ShouldPreserveStreamsAndObserveOnce(bool asynchronous,
        bool factory)
    {
        var calls = 0;
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2).PageFooterNamed("Sign", "SIGN",
                footer =>
                {
                    if (factory)
                        footer.Cell("A2", _ => { calls++; return " SIGN "; });
                    else
                        footer.Cell("A2", "SIGN");
                }).Footer("END")));
        using var template = CreateTemplate();
        var templateBytes = template.ToArray();
        using var target = new MemoryStream();
        var original = new byte[] { 3, 1, 4, 1, 5 };
        target.Write(original);
        var observer = new EntityPageAnchorFailureObserver();
        var exporter = new ClosedXmlExcelExporter(exceptionObservers: new[] { observer });
        BingOfficesConfigurationException error;
        if (asynchronous)
            error = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                exporter.ExportForTemplateAsync(CreateDocument(1), layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), target));
        else
            error = Assert.Throws<BingOfficesConfigurationException>(() =>
                exporter.ExportForTemplate(CreateDocument(1), layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), target));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(factory ? 1 : 0, calls);
        Assert.Equal(1, observer.Count);
        Assert.Same(error, observer.LastException);
        Assert.Equal(original, target.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
        Assert.True(target.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证首页内容不能占用页尾扫描标记。
    /// </summary>
    [Fact]
    public void PageHeader_FooterMarkerContent_ShouldReportSourceCoordinate()
    {
        var calls = 0;
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A3",
            document => document.Lines, region => region.PageBreak(2)
                .PageHeaderNamed("Head", "A1:C3", "C2", "HEADER",
                    header => header.Cell("A1", _ => { calls++; return "SIGN"; }))
                .PageFooterNamed("Sign", "SIGN").Footer("END")));
        using var template = CreateTemplate();
        using var target = new MemoryStream();
        var original = new byte[] { 2, 7, 1 };
        target.Write(original);
        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(CreateDocument(1), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), target));
        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, error.Code);
        Assert.Equal(BingOfficesStage.Plan, error.Stage);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(1, calls);
        Assert.Equal(original, target.ToArray());
    }

    /// <summary>
    /// 验证非扫描列可使用相同文本且工厂每页只执行一次。
    /// </summary>
    [Fact]
    public void PageFooter_MarkerInOtherColumn_ShouldRoundTripCompleteWorkbook()
    {
        var calls = 0;
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.Cell("A2", "审核")
                    .Cell("B2", _ => { calls++; return "SIGN"; })).Footer("END")));
        var source = CreateDocument(3);
        using var template = CreateTemplate();
        using var target = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), target);
        Assert.Equal(2, calls);
        using (var workbook = new XLWorkbook(new MemoryStream(target.ToArray(), writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal(9, sheet.LastRowUsed().RowNumber());
            var expected = new[] { "Group|Code|Amount", "G|L1|1", "G|L2|2", "SIGN||",
                "审核|SIGN|", "G|L3|3", "SIGN||", "审核|SIGN|", "END||" };
            Assert.Equal(expected, Enumerable.Range(1, 9).Select(row => string.Join("|",
                Enumerable.Range(1, 3).Select(column => sheet.Cell(row, column).GetString()))));
            Assert.Equal(new[] { "Sign_001", "Sign_002" }, workbook.DefinedNames.Select(name => name.Name));
        }
        target.Position = 0;
        var result = new ClosedXmlExcelImporter().ImportEntity(target, layout);
        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        Assert.Equal(source.Lines.Select(line => (line.Group, line.Code, line.Amount)),
            result.Entity.Lines.Select(line => (line.Group, line.Code, line.Amount)));
        Assert.True(target.CanRead);
    }

    /// <summary>
    /// 创建空模板。
    /// </summary>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplate()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("明细");
        var stream = new MemoryStream();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建测试单据。
    /// </summary>
    /// <param name="count">明细数量。</param>
    /// <returns>测试单据。</returns>
    private static PageDocument CreateDocument(int count) => new()
    {
        Lines = Enumerable.Range(1, count).Select(index => new PageLine
            { Group = "G", Code = "L" + index, Amount = index }).ToList()
    };
}
