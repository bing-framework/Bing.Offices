using System;
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
using PageDocument = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageDocument;
using PageLine = Bing.Offices.Tests.EntityRepeatedPageAnchorTest.PageLine;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证逐页布局保留标记的导出和导入一致性。
/// </summary>
public sealed class EntityIndependentPageMarkerTest
{
    /// <summary>
    /// 验证普通内容不能在扫描列重复使用页尾标记。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="factory">是否通过工厂生成冲突值。</param>
    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, true)]
    [InlineData(false, false, true)]
    [InlineData(true, true, true)]
    [InlineData(false, true, true)]
    public async Task PageFooter_DuplicateContent_ShouldPreserveStreamsAndObserveOnce(bool xls,
        bool asynchronous, bool factory)
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
        using var template = CreateTemplate(xls);
        var templateBytes = template.ToArray();
        using var target = new MemoryStream();
        var original = new byte[] { 3, 1, 4, 1, 5 };
        target.Write(original);
        var observer = new IndependentPageAnchorRecordingObserver();
        var exporter = new NpoiExcelExporter(exceptionObservers: new[] { observer });
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
        Assert.Same(error, Assert.Single(observer.Exceptions));
        Assert.Equal(original, target.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
        Assert.True(target.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证首页内容不能占用页尾扫描标记。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageHeader_FooterMarkerContent_ShouldReportSourceCoordinate(bool xls)
    {
        var calls = 0;
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A3",
            document => document.Lines, region => region.PageBreak(2)
                .PageHeaderNamed("Head", "A1:C3", "C2", "HEADER",
                    header => header.Cell("A1", _ => { calls++; return "SIGN"; }))
                .PageFooterNamed("Sign", "SIGN").Footer("END")));
        using var template = CreateTemplate(xls);
        using var target = new MemoryStream();
        var original = new byte[] { 2, 7, 1 };
        target.Write(original);
        var error = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(CreateDocument(1), layout,
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
    /// <param name="xls">是否使用 HSSF。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PageFooter_MarkerInOtherColumn_ShouldRoundTripCompleteWorkbook(bool xls)
    {
        var calls = 0;
        var layout = ExcelEntity.Layout<PageDocument>(builder => builder.ListRegion("明细", "A1",
            document => document.Lines, region => region.PageBreak(2)
                .PageFooterNamed("Sign", "SIGN", footer => footer.Cell("A2", "审核")
                    .Cell("B2", _ => { calls++; return "SIGN"; })).Footer("END")));
        var source = CreateDocument(3);
        using var template = CreateTemplate(xls);
        using var target = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), target);
        Assert.Equal(2, calls);
        using (var workbook = WorkbookFactory.Create(new MemoryStream(target.ToArray(), writable: false)))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(8, sheet.LastRowNum);
            var expected = new[] { "Group|Code|Amount", "G|L1|1", "G|L2|2", "SIGN||",
                "审核|SIGN|", "G|L3|3", "SIGN||", "审核|SIGN|", "END||" };
            Assert.Equal(expected, Enumerable.Range(0, 9).Select(row => string.Join("|",
                Enumerable.Range(0, 3).Select(column => sheet.GetRow(row)?.GetCell(column)?.ToString() ?? ""))));
            Assert.Equal(new[] { "Sign_001", "Sign_002" }, workbook.GetAllNames().Select(name => name.NameName));
        }
        target.Position = 0;
        var result = new NpoiExcelImporter().ImportEntity(target, layout);
        Assert.True(result.IsSuccess);
        Assert.Empty(result.Errors);
        Assert.Equal(source.Lines.Select(line => (line.Group, line.Code, line.Amount)),
            result.Entity.Lines.Select(line => (line.Group, line.Code, line.Amount)));
        Assert.True(target.CanRead);
    }

    /// <summary>
    /// 创建空模板。
    /// </summary>
    /// <param name="xls">是否使用 HSSF。</param>
    /// <returns>调用方持有的模板流。</returns>
    private static MemoryStream CreateTemplate(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("明细");
        var stream = new MemoryStream();
        workbook.Write(stream, true);
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
