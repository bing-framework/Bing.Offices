using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Internals;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML Provider 的直接职责测试；所有 Workbook 断言均通过真实 ClosedXML 读回。
/// </summary>
public sealed class ClosedXmlProviderTest
{
    /// <summary>
    /// 验证 ClosedXML 实体布局按模板命名锚点写入并读回固定单元格和列表区域。
    /// </summary>
    [Fact]
    public void EntityLayout_NamedAnchors_ShouldRoundTripTemplate()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .CellNamed("单据", "InvoiceCustomer", item => item.Customer)
            .ListRegionNamed("明细", "InvoiceLines", item => item.Lines));
        var source = new EntityInvoice
        {
            Customer = "客户 A",
            Lines = new List<EntityLine>
            {
                new EntityLine { Code = "A", Quantity = 2 },
                new EntityLine { Code = "B", Quantity = 3 }
            }
        };

        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var invoice = workbook.Worksheets.Add("单据");
            var lines = workbook.Worksheets.Add("明细");
            workbook.DefinedNames.Add("InvoiceCustomer", invoice.Range("B2:B2"));
            workbook.DefinedNames.Add("InvoiceLines", lines.Range("A5:A5"));
            workbook.SaveAs(template);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            Assert.Equal("客户 A", workbook.Worksheet("单据").Cell("B2").GetString());
            Assert.Equal("Code", workbook.Worksheet("明细").Cell("A5").GetString());
            Assert.Equal("B", workbook.Worksheet("明细").Cell("A7").GetString());
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("客户 A", result.Entity.Customer);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证 ClosedXML 实体布局缺少命名锚点时在模板预检阶段失败。
    /// </summary>
    [Fact]
    public void EntityLayout_NamedAnchor_ShouldRejectMissingTemplateName()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.CellNamed("单据", "InvoiceCustomer", item => item.Customer));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("单据");
            workbook.SaveAs(template);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice(), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Contains("命名锚点", exception.Message);
    }

    /// <summary>
    /// 验证 ClosedXML 实体布局拒绝重复命名锚点和命名列表的绝对结束地址。
    /// </summary>
    [Fact]
    public void EntityLayout_NamedAnchor_ShouldRejectDuplicateAndAbsoluteEnd()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder => builder
            .CellNamed("单据", "InvoiceCustomer", item => item.Customer)
            .CellNamed("单据", "InvoiceCustomer", item => item.Title)));

        var exception = Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.ListRegionNamed("明细", "InvoiceLines", item => item.Lines,
                region => region.End("D20"))));
        Assert.Contains("绝对 End", exception.Message);
    }

    /// <summary>
    /// 验证最终 Footer 命名锚点在新工作簿导出后指向真实 marker，且导入仍按 marker 截断明细。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterNamed_ShouldCreateNameAndRoundTrip()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计")
                .FooterNamed("OrderFooter", "订单合计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount)))));

        foreach (var count in new[] { 0, 3 })
        {
            var source = new GroupSubtotalDocument
            {
                Lines = Enumerable.Range(1, count).Select(index => new GroupSubtotalLine
                {
                    Category = $"C{index}", Amount = index
                }).ToList()
            };
            using var output = new MemoryStream();
            new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
            using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
            {
                var name = Assert.Single(workbook.DefinedNames.ValidNamedRanges()
                    .Where(item => item.Name == "OrderFooter"));
                var range = Assert.Single(name.Ranges);
                Assert.Equal("明细", range.Worksheet.Name);
                Assert.Equal(count == 0 ? 2 : 6, range.RangeAddress.FirstAddress.RowNumber);
                Assert.Equal(1, range.RangeAddress.FirstAddress.ColumnNumber);
                Assert.Equal(range.RangeAddress.FirstAddress.RowNumber, range.RangeAddress.LastAddress.RowNumber);
                Assert.Equal(range.RangeAddress.FirstAddress.ColumnNumber, range.RangeAddress.LastAddress.ColumnNumber);
                Assert.Equal("订单合计", workbook.Worksheet("明细").Cell(count == 0 ? "A2" : "A6").GetString());
                Assert.DoesNotContain(workbook.DefinedNames.ValidNamedRanges(), item => item.Name == "页小计");
            }
            Assert.Null(ReadDefinedNameLocalSheetId(output.ToArray(), "OrderFooter"));

            using var input = new MemoryStream(output.ToArray(), writable: false);
            var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(item => item.Category), result.Entity.Lines.Select(item => item.Category));
            Assert.Equal(source.Lines.Select(item => item.Amount), result.Entity.Lines.Select(item => item.Amount));
        }
    }

    /// <summary>
    /// 验证模板中的 global/local 单格名称会随最终 Footer marker 重定位。
    /// </summary>
    /// <param name="localScope">名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EntityLayout_FooterNamed_ShouldRelocateTemplateName(bool localScope)
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计")
                .FooterNamed("OrderFooter", "订单合计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount)))));

        foreach (var count in new[] { 0, 3 })
        {
            var source = new GroupSubtotalDocument
            {
                Lines = Enumerable.Range(1, count).Select(index => new GroupSubtotalLine
                {
                    Category = $"C{index}", Amount = index
                }).ToList()
            };
            using var template = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("说明");
                var sheet = workbook.Worksheets.Add("明细");
                if (localScope)
                    sheet.DefinedNames.Add("OrderFooter", sheet.Range("Z20"));
                else
                    workbook.DefinedNames.Add("OrderFooter", sheet.Range("Z20"));
                workbook.SaveAs(template);
            }

            template.Position = 0;
            using var output = new MemoryStream();
            new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false)))
            {
                var definedNames = localScope
                    ? workbook.Worksheet("明细").DefinedNames.ValidNamedRanges()
                    : workbook.DefinedNames.ValidNamedRanges();
                var name = Assert.Single(definedNames
                    .Where(item => item.Name == "OrderFooter"));
                var range = Assert.Single(name.Ranges);
                Assert.Equal("明细", range.Worksheet.Name);
                Assert.Equal(count == 0 ? 2 : 6, range.RangeAddress.FirstAddress.RowNumber);
                Assert.Equal(1, range.RangeAddress.FirstAddress.ColumnNumber);
                Assert.Equal(range.RangeAddress.FirstAddress.RowNumber, range.RangeAddress.LastAddress.RowNumber);
                Assert.Equal(range.RangeAddress.FirstAddress.ColumnNumber, range.RangeAddress.LastAddress.ColumnNumber);
            }
            Assert.Equal(localScope ? "1" : null,
                ReadDefinedNameLocalSheetId(output.ToArray(), "OrderFooter"));

            using var input = new MemoryStream(output.ToArray(), writable: false);
            var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(item => item.Category), result.Entity.Lines.Select(item => item.Category));
            Assert.Equal(source.Lines.Select(item => item.Amount), result.Entity.Lines.Select(item => item.Amount));
            Assert.True(template.CanRead);
        }
    }

    /// <summary>
    /// 验证最终尾部名称冲突会在布局阶段拒绝，且普通 Footer 会清除先前名称。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterNamed_ShouldRejectDuplicateAnchorsAndClearOnFooter()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder => builder
            .CellNamed("单据", "OrderFooter", item => item.Title)
            .ListRegion("明细", "A1", item => item.Lines,
                region => region.FooterNamed("OrderFooter", "订单合计"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder =>
            builder.ListRegionNamed("明细", "OrderFooter", document => document.Lines,
                region => region.FooterNamed("OrderFooter", "订单合计"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.FooterNamed("OrderFooter", "订单合计"))
            .ListRegion("副本", "A1", document => document.Lines,
                region => region.FooterNamed("OrderFooter", "副本合计"))));

        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .FooterNamed("DiscardedFooter", "旧 marker")
                .Footer("最终 marker")));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new GroupSubtotalDocument(), layout, output);
        using var workbook = new XLWorkbook(new MemoryStream(output.ToArray(), writable: false));
        Assert.DoesNotContain(workbook.DefinedNames.ValidNamedRanges(), item => item.Name == "DiscardedFooter");
        Assert.Equal("最终 marker", workbook.Worksheet("明细").Cell("A2").GetString());
    }

    /// <summary>
    /// 验证最终 Footer 名称冲突和预取消都不会改写既有目标流。
    /// </summary>
    /// <param name="localScope">冲突名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EntityLayout_FooterNamed_ShouldPreserveOutputOnConflictAndCancellation(bool localScope)
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .FooterNamed("OrderFooter", "订单合计")));
        var source = new GroupSubtotalDocument
        {
            Lines = new List<GroupSubtotalLine>
            {
                new GroupSubtotalLine { Category = "A", Amount = 1 }
            }
        };
        var original = new byte[] { 1, 2, 3, 4 };

        using (var template = new MemoryStream())
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("说明");
            var sheet = workbook.Worksheets.Add("明细");
            if (localScope)
                sheet.DefinedNames.Add("OrderFooter", sheet.Range("Z20:Z21"));
            else
                workbook.DefinedNames.Add("OrderFooter", sheet.Range("Z20:Z21"));
            workbook.SaveAs(template);
            template.Position = 0;
            using var output = new MemoryStream();
            output.Write(original, 0, original.Length);
            output.Position = 0;
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Equal(original, output.ToArray());
            Assert.True(template.CanRead);
        }

        using var canceledOutput = new MemoryStream();
        canceledOutput.Write(original, 0, original.Length);
        canceledOutput.Position = 0;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportEntityAsync(source, layout, canceledOutput, cancellation.Token));
        Assert.Equal(original, canceledOutput.ToArray());
    }

    /// <summary>
    /// 验证 ClosedXML 实体布局计算列使用行上下文并在导入时跳过。
    /// </summary>
    [Fact]
    public void EntityLayout_CalculatedColumns_ShouldUseRowContextAndSkipOnImport()
    {
        var contexts = new List<(int Index, int RowNumber, int ColumnNumber, int ItemCount, string Sheet)>();
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("明细", "B2", item => item.Lines, region => region
                .CalculatedColumn<int>("line-no", "序号", context =>
                {
                    contexts.Add((context.Index, context.RowNumber, context.ColumnNumber,
                        context.Items.Count, context.SheetName));
                    return context.Index + 1;
                }, calculated => calculated.Placement(ExcelColumnPlacement.Before("Code")))
                .CalculatedColumn<decimal>("amount", "金额", context => context.Item.Quantity * 10m,
                    calculated => calculated.Placement(ExcelColumnPlacement.After("Quantity"))
                        .NumberFormat("0.00"))));
        var source = new EntityInvoice
        {
            Lines = new List<EntityLine>
            {
                new EntityLine { Code = "A", Quantity = 2 },
                new EntityLine { Code = "B", Quantity = 3 }
            }
        };

        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);
        Assert.Equal(new[] { (0, 3, 2, 2, "明细"), (1, 4, 2, 2, "明细") }, contexts);
        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("序号", sheet.Cell("B2").GetString());
            Assert.Equal("Code", sheet.Cell("C2").GetString());
            Assert.Equal("金额", sheet.Cell("E2").GetString());
            Assert.Equal(1, sheet.Cell("B3").GetDouble());
            Assert.Equal(20, sheet.Cell("E3").GetDouble());
            Assert.Equal("0.00", sheet.Cell("E3").Style.NumberFormat.Format);
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(item => item.Quantity));
        Assert.Equal(2, contexts.Count);
    }

    /// <summary>
    /// 验证 ClosedXML 计算列委托异常包含完整坐标。
    /// </summary>
    [Fact]
    public void EntityLayout_CalculatedColumnFailure_ShouldReportCoordinates()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .CalculatedColumn<decimal>("amount", "金额", _ => throw new InvalidOperationException("计算失败"))));

        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesExportException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice
            {
                Lines = new List<EntityLine> { new EntityLine { Code = "A", Quantity = 1 } }
            }, layout, output));
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(3, exception.ColumnIndex);
        Assert.Equal("amount", exception.PropertyName);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证 ClosedXML 能力声明、DI 注册和首次注册优先语义。
    /// </summary>
    [Fact]
    public void CapabilitiesAndDi_ShouldBeProviderSpecificAndFirstRegistrationWins()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesClosedXml()
            .BuildServiceProvider();

        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        Assert.IsType<ClosedXmlExcelExporter>(exporter);
        Assert.IsType<ClosedXmlExcelImporter>(importer);
        Assert.Equal("ClosedXML", ((IExcelProviderCapabilities)exporter).ProviderName);
        Assert.True(((IExcelProviderCapabilities)exporter).Supports(
            ExcelProviderCapabilities.List | ExcelProviderCapabilities.Workbook
            | ExcelProviderCapabilities.Entity
            | ExcelProviderCapabilities.Xlsx));
        Assert.IsAssignableFrom<IExcelEntityExporter>(exporter);
        Assert.IsAssignableFrom<IExcelEntityImporter>(importer);
        Assert.IsType<ClosedXmlExcelExporter>(provider.GetRequiredService<IExcelEntityExporter>());
        Assert.IsType<ClosedXmlExcelImporter>(provider.GetRequiredService<IExcelEntityImporter>());

        using var first = new ServiceCollection()
            .AddBingOfficesClosedXml()
            .AddBingOfficesClosedXml()
            .BuildServiceProvider();
        Assert.IsType<ClosedXmlExcelExporter>(first.GetRequiredService<IExcelExporter>());

        using var configured = new ServiceCollection()
            .AddBingOfficesClosedXml(options =>
            {
                options.MaxConcurrentWorkbooks = 2;
                options.MaxQueuedOperations = 3;
            })
            .BuildServiceProvider();
        var configuredOptions = configured.GetRequiredService<ClosedXmlProviderOptions>();
        Assert.Equal(2, configuredOptions.MaxConcurrentWorkbooks);
        Assert.Equal(3, configuredOptions.MaxQueuedOperations);
    }

    /// <summary>
    /// 验证多工作表、样式和公式导出后可被 ClosedXML 读回。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteMultipleSheetsStylesAndFormula()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(new ExcelCellStyle { Bold = true })
                    .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00" }))
            .AddSheet("Formula", new[] { new FormulaRow { Formula = "=1+1" } }));
        using var stream = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, stream);
        Assert.True(stream.Length > 0);
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);
        Assert.Equal(new[] { "People", "Formula" }, workbook.Worksheets.Select(sheet => sheet.Name));
        var people = workbook.Worksheet("People");
        Assert.Equal("Name", people.Cell(1, 1).GetString());
        Assert.Equal("Alice", people.Cell(2, 1).GetString());
        Assert.Equal(30d, people.Cell(2, 2).GetDouble());
        Assert.True(people.Cell(1, 1).Style.Font.Bold);
        Assert.Equal("0.00", people.Cell(2, 3).Style.NumberFormat.Format);
        Assert.Equal("1+1", workbook.Worksheet("Formula").Cell(2, 1).FormulaA1);
    }

    /// <summary>
    /// 验证边框样式覆盖和模板对齐样式保留。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteBorderStyleAndPreserveTemplateAlignment()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell(1, 1).Value = "Template Header";
            sheet.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var style = new ExcelCellStyle
        {
            Bold = true,
            TopBorder = new ExcelBorderStyle
            {
                LineStyle = ExcelBorderLineStyle.Thin,
                Color = new ExcelColor { Argb = "112233" }
            },
            BottomBorder = new ExcelBorderStyle { LineStyle = ExcelBorderLineStyle.Double }
        };
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(style)));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var header = workbook.Worksheet("People").Cell(1, 1);
        Assert.True(header.Style.Font.Bold);
        Assert.Equal(XLBorderStyleValues.Thin, header.Style.Border.TopBorder);
        Assert.Equal(XLBorderStyleValues.Double, header.Style.Border.BottomBorder);
        Assert.Equal(XLAlignmentHorizontalValues.Center, header.Style.Alignment.Horizontal);
    }

    /// <summary>
    /// 验证显式样式重置在模板样式叠加前生效。
    /// </summary>
    [Fact]
    public void Export_ShouldApplyExplicitStyleResetBeforeOverlay()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            var templateHeader = sheet.Cell(1, 1);
            templateHeader.Value = "Old";
            templateHeader.Style.Font.Bold = true;
            templateHeader.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            templateHeader.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderStyle(new ExcelCellStyle
                {
                    Reset = new ExcelCellStyleReset
                    {
                        Bold = true,
                        HorizontalAlignment = true,
                        TopBorder = true
                    }
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var header = workbook.Worksheet("People").Cell(1, 1);
        Assert.False(header.Style.Font.Bold);
        Assert.Equal(XLAlignmentHorizontalValues.General, header.Style.Alignment.Horizontal);
        Assert.Equal(XLBorderStyleValues.None, header.Style.Border.TopBorder);
    }

    /// <summary>
    /// 验证自定义表头跨度和批注可以写入真实工作簿。
    /// </summary>
    [Fact]
    public void Export_ShouldWriteCustomHeaderSpansAndComments()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new Person { Name = "Alice", Age = 30, Score = 1.25m } },
                sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2).HeaderRows(new[]
                {
                    new ExcelHeaderRow(0, new[]
                    {
                        new ExcelHeaderCell(0, "Profile", columnSpan: 2,
                            comment: new ExcelComment("generated", "tester")),
                        new ExcelHeaderCell(2, "Score")
                    })
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Profile", sheet.Cell("A1").GetString());
        Assert.Equal("Score", sheet.Cell("C1").GetString());
        Assert.Equal("A1:B1", Assert.Single(sheet.MergedRanges).RangeAddress.ToString());
        Assert.True(sheet.Cell("A1").HasComment);
        Assert.Equal("generated", sheet.Cell("A1").GetComment().ToString());
        Assert.Equal("tester", sheet.Cell("A1").GetComment().Author);
    }

    /// <summary>
    /// 验证映射、日期和数值字段能够完成导出导入往返。
    /// </summary>
    [Fact]
    public void RoundTrip_ShouldImportMappingAndDateNumberValues()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "Bob", Age = 41, Score = 12.5m } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var person = Assert.Single(result.Workbook.People);
        Assert.Equal("Bob", person.Name);
        Assert.Equal(41, person.Age);
        Assert.Equal(12.5m, person.Score);
        Assert.Equal(new[] { 1 }, result.Sheets.Single().SourceRows);
    }

    /// <summary>
    /// 验证最大行数资源限制在 ClosedXML 加载前报告结构化错误。
    /// </summary>
    [Fact]
    public void Import_MaxRows_ShouldRejectBeforeClosedXmlLoadWithStructuredResourceError()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[]
            {
                new Person { Name = "First", Age = 1, Score = 1m },
                new Person { Name = "Second", Age = 2, Score = 2m }
            }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Single(result.Sheets);
        Assert.Empty(result.Sheets[0].SourceRows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        Assert.Equal("People", error.SheetName);
        Assert.Equal(3, error.RowIndex);
    }

    /// <summary>
    /// 验证取消令牌在 ClosedXML 加载前阻止导入。
    /// </summary>
    [Fact]
    public void Import_Cancellation_ShouldStopBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "cancel" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().Import(source, request, cancellation.Token));
    }

    /// <summary>
    /// 验证异步非定位输入在超过输入上限时边读边拒绝。
    /// </summary>
    [Fact]
    public async Task ImportAsync_NonSeekableInputOverLimit_ShouldRejectDuringAsyncBuffering()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = new string('x', 512) } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        const int maximum = 32;
        using var source = new TrackingNonSeekableReadStream(bytes);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = maximum })
            .Sheet<Person>("People", root => root.People));

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().ImportAsync(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.InRange(source.BytesRead, maximum + 1, maximum + 1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证定位输入恰好达到输入字节上限时可以成功导入。
    /// </summary>
    [Fact]
    public async Task ImportAsync_SeekableInputExactlyAtMaxInputBytes_ShouldSucceed()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "exact-limit" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = generated.Length })
            .Sheet<Person>("People", root => root.People));

        generated.Position = 0;
        var result = await new ClosedXmlExcelImporter().ImportAsync(generated, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("exact-limit", Assert.Single(result.Workbook.People).Name);
        Assert.True(generated.CanRead);
    }

    /// <summary>
    /// 验证图片资源限制在 ClosedXML 加载前生效。
    /// </summary>
    [Fact]
    public void Import_PictureResourceLimit_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "picture" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.CreateEntry("xl/media/image1.png");
            using var image = entry.Open();
            image.Write(new byte[] { 1, 2, 3 }, 0, 3);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxPictureBytes = 2 })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Contains("图片", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证图片数量和总资源限制在 DOM 创建前生效。
    /// </summary>
    [Fact]
    public void Import_PictureCountAndTotalResourceLimits_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "picture" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var pair in new[]
            {
                (Name: "xl/media/image1.png", Bytes: new byte[] { 1, 2, 3 }),
                (Name: "xl/media/image2.png", Bytes: new byte[] { 4, 5, 6 })
            })
            {
                var entry = archive.CreateEntry(pair.Name);
                using var image = entry.Open();
                image.Write(pair.Bytes, 0, pair.Bytes.Length);
            }
        }
        var bytes = source.ToArray();

        using var countSource = new MemoryStream(bytes, writable: false);
        var countRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxPictures = 1 })
            .Sheet<Person>("People", root => root.People));
        var countException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(countSource, countRequest));

        using var totalSource = new MemoryStream(bytes, writable: false);
        var totalRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTotalPictureBytes = 5 })
            .Sheet<Person>("People", root => root.People));
        var totalException = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(totalSource, totalRequest));

        Assert.Equal(BingOfficesStage.Preflight, countException.Stage);
        Assert.Equal(BingOfficesStage.Preflight, totalException.Stage);
        Assert.Contains("图片", countException.Message, StringComparison.Ordinal);
        Assert.Contains("图片", totalException.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 验证公式字符串属性保留导入的公式文本。
    /// </summary>
    [Fact]
    public void FormulaImport_ShouldPreserveFormulaTextForStringProperty()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Formula",
            new[] { new FormulaRow { Formula = "=1+1" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<FormulaWorkbook>(workbook =>
            workbook.Sheet<FormulaRow>("Formula", root => root.Rows));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("=1+1", Assert.Single(result.Workbook.Rows).Formula);
    }

    /// <summary>
    /// 验证隐藏行列、合并锚点、空白和错误单元格的导入边界。
    /// </summary>
    [Fact]
    public void ImportBoundaries_ShouldIncludeHiddenRowsColumnsAndReadMergeAnchorBlankAndError()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Boundary");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = "Empty";
            sheet.Cell("C1").Value = "Error";
            sheet.Cell("D1").Value = "Merged";
            sheet.Cell("A2").Value = "first";
            sheet.Cell("B2").Value = string.Empty;
            sheet.Cell("C2").Value = XLError.DivisionByZero;
            sheet.Cell("D2").Value = "anchor";
            sheet.Range("D2:E2").Merge();
            sheet.Cell("A3").Value = "hidden";
            sheet.Cell("B3").Value = string.Empty;
            sheet.Cell("C3").Value = XLError.NumberInvalid;
            sheet.Cell("D3").Value = "hidden-anchor";
            sheet.Row(3).Hide();
            sheet.Column(2).Hide();
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<BoundaryWorkbook>(workbook => workbook
            .Sheet<BoundaryRow>("Boundary", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "first", "hidden" }, result.Workbook.Rows.Select(row => row.Name));
        Assert.All(result.Workbook.Rows, row => Assert.Null(row.Empty));
        Assert.Equal("DivisionByZero", result.Workbook.Rows[0].Error);
        Assert.Equal("anchor", result.Workbook.Rows[0].Merged);
        Assert.Equal("hidden-anchor", result.Workbook.Rows[1].Merged);
    }

    /// <summary>
    /// 验证错误单元格转换到不兼容目标类型时报告转换错误。
    /// </summary>
    [Fact]
    public void ImportErrorCell_ShouldReportConversionForIncompatibleTargetType()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Errors");
            sheet.Cell("A1").Value = "Amount";
            sheet.Cell("A2").Value = XLError.DivisionByZero;
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook =>
            workbook.Sheet<ErrorNumberRow>("Errors", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ValueConversion
            && error.PropertyName == nameof(ErrorNumberRow.Amount));
        Assert.Empty(result.Workbook.Rows);
    }

    /// <summary>
    /// 验证 1904 日期系统的数值序列解析语义。
    /// </summary>
    [Fact]
    public void Date1904NumericSerial_ShouldUseWorkbookDateSystem()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Dates");
            sheet.Cell(1, 1).Value = "Date";
            sheet.Cell(2, 1).Value = 1d;
            workbook.SaveAs(source);
        }
        MarkWorkbookDate1904(source);
        using (var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true))
        using (var reader = new StreamReader(archive.GetEntry("xl/workbook.xml")!.Open(), Encoding.UTF8, true))
        {
            var workbookXml = reader.ReadToEnd();
            Assert.Contains("date1904=\"1\"", workbookXml);
        }
        Assert.True(ExcelXlsxZipPreflight.GetDate1904(source));
        source.Position = 0;
        var request = ExcelImport.Workbook<DateWorkbook>(workbook =>
            workbook.Sheet<DateRow>("Dates", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new DateTime(1904, 1, 2), Assert.Single(result.Workbook.Rows).Date);
    }

    /// <summary>
    /// 验证实体固定单元格、列表区域和合并区域的往返语义。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRoundTripFixedCellsListAndMerge()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "B2", item => item.Number)
            .Cell("Invoice", "B3", item => item.Customer)
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "B1", item => item.Title)
            .ListRegion("Lines", "A1", item => item.Lines));
        var source = new EntityInvoice
        {
            Number = 1001,
            Customer = "Customer A",
            Title = "Sales Order",
            Lines = new List<EntityLine>
            {
                new EntityLine { Code = "A", Quantity = 2 },
                new EntityLine { Code = "B", Quantity = 3 }
            }
        };
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var invoice = workbook.Worksheet("Invoice");
            Assert.Equal("Sales Order", invoice.Cell("A1").GetString());
            Assert.Equal(1001d, invoice.Cell("B2").GetDouble());
            Assert.Single(invoice.MergedRanges);
            var lines = workbook.Worksheet("Lines");
            Assert.Equal("Code", lines.Cell("A1").GetString());
            Assert.Equal("B", lines.Cell("A3").GetString());
        }

        using var input = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(1001, result.Entity.Number);
        Assert.Equal("Customer A", result.Entity.Customer);
        Assert.Equal("Sales Order", result.Entity.Title);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(line => line.Quantity));
    }

    /// <summary>
    /// 验证实体导入选项在 ClosedXML 加载工作簿前执行输入字节限制。
    /// </summary>
    [Fact]
    public void EntityImportOptions_MaxInputBytes_ShouldRejectBeforeClosedXmlLoad()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(
            new EntityInvoice { Title = "Resource limit" }, layout, generated);
        using var source = new MemoryStream(generated.ToArray());
        IExcelEntityResourceImporter importer = new ClosedXmlExcelImporter();
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxInputBytes = 1 });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntity(source, layout, options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体列表区域按 Order、PlacementKey 和 ColumnIndex 导出并导入动态字典列。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldPreserveHeadersValuesAndRoundTrip()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" },
                new() { PropertyName = nameof(EntityDynamicLine.Quantity), Title = "Qty" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "zone", Title = "Zone", DataTypeName = "string", Order = 0,
                    PlacementKey = $"before:{nameof(EntityDynamicLine.Quantity)}"
                },
                new() { Key = "early", Title = "Early", DataTypeName = "int32", Order = 10 },
                new() { Key = "late", Title = "Late", DataTypeName = "string", Order = 20 },
                new()
                {
                    Key = "indexed", Title = "Indexed", DataTypeName = "boolean", Order = 30,
                    ColumnIndex = 5
                }
            }
        };
        var layout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "B2", root => root.Lines, region => region.Mapping(mapping)));
        var source = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A", Quantity = 2,
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North", ["early"] = 7, ["late"] = "tail", ["indexed"] = true
                    }
                }
            }
        };
        using var exported = new MemoryStream();

        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("Dynamic");
            Assert.Equal(new[] { "Item", "Zone", "Qty", "Early", "Late", "Indexed" },
                sheet.Range("B2:G2").Cells().Select(cell => cell.GetString()));
            Assert.Equal("A", sheet.Cell("B3").GetString());
            Assert.Equal("North", sheet.Cell("C3").GetString());
            Assert.Equal(2d, sheet.Cell("D3").GetDouble());
            Assert.Equal(7d, sheet.Cell("E3").GetDouble());
            Assert.Equal("tail", sheet.Cell("F3").GetString());
            Assert.True(sheet.Cell("G3").GetBoolean());
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("A", line.Name);
        Assert.Equal(2, line.Quantity);
        Assert.Equal("North", line.CustomFields["zone"]);
        Assert.Equal(7, line.CustomFields["early"]);
        Assert.Equal("tail", line.CustomFields["late"]);
        Assert.Equal(true, line.CustomFields["indexed"]);
    }

    /// <summary>
    /// 验证实体动态列复用命名转换器和内置校验绑定。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldApplyConverterAndValidation()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "zone", Title = "Zone", DataTypeName = "string", ConverterName = "entity-title",
                    PlacementKey = $"before:{nameof(EntityDynamicLine.Quantity)}",
                    ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                    {
                        new() { Name = "maxLength", MaxLength = 20 }
                    }
                }
            }
        };
        var layout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines, region => region.Mapping(mapping)));
        var source = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        using var exported = new MemoryStream();
        var converters = new[] { new EntityTitleConverter() };
        new ClosedXmlExcelExporter(converters).ExportEntity(source, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("entity:North", workbook.Worksheet("Dynamic").Cell("B2").GetString());

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter(valueConverters: converters).ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("North", Assert.Single(result.Entity.Lines).CustomFields["zone"]);

        var invalid = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "B",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = new string('x', 21)
                    }
                }
            }
        };
        using var invalidOutput = new MemoryStream();
        Assert.Throws<BingOfficesExportException>(() =>
            new ClosedXmlExcelExporter(converters).ExportEntity(invalid, layout, invalidOutput));
        Assert.Empty(invalidOutput.ToArray());
    }

    /// <summary>
    /// 验证 ClosedXML 实体显式动态列组执行内置校验规则。
    /// </summary>
    [Fact]
    public void EntityLayout_ExplicitDynamicValidation_ShouldReportStructuredError()
    {
        var layout = ExcelEntity.Layout<ExplicitDynamicRoot>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .DynamicColumnGroup("商品", item => item.Fields, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "weight", Title = "重量", DataType = typeof(decimal),
                        ValidationRules = new[]
                        {
                            new ExcelMappingDynamicValidationConfiguration
                            {
                                Name = "range", Min = 1, Max = 10
                            }
                        }
                    }
                })));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new ExplicitDynamicRoot
        {
            Lines = new List<ExplicitDynamicLine>
            {
                new() { Code = "A", Fields = new Dictionary<string, object>
                    { ["weight"] = 5m } }
            }
        }, layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = new XLWorkbook(exported))
        {
            workbook.Worksheet("明细").Cell("B2").Value = 99d;
            using var input = new MemoryStream();
            workbook.SaveAs(input);
            inputBytes = input.ToArray();
        }

        var result = new ClosedXmlExcelImporter().ImportEntity(
            new MemoryStream(inputBytes, writable: false), layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors.Where(item => item.PropertyName == "weight"));
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
    }

    /// <summary>
    /// 验证 ClosedXML 实体显式动态列覆盖全部内置校验规则和转换失败。
    /// </summary>
    [Fact]
    public void EntityLayout_ExplicitDynamicValidation_ShouldCoverAllRules()
    {
        void AssertFailure(ExcelDynamicColumnDefinition definition, IReadOnlyList<object> values,
            Action<IXLCell> mutate, ExcelImportErrorCode code, int rowIndex = 2)
        {
            var layout = ExcelEntity.Layout<ExplicitDynamicRoot>(builder => builder
                .ListRegion("动态验证", "A1", item => item.Lines, region => region
                    .DynamicColumnGroup("商品", item => item.Fields, new[] { definition })));
            using var exported = new MemoryStream();
            new ClosedXmlExcelExporter().ExportEntity(new ExplicitDynamicRoot
            {
                Lines = values.Select((value, index) => new ExplicitDynamicLine
                {
                    Code = $"L{index + 1}",
                    Fields = new Dictionary<string, object> { [definition.Key] = value }
                }).ToList()
            }, layout, exported);
            exported.Position = 0;
            byte[] inputBytes;
            using (var workbook = new XLWorkbook(exported))
            {
                var sheet = workbook.Worksheet("动态验证");
                var columnIndex = sheet.Row(1).CellsUsed()
                    .Single(cell => cell.GetString() == definition.Title).Address.ColumnNumber;
                mutate(sheet.Cell(rowIndex, columnIndex));
                using var input = new MemoryStream();
                workbook.SaveAs(input);
                inputBytes = input.ToArray();
            }

            using var source = new MemoryStream(inputBytes, writable: false);
            var result = new ClosedXmlExcelImporter().ImportEntity(source, layout);
            var error = result.Errors.FirstOrDefault(item => item.Code == code
                && item.ColumnKey == definition.Key && item.RowIndex == rowIndex);
            Assert.False(result.IsSuccess, string.Join(";", result.Errors.Select(item => item.Message)));
            Assert.NotNull(error);
            Assert.Equal("动态验证", error.SheetName);
            Assert.Equal(rowIndex, error.RowIndex);
            Assert.True(error.ColumnIndex > 0);
            Assert.Equal(definition.Key, error.ColumnKey);
        }

        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "required", Title = "必填", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "required" }
            }
        }, new object[] { "有效" }, cell => cell.Value = string.Empty,
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "length", Title = "长度", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "maxLength", MaxLength = 3 }
            }
        }, new object[] { "有效" }, cell => cell.Value = "超长数据",
            ExcelImportErrorCode.MaxLength);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "regex", Title = "正则", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "regex", Pattern = "^OK-" }
            }
        }, new object[] { "OK-001" }, cell => cell.Value = "BAD",
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "date", Title = "日期", DataType = typeof(DateTime),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "date", Format = "yyyy-MM-dd" }
            }
        }, new object[] { new DateTime(2024, 1, 1) }, cell => cell.Value = "2024/01/01",
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "max", Title = "最大值", DataType = typeof(decimal),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "maxValue", MaxValue = 10 }
            }
        }, new object[] { 5m }, cell => cell.Value = 11d, ExcelImportErrorCode.MaxValue);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "unique", Title = "唯一", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "unique", IgnoreEmpty = false }
            }
        }, new object[] { "U1", "U2" }, cell => cell.Value = "U1",
            ExcelImportErrorCode.Validation, rowIndex: 3);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "number", Title = "数值", DataType = typeof(decimal)
        }, new object[] { 5m }, cell => cell.Value = "不是数字",
            ExcelImportErrorCode.ValueConversion);
    }

    /// <summary>
    /// 验证 ClosedXML 实体导入可选择收集当前行的全部校验错误。
    /// </summary>
    [Fact]
    public void EntityLayout_ValidationFailureMode_ShouldControlRowErrors()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateValidationRoot(), layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("验证");
            sheet.Cell("A2").Value = string.Empty;
            sheet.Cell("D2").Value = "BAD";
            using var input = new MemoryStream();
            workbook.SaveAs(input);
            inputBytes = input.ToArray();
        }

        var stopResult = new ClosedXmlExcelImporter().ImportEntity(
            new MemoryStream(inputBytes, writable: false), layout);
        var continueResult = new ClosedXmlExcelImporter().ImportEntity(
            new MemoryStream(inputBytes, writable: false), layout,
            new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue));

        Assert.Single(stopResult.Errors.Where(error => error.RowIndex == 2));
        Assert.True(continueResult.Errors.Count(error => error.RowIndex == 2) >= 2);
        Assert.Contains(continueResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Required));
        Assert.Contains(continueResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Code));
    }

    /// <summary>
    /// 验证 ClosedXML 实体异步和模板异步导入遵守继续收集行校验错误的策略。
    /// </summary>
    [Fact]
    public async Task EntityLayout_ValidationFailureMode_ShouldContinueForAsyncAndTemplate()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateValidationRoot(), layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("验证");
            sheet.Cell("A2").Value = string.Empty;
            sheet.Cell("D2").Value = "BAD";
            using var input = new MemoryStream();
            workbook.SaveAs(input);
            inputBytes = input.ToArray();
        }

        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue);
        var asyncResult = await new ClosedXmlExcelImporter().ImportEntityAsync(
            new MemoryStream(inputBytes, writable: false), layout, options);
        Assert.True(asyncResult.Errors.Count(error => error.RowIndex == 2) >= 2);
        Assert.Contains(asyncResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Required));
        Assert.Contains(asyncResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Code));

        using var template = new MemoryStream(inputBytes, writable: false);
        var templateResult = await new ClosedXmlExcelImporter().ImportForTemplateAsync(
            new MemoryStream(inputBytes, writable: false), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), options);
        Assert.True(templateResult.Errors.Count(error => error.RowIndex == 2) >= 2);
        Assert.Contains(templateResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Required));
        Assert.Contains(templateResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Code));
    }

    /// <summary>
    /// 验证实体动态列与固定列冲突或超出区域边界时在写入前失败。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicColumns_ShouldRejectCollisionAndOverflowBeforeWriting()
    {
        var collisionMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" },
                new() { PropertyName = nameof(EntityDynamicLine.Quantity), Title = "Qty" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string", ColumnIndex = 1 }
            }
        };
        var entity = new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new()
                {
                    Name = "A", Quantity = 1,
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        var collisionLayout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines,
                region => region.Mapping(collisionMapping)));
        using var collisionOutput = new MemoryStream();
        var collision = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(entity, collisionLayout, collisionOutput));
        Assert.Equal(BingOfficesStage.Plan, collision.Stage);
        Assert.Empty(collisionOutput.ToArray());

        var overflowMapping = new ExcelMappingConfiguration
        {
            Columns = collisionMapping.Columns,
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string" }
            }
        };
        var overflowLayout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("Dynamic", "A1", root => root.Lines,
                region => region.End("B2").Mapping(overflowMapping)));
        using var overflowOutput = new MemoryStream();
        var overflow = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(entity, overflowLayout, overflowOutput));
        Assert.Equal(BingOfficesStage.Plan, overflow.Stage);
        Assert.Contains("列表区域超出声明边界", overflow.Message);
        Assert.Empty(overflowOutput.ToArray());
    }

    /// <summary>
    /// 验证属性式固定单元格、多个动态字典组和尾部聚合可在 ClosedXML 中往返。
    /// </summary>
    [Fact]
    public void EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip()
    {
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition { Key = "goods", Title = "商品字段", DataType = typeof(string) },
            new ExcelDynamicColumnDefinition { Key = "product", Title = "产品字段", DataType = typeof(string) }
        };
        var layout = ExcelEntity.LayoutFromAttributes<EntityAttributeRoot>(builder => builder
            .ListRegion("采购单", "A4", root => root.Lines, region => region
                .DynamicColumnGroup("商品组", item => item.Goods, new[] { definitions[0] })
                .DynamicColumnGroup("产品组", item => item.Product, new[] { definitions[1] })
                .Footer("合计", footer => footer
                    .GapRows(1)
                    .MarkerStyle(new ExcelCellStyle { Bold = true })
                    .Cell("B1", items => items.Sum(item => item.Quantity), numberFormat: "0")
                    .Cell("C1", items => items.Sum(item => item.Amount), numberFormat: "0.00")
                    .Cell("D1", "说明", new ExcelCellStyle { Bold = true })
                    .Merge("D1:E1"))));
        var source = new EntityAttributeRoot
        {
            Code = "PO-001",
            Lines = new List<EntityGroupedLine>
            {
                new()
                {
                    Name = "A", Quantity = 2, Amount = 12.5m,
                    Goods = new Dictionary<string, object> { ["goods"] = "G" },
                    Product = new Dictionary<string, object> { ["product"] = "P" }
                },
                new()
                {
                    Name = "B", Quantity = 3, Amount = 7.5m,
                    Goods = new Dictionary<string, object> { ["goods"] = "G2" },
                    Product = new Dictionary<string, object> { ["product"] = "P2" }
                }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("采购单");
            Assert.Equal("PO-001", sheet.Cell("B2").GetString());
            Assert.Equal(new[] { "Name", "Quantity", "Amount", "商品字段", "产品字段" },
                sheet.Range("A4:E4").Cells().Select(cell => cell.GetString()));
            Assert.Equal("合计", sheet.Cell("A8").GetString());
            Assert.True(sheet.Cell("A8").Style.Font.Bold);
            Assert.Equal(5d, sheet.Cell("B8").GetDouble());
            Assert.Equal(20d, sheet.Cell("C8").GetDouble());
            Assert.True(sheet.Cell("D8").Style.Font.Bold);
            Assert.Single(sheet.MergedRanges.Where(range => range.RangeAddress.ToString() == "D8:E8"));
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("PO-001", result.Entity.Code);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Name));
        Assert.Equal("G2", result.Entity.Lines[1].Goods["goods"]);
        Assert.Equal("P", result.Entity.Lines[0].Product["product"]);
    }

    /// <summary>
    /// 验证 ClosedXML 实体列表按连续分组写入小计，并在导入时跳过小计行。
    /// </summary>
    [Fact]
    public void EntityLayout_GroupSubtotal_ShouldRoundTripAndSkipSubtotalRows()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .GroupSubtotal(item => item.Category, "小计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00"))
                .Footer("总计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00"))));
        var source = new GroupSubtotalDocument
        {
            Lines = new List<GroupSubtotalLine>
            {
                new() { Category = "A", Amount = 1.25m },
                new() { Category = "A", Amount = 2.75m },
                new() { Category = "B", Amount = 4m }
            }
        };

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("小计", sheet.Cell("A4").GetString());
            Assert.Equal(4d, sheet.Cell("B4").GetDouble(), 2);
            Assert.Equal("小计", sheet.Cell("A6").GetString());
            Assert.Equal(4d, sheet.Cell("B6").GetDouble(), 2);
            Assert.Equal("总计", sheet.Cell("A7").GetString());
            Assert.Equal(8d, sheet.Cell("B7").GetDouble(), 2);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "A", "B" }, result.Entity.Lines.Select(item => item.Category));
        Assert.Equal(new[] { 1.25m, 2.75m, 4m }, result.Entity.Lines.Select(item => item.Amount));
    }

    /// <summary>
    /// 验证 ClosedXML 实体列表按明细行数写入水平分页符。
    /// </summary>
    [Fact]
    public void EntityLayout_PageBreak_ShouldWriteRowBreaks()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageBreak(2)));
        var source = new EntityInvoice
        {
            Lines = Enumerable.Range(1, 5)
                .Select(index => new EntityLine { Code = $"C{index}", Quantity = index })
                .ToList()
        };

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);

        Assert.Equal(new[] { 3, 5 }, workbook.Worksheet("明细").PageSetup.RowBreaks);
    }

    /// <summary>
    /// 验证签字区推移末页明细后仍在写入目标流前拒绝越界。
    /// </summary>
    /// <param name="physical">是否验证工作表物理边界。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageSignatures_ShouldRejectShiftedDetailOverflow(bool physical)
    {
        var start = physical ? "A1048571" : "A1";
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", start, document => document.Lines, region =>
            {
                region.PageBreak(2).PageSubtotal("页小计", footer => footer
                    .Cell("A2", "签字：").Merge("A2:B3"));
                if (!physical)
                    region.End("B6");
            }));
        var source = new GroupSubtotalDocument
        {
            Lines = Enumerable.Range(1, 3).Select(index =>
                new GroupSubtotalLine { Category = $"C{index}", Amount = index }).ToList()
        };
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        var original = new byte[] { 1, 2, 3, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证分页签字区的完整内容、合并、模板保留和导入边界。
    /// </summary>
    /// <param name="header">是否包含表头。</param>
    /// <param name="asynchronous">是否通过异步入口执行。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PageSignatures_ShouldRoundTripTemplate(bool header, bool asynchronous)
    {
        foreach (var count in new[] { 0, 2, 5 })
        {
            var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
                .ListRegion("明细", "A1", document => document.Lines, region => region
                    .Header(header).PageBreak(2)
                    .PageSubtotal("页小计", footer => footer.GapRows(1)
                        .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00")
                        .Cell("A2", "经办人：", new ExcelCellStyle { Bold = true }).Merge("A2:B3"))
                    .Footer("总计", footer => footer.GapRows(1)
                        .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00")
                        .Cell("A2", "审核人：", new ExcelCellStyle { Bold = true }).Merge("A2:B3"))));
            var source = new GroupSubtotalDocument
            {
                Lines = Enumerable.Range(1, count)
                    .Select(index => new GroupSubtotalLine { Category = $"C{index}", Amount = index }).ToList()
            };
            // 明确列出业务期望，不用生产布局计算器生成断言。
            var expected = count switch
            {
                0 => new[] { "|", "总计|0", "审核人：|", "|" },
                2 => new[] { "C1|1", "C2|2", "|", "总计|3", "审核人：|", "|" },
                _ => new[] { "C1|1", "C2|2", "|", "页小计|3", "经办人：|", "|",
                    "C3|3", "C4|4", "|", "页小计|7", "经办人：|", "|",
                    "C5|5", "|", "总计|15", "审核人：|", "|" }
            };
            var headerRows = header ? 1 : 0;
            var signatureRows = count switch { 0 => new[] { 2 }, 2 => new[] { 4 }, _ => new[] { 4, 10, 15 } };
            using var template = new MemoryStream();
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("明细");
                sheet.Column(1).Width = 24;
                foreach (var row in signatureRows)
                {
                    sheet.Row(row + headerRows + 1).Height = 29;
                    sheet.Range(row + headerRows + 1, 1, row + headerRows + 2, 2).Merge();
                }
                workbook.SaveAs(template);
            }
            template.Position = 0;
            using var output = new MemoryStream();
            var exporter = new ClosedXmlExcelExporter();
            if (asynchronous)
                await exporter.ExportForTemplateAsync(source, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            else
                exporter.ExportForTemplate(source, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            Assert.True(template.CanRead);
            Assert.True(output.CanWrite);
            using (var read = new MemoryStream(output.ToArray(), writable: false))
            using (var workbook = new XLWorkbook(read))
            {
                var sheet = workbook.Worksheet("明细");
                var actual = Enumerable.Range(headerRows + 1, expected.Length).Select(row =>
                    sheet.Cell(row, 1).GetString() + "|" + sheet.Cell(row, 2).GetString()).ToArray();
                Assert.Equal(expected, actual);
                Assert.Equal(signatureRows.Length, sheet.MergedRanges.Count);
                Assert.Equal(24d, sheet.Column(1).Width);
                foreach (var row in signatureRows)
                {
                    Assert.Contains(sheet.MergedRanges, merge => merge.RangeAddress.FirstAddress.RowNumber == row + headerRows + 1
                        && merge.RangeAddress.LastAddress.RowNumber == row + headerRows + 2
                        && merge.RangeAddress.FirstAddress.ColumnNumber == 1 && merge.RangeAddress.LastAddress.ColumnNumber == 2);
                    Assert.Equal(29d, sheet.Row(row + headerRows + 1).Height);
                    Assert.True(sheet.Cell(row + headerRows + 1, 1).Style.Font.Bold);
                }
                Assert.Equal(count == 5 ? new[] { 6 + headerRows, 12 + headerRows } : Array.Empty<int>(), sheet.PageSetup.RowBreaks);
            }
            using var input = new MemoryStream(output.ToArray(), writable: false);
            var importer = new ClosedXmlExcelImporter();
            var result = asynchronous
                ? await importer.ImportEntityAsync(input, layout)
                : importer.ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(item => item.Category), result.Entity.Lines.Select(item => item.Category));
            Assert.Equal(source.Lines.Select(item => item.Amount), result.Entity.Lines.Select(item => item.Amount));
            Assert.True(input.CanRead);
        }
    }

    /// <summary>
    /// 验证 ClosedXML 分页小计按页聚合、写入分页符并在导入时跳过小计行。
    /// </summary>
    [Fact]
    public void EntityLayout_PageSubtotal_ShouldRoundTripAndBreakAfterSubtotal()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00"))
                .Footer("总计", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Amount), numberFormat: "0.00"))));
        var source = new GroupSubtotalDocument
        {
            Lines = new List<GroupSubtotalLine>
            {
                new() { Category = "A", Amount = 1m },
                new() { Category = "B", Amount = 2m },
                new() { Category = "C", Amount = 3m },
                new() { Category = "D", Amount = 4m }
            }
        };

        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("页小计", sheet.Cell("A4").GetString());
            Assert.Equal(3d, sheet.Cell("B4").GetDouble(), 2);
            Assert.Equal("总计", sheet.Cell("A7").GetString());
            Assert.Equal(10d, sheet.Cell("B7").GetDouble(), 2);
            Assert.Equal(new[] { 4 }, sheet.PageSetup.RowBreaks);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { 1m, 2m, 3m, 4m }, result.Entity.Lines.Select(item => item.Amount));
    }

    /// <summary>
    /// 验证分页小计 marker 缺失或位置错误时返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计")
                .Footer("总计")));

        static byte[] CreateWorkbook(params string[] rows)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("明细");
            sheet.Cell("A1").Value = "Category";
            sheet.Cell("B1").Value = "Amount";
            for (var index = 0; index < rows.Length; index++)
            {
                sheet.Cell(index + 2, 1).Value = rows[index];
                sheet.Cell(index + 2, 2).Value = index + 1;
            }
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }

        using (var missing = new MemoryStream(CreateWorkbook("A", "B", "C", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(missing, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("缺少分页小计标记", exception.Message);
        }
        using (var terminal = new MemoryStream(CreateWorkbook("A", "B", "页小计", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(terminal, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("分页小计之后缺少明细", exception.Message);
        }
        using (var invalid = new MemoryStream(CreateWorkbook("A", "页小计", "B", "C", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(invalid, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("分页小计标记位置无效", exception.Message);
        }
    }

    /// <summary>
    /// 验证分页配置的非法行数和分组小计冲突会在布局阶段失败。
    /// </summary>
    [Fact]
    public void EntityLayout_PageBreak_ShouldRejectInvalidCombinations()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageBreak(0))));

        Assert.Throws<InvalidOperationException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .GroupSubtotal(item => item.Category, "小计"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageSubtotal("页小计"))));
    }

    /// <summary>
    /// 验证商品档案支持三个动态字典组、空字典和未知键忽略。
    /// </summary>
    [Fact]
    public void EntityLayout_ProductArchive_ShouldRoundTripThreeDynamicGroups()
    {
        var layout = ExcelEntity.Layout<ProductArchive>(builder => builder
            .ListRegion("商品档案", "A1", item => item.Items, region => region
                .DynamicColumnGroup("商品", item => item.GoodsFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "brand", Title = "品牌", DataType = typeof(string), Order = 0 }
                })
                .DynamicColumnGroup("产品", item => item.ProductFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "weight", Title = "重量", DataType = typeof(string), Order = 1 }
                })
                .DynamicColumnGroup("扩展", item => item.ExtensionFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "enabled", Title = "启用", DataType = typeof(string), Order = 2 }
                })
                .UnknownDynamicValues(ExcelUnknownDynamicValuePolicy.Ignore)));
        var source = new ProductArchive
        {
            Items = new List<ProductArchiveLine>
            {
                new()
                {
                    Code = "SKU-1", Name = "商品一",
                    GoodsFields = new Dictionary<string, object> { ["brand"] = "Bing", ["unknown"] = "忽略" },
                    ProductFields = new Dictionary<string, object> { ["weight"] = "2.5" },
                    ExtensionFields = new Dictionary<string, object> { ["enabled"] = "true" }
                },
                new() { Code = "SKU-2", Name = "商品二" }
            }
        };
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("商品档案");
            Assert.Equal(new[] { "Code", "Name", "品牌", "重量", "启用" },
                sheet.Range("A1:E1").Cells().Select(cell => cell.GetString()));
            Assert.Equal("Bing", sheet.Cell("C2").GetString());
            Assert.Equal("2.5", sheet.Cell("D2").GetString());
            Assert.Equal("true", sheet.Cell("E2").GetString());
        }
        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, result.Entity.Items.Count);
        Assert.Equal("Bing", result.Entity.Items[0].GoodsFields["brand"]);
        Assert.Equal("2.5", result.Entity.Items[0].ProductFields["weight"]);
        Assert.Equal("true", result.Entity.Items[0].ExtensionFields["enabled"]);
        Assert.NotNull(result.Entity.Items[1].GoodsFields);
        Assert.NotNull(result.Entity.Items[1].ProductFields);
        Assert.NotNull(result.Entity.Items[1].ExtensionFields);
    }

    /// <summary>
    /// 验证赠品单无表头且无明细时只输出尾部标记并成功读回。
    /// </summary>
    [Fact]
    public void EntityLayout_GiftOrderWithoutHeaderAndEmptyDetails_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("赠品单", "A1", item => item.Items, region => region
                .Header(false)
                .Footer("合计", footer => footer.Cell("B1", items => items.Count))));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new GiftOrder(), layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("赠品单");
            Assert.Equal("合计", sheet.Cell("A1").GetString());
            Assert.Equal(0, sheet.Cell("B1").GetDouble());
        }
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Empty(result.Entity.Items);
    }

    /// <summary>
    /// 验证 XLSX 模板 Footer 公式写入真实公式单元格，且标记后的公式不会导入为明细。
    /// </summary>
    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void EntityLayout_FooterFormula_ShouldRoundTripAndSkipFooter(int lineCount)
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .Footer("合计", footer => footer
                    .Formula("C1", "=SUM(B1:B2)", numberFormat: "0.00"))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using (var exported = new XLWorkbook(output))
        {
            var sheet = exported.Worksheet("明细");
            var footerRow = lineCount + 1;
            Assert.Equal("合计", sheet.Cell(footerRow, 1).GetString());
            Assert.True(sheet.Cell(footerRow, 3).HasFormula);
            Assert.Equal("SUM(B1:B2)", sheet.Cell(footerRow, 3).FormulaA1);
            Assert.Equal("0.00", sheet.Cell(footerRow, 3).Style.NumberFormat.Format);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(new MemoryStream(template.ToArray(), writable: false)));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
        {
            Assert.Equal("L1", result.Entity.Lines[0].Code);
            Assert.Equal(1.25m, result.Entity.Lines[0].Amount);
        }
    }

    /// <summary>
    /// 验证 Footer 公式必须在布局构建阶段以等号开头并包含公式内容。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterFormula_ShouldRejectInvalidFormulaAtBuild()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Footer("合计", footer => footer.Formula("C1", "SUM(B1:B2)")))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Footer("合计", footer => footer.Formula("C1", "")))));
    }

    /// <summary>
    /// 验证 XLSX 连续明细求和公式会根据空行和公式相对行定位。
    /// </summary>
    /// <param name="lineCount">明细行数。</param>
    /// <param name="setGapBeforeFormula">是否在公式前设置空行数。</param>
    [Theory]
    [InlineData(0, true)]
    [InlineData(0, false)]
    [InlineData(3, true)]
    [InlineData(3, false)]
    public void EntityLayout_FormulaSumContiguousRowsAbove_ShouldRoundTrip(int lineCount, bool setGapBeforeFormula)
    {
        ExcelEntityListFooterBuilder<FooterFormulaLine> footerBuilder = null;
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .Footer("合计", footer =>
                {
                    footerBuilder = footer;
                    if (setGapBeforeFormula)
                        footer.GapRows(1);
                    footer.FormulaSumContiguousRowsAbove("C2", "b", numberFormat: "0.00");
                    if (!setGapBeforeFormula)
                        footer.GapRows(1);
                })));
        footerBuilder.GapRows(4);
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var exported = new XLWorkbook(output))
        {
            var sheet = exported.Worksheet("明细");
            var markerRow = lineCount + 2;
            var formulaRow = markerRow + 1;
            Assert.Equal("合计", sheet.Cell(markerRow, 1).GetString());
            var formula = sheet.Cell(formulaRow, 3);
            Assert.True(formula.HasFormula);
            Assert.Equal(lineCount == 0
                ? "0"
                : $"SUM(INDEX(B:B,ROW()-{lineCount + 2}):INDEX(B:B,ROW()-3))",
                formula.FormulaA1);
            Assert.Equal("0.00", formula.Style.NumberFormat.Format);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(new MemoryStream(template.ToArray(), writable: false)));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
        {
            Assert.Equal("L1", result.Entity.Lines[0].Code);
            Assert.Equal(1.25m, result.Entity.Lines[0].Amount);
        }
    }

    /// <summary>
    /// 验证分页小计可使用连续明细求和公式，导入结果仍只包含明细。
    /// </summary>
    [Fact]
    public void EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForPageSubtotal()
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .PageBreak(2)
                .PageSubtotal("页小计", footer => footer
                    .FormulaSumContiguousRowsAbove("C1", "B", numberFormat: "0.00"))
                .Footer("总计", footer => footer.Cell("C1", items => items.Sum(item => item.Amount)))));
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, 3)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("页小计", sheet.Cell(3, 1).GetString());
            Assert.Equal("SUM(INDEX(B:B,ROW()-2):INDEX(B:B,ROW()-1))", sheet.Cell(3, 3).FormulaA1);
            Assert.Equal("总计", sheet.Cell(5, 1).GetString());
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证分组小计可为每组写入连续明细求和公式并保持导入边界。
    /// </summary>
    [Fact]
    public void EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForGroupSubtotal()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .GroupSubtotal(item => item.Category, "小计", footer => footer
                    .FormulaSumContiguousRowsAbove("C1", "B", numberFormat: "0.00"))));
        var source = new GroupSubtotalDocument
        {
            Lines = new List<GroupSubtotalLine>
            {
                new() { Category = "A", Amount = 1.25m },
                new() { Category = "A", Amount = 2.75m },
                new() { Category = "B", Amount = 4m }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("小计", sheet.Cell(4, 1).GetString());
            Assert.Equal("SUM(INDEX(B:B,ROW()-2):INDEX(B:B,ROW()-1))", sheet.Cell(4, 3).FormulaA1);
            Assert.Equal("小计", sheet.Cell(6, 1).GetString());
            Assert.Equal("SUM(INDEX(B:B,ROW()-1):INDEX(B:B,ROW()-1))", sheet.Cell(6, 3).FormulaA1);
            Assert.Equal("0.00", sheet.Cell(4, 3).Style.NumberFormat.Format);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "A", "B" }, result.Entity.Lines.Select(item => item.Category));
    }

    /// <summary>
    /// 验证非法求和列及最终 Footer 与分页或分组小计冲突时会在布局阶段拒绝。
    /// </summary>
    [Fact]
    public void EntityLayout_FormulaSumContiguousRowsAbove_ShouldRejectInvalidColumnsAndFinalFooterConflicts()
    {
        foreach (var column in new[] { "", "A1", "1", "XFE" })
        {
            Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
                .ListRegion("明细", "A1", document => document.Lines, region => region
                    .Footer("合计", footer => footer.FormulaSumContiguousRowsAbove("C1", column)))));
        }

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计")
                .Footer("总计", footer => footer.FormulaSumContiguousRowsAbove("C1", "B")))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .GroupSubtotal(item => item.Category, "分组小计")
                .Footer("总计", footer => footer.FormulaSumContiguousRowsAbove("C1", "B")))));
    }

    /// <summary>
    /// 验证连续求和公式尾部越过声明边界时保留调用方输出流。
    /// </summary>
    [Fact]
    public void EntityLayout_FormulaSumContiguousRowsAbove_ShouldPreserveOutputOnPlanFailure()
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .End("B10")
                .Footer("合计", footer => footer.FormulaSumContiguousRowsAbove("C2", "B"))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var original = new byte[] { 4, 3, 2, 1 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new FooterFormulaDocument(), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
    }

    /// <summary>
    /// 验证最终 Footer 求和公式会跳过分页小计之间的间隔行及多行内容。
    /// </summary>
    /// <param name="lineCount">明细行数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    public void EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipPageSubtotals(int lineCount)
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .PageBreak(2)
                .PageSubtotal("页小计", footer => footer.GapRows(1)
                    .Cell("B1", items => items.Sum(item => item.Amount))
                    .Cell("D2", "经办人签字"))
                .Footer("总计", footer => footer.GapRows(1)
                    .FormulaSumDetailRowsAbove("C2", "B", numberFormat: "0.00")
                    .Cell("D2", "复核签字"))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            if (lineCount == 0)
            {
                Assert.Equal("总计", sheet.Cell("A2").GetString());
                Assert.True(sheet.Cell("C3").HasFormula);
                Assert.Equal("0", sheet.Cell("C3").FormulaA1);
                Assert.Equal("复核签字", sheet.Cell("D3").GetString());
            }
            else
            {
                Assert.Equal("页小计", sheet.Cell("A4").GetString());
                Assert.Equal("经办人签字", sheet.Cell("D5").GetString());
                Assert.Equal("总计", sheet.Cell("A9").GetString());
                Assert.True(sheet.Cell("C10").HasFormula);
                Assert.Equal("SUM(B1:B2,B6:B7)", sheet.Cell("C10").FormulaA1);
                Assert.Equal("复核签字", sheet.Cell("D10").GetString());
                Assert.Equal("0.00", sheet.Cell("C10").Style.NumberFormat.Format);
            }
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(new MemoryStream(template.ToArray(), writable: false)));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
            Assert.Equal(new[] { "L1", "L2", "L3", "L4" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 验证最终 Footer 求和公式只汇总分组明细，并跳过多行分组小计。
    /// </summary>
    /// <param name="lineCount">明细行数。</param>
    [Theory]
    [InlineData(0)]
    [InlineData(4)]
    public void EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipGroupSubtotals(int lineCount)
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .GroupSubtotal(item => item.Category, "分组小计", footer => footer.GapRows(1)
                    .Cell("B1", items => items.Sum(item => item.Amount))
                    .Cell("D2", "组内签字"))
                .Footer("总计", footer => footer.GapRows(1)
                    .FormulaSumDetailRowsAbove("C2", "B", numberFormat: "0.00")
                    .Cell("D2", "复核签字"))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("明细");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var source = new GroupSubtotalDocument
        {
            Lines = lineCount == 0 ? new List<GroupSubtotalLine>() : new List<GroupSubtotalLine>
            {
                new() { Category = "A", Amount = 1.25m },
                new() { Category = "A", Amount = 2.75m },
                new() { Category = "B", Amount = 3.25m },
                new() { Category = "B", Amount = 4.75m }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("明细");
            if (lineCount == 0)
            {
                Assert.Equal("总计", sheet.Cell("A2").GetString());
                Assert.True(sheet.Cell("C3").HasFormula);
                Assert.Equal("0", sheet.Cell("C3").FormulaA1);
            }
            else
            {
                Assert.Equal("分组小计", sheet.Cell("A4").GetString());
                Assert.Equal("组内签字", sheet.Cell("D5").GetString());
                Assert.Equal("分组小计", sheet.Cell("A9").GetString());
                Assert.Equal("组内签字", sheet.Cell("D10").GetString());
                Assert.Equal("总计", sheet.Cell("A12").GetString());
                Assert.True(sheet.Cell("C13").HasFormula);
                Assert.Equal("SUM(B1:B2,B6:B7)", sheet.Cell("C13").FormulaA1);
                Assert.Equal("0.00", sheet.Cell("C13").Style.NumberFormat.Format);
            }
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(new MemoryStream(template.ToArray(), writable: false)));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
            Assert.Equal(new[] { "A", "A", "B", "B" }, result.Entity.Lines.Select(line => line.Category));
    }

    /// <summary>
    /// 验证明细段求和公式不能配置在分页或分组中间小计。
    /// </summary>
    [Fact]
    public void EntityLayout_FormulaSumDetailRowsAbove_ShouldRejectIntermediateSubtotalAtBuild()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计", footer => footer
                    .FormulaSumDetailRowsAbove("B1", "B")))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .GroupSubtotal(item => item.Category, "分组小计", footer => footer
                    .FormulaSumDetailRowsAbove("B1", "B")))));
    }


    /// <summary>
    /// 验证盘点单支持多工作表、多列表区域和固定汇总单元格。
    /// </summary>
    [Fact]
    public void EntityLayout_Inventory_ShouldRoundTripMultipleSheetsAndRegions()
    {
        var layout = ExcelEntity.Layout<InventoryOrder>(builder => builder
            .Cell("汇总", "B2", item => item.Difference)
            .ListRegion("库存", "A1", item => item.Stocks)
            .ListRegion("调整", "A1", item => item.Adjustments));
        var source = new InventoryOrder
        {
            Difference = -3,
            Stocks = new List<InventoryStock> { new() { Code = "A", Quantity = 10 } },
            Adjustments = new List<InventoryAdjustment> { new() { Code = "A", Quantity = -3 } }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(-3, result.Entity.Difference);
        Assert.Equal("A", Assert.Single(result.Entity.Stocks).Code);
        Assert.Equal(-3, Assert.Single(result.Entity.Adjustments).Quantity);
    }

    /// <summary>
    /// 验证 Footer marker 缺失、重复或早于 GapRows 时返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterMarkerErrors_ShouldBeStructured()
    {
        var layout = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("赠品单", "A1", item => item.Items, region =>
                region.Footer("合计", footer => footer.GapRows(1))));
        static byte[] CreateInput(params string[] markerRows)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("赠品单");
            sheet.Cell("A1").Value = "Code";
            for (var index = 0; index < markerRows.Length; index++)
                sheet.Cell(index + 2, 1).Value = markerRows[index];
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
        using (var missing = new MemoryStream(CreateInput("项目"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(missing, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("缺少尾部标记", exception.Message);
        }
        using (var duplicate = new MemoryStream(CreateInput("合计", "合计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(duplicate, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("尾部标记重复", exception.Message);
        }
        using (var early = new MemoryStream(CreateInput("合计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(early, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("起始位置之前", exception.Message);
        }
    }

    /// <summary>
    /// 验证动态字典为空时自动创建，无法创建时统一返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicDictionaryCreationErrors_ShouldBeStructured()
    {
        static byte[] CreateInput()
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("动态");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("B1").Value = "值";
            sheet.Cell("A2").Value = "A";
            sheet.Cell("B2").Value = "v";
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            return stream.ToArray();
        }
        var definition = new[]
        {
            new ExcelDynamicColumnDefinition { Key = "value", Title = "值", DataType = typeof(string) }
        };
        var writableLayout = ExcelEntity.Layout<WritableDynamicRoot>(builder => builder
            .ListRegion("动态", "A1", item => item.Items, region =>
                region.DynamicColumnGroup("字段", item => item.Values, definition)));
        using (var writable = new MemoryStream(CreateInput(), writable: false))
        {
            var result = new ClosedXmlExcelImporter().ImportEntity(writable, writableLayout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal("v", Assert.Single(result.Entity.Items).Values["value"]);
        }
        var readOnlyLayout = ExcelEntity.Layout<ReadOnlyDynamicRoot>(builder => builder
            .ListRegion("动态", "A1", item => item.Items, region =>
                region.DynamicColumnGroup("字段", item => item.Values, definition)));
        var noConstructorLayout = ExcelEntity.Layout<NoConstructorDynamicRoot>(builder => builder
            .ListRegion("动态", "A1", item => item.Items, region =>
                region.DynamicColumnGroup("字段", item => item.Values, definition)));
        var incompatibleLayout = ExcelEntity.Layout<IncompatibleDynamicRoot>(builder => builder
            .ListRegion("动态", "A1", item => item.Items, region =>
                region.DynamicColumnGroup("字段", item => item.Values, definition)));
        static void AssertCreationFailure<TEntity>(ExcelEntityLayout<TEntity> layout) where TEntity : class, new()
        {
            using var source = new MemoryStream(CreateInput(), writable: false);
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new ClosedXmlExcelImporter().ImportEntity(source, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("动态列属性", exception.Message);
        }
        AssertCreationFailure(readOnlyLayout);
        AssertCreationFailure(noConstructorLayout);
        AssertCreationFailure(incompatibleLayout);
    }

    /// <summary>
    /// 验证 ClosedXML 在写入前拒绝 Footer 物理行列溢出。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterPhysicalBounds_ShouldBePreflighted()
    {
        var rowOverflow = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "A1048576", item => item.Items, region =>
                region.Header(false).Footer("合计", footer => footer.GapRows(1))));
        using var rowOutput = new MemoryStream();
        var rowException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new GiftOrder(), rowOverflow, rowOutput));
        Assert.Equal(BingOfficesStage.Plan, rowException.Stage);
        Assert.Empty(rowOutput.ToArray());

        var columnOverflow = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "XFD1", item => item.Items, region =>
                region.Header(false).Footer("合计", footer => footer.Cell("B1", "越界"))));
        using var columnOutput = new MemoryStream();
        var columnException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new GiftOrder(), columnOverflow, columnOutput));
        Assert.Equal(BingOfficesStage.Plan, columnException.Stage);
        Assert.Empty(columnOutput.ToArray());

        var valid = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "A1048576", item => item.Items, region =>
                region.Header(false).Footer("合计")));
        using var validOutput = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new GiftOrder(), valid, validOutput);
        Assert.NotEmpty(validOutput.ToArray());
    }

    /// <summary>
    /// 验证 ClosedXML 导入只读固定属性时返回结构化配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectReadOnlyPropertyOnImport()
    {
        var layout = ExcelEntity.Layout<ReadOnlyEntity>(builder => builder
            .Cell("单据", "A1", item => item.Title));
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("单据").Cell("A1").Value = "只读";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().ImportEntity(source, layout));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("不可写入", exception.Message);
    }

    /// <summary>
    /// 验证 ClosedXML Footer 超出列表边界时在写入前失败。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectFooterOutsideDeclaredBounds()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("单据", "A1", item => item.Lines, region => region
                .End("C2")
                .Footer("合计", footer => footer.Cell("A2", "越界"))));
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice
            {
                Lines = new List<EntityLine> { new() { Code = "A", Quantity = 1 } }
            }, layout, output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("列表尾部超出声明边界", exception.Message);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证 ClosedXML 模板模式会按运行时明细行创建 Footer 合并区。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldCreateVariableFooterMerge()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("单据", "A1", item => item.Lines, region => region
                .Footer("合计", footer => footer.Merge("B1:C1"))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("单据").Cell("A1").Style.Font.Bold = true;
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice
        {
            Lines = new List<EntityLine> { new() { Code = "A", Quantity = 1 } }
        }, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using var exported = new XLWorkbook(output);
        var sheet = exported.Worksheet("单据");
        Assert.Equal("合计", sheet.Cell("A3").GetString());
        Assert.Single(sheet.MergedRanges.Where(range => range.RangeAddress.ToString() == "B3:C3"));
        Assert.True(sheet.Cell("A1").Style.Font.Bold);
    }

    /// <summary>
    /// 验证模板 Footer 间隔行已有内容时不会被导入为明细。
    /// </summary>
    [Fact]
    public void EntityTemplate_FooterGapWithTemplateContent_ShouldNotImportGapAsDetail()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("单据", "A1", item => item.Lines, region => region
                .Footer("合计", footer => footer.GapRows(1))));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("单据");
            sheet.Cell("A3").Value = "模板间隔内容";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice
        {
            Lines = new List<EntityLine> { new() { Code = "A", Quantity = 1 } }
        }, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using var exported = new XLWorkbook(output);
        Assert.Equal("模板间隔内容", exported.Worksheet("单据").Cell("A3").GetString());
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(new MemoryStream(template.ToArray(), writable: false)));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Single(result.Entity.Lines);
        Assert.Equal("A", result.Entity.Lines[0].Code);
    }

    /// <summary>
    /// 验证实体模板导入前检查合并结构并保留模板流。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldValidateMergeAndPreserveTemplateStream()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "B1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Invoice");
            sheet.Cell("A1").Value = "Template";
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Range("A1:B1").Merge();
            workbook.SaveAs(template);
        }
        var original = template.ToArray();
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice { Title = "Filled" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        Assert.Equal(original, template.ToArray());
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            Assert.Equal("Filled", workbook.Worksheet("Invoice").Cell("A1").GetString());
            Assert.True(workbook.Worksheet("Invoice").Cell("A1").Style.Font.Bold);
            Assert.Single(workbook.Worksheet("Invoice").MergedRanges);
        }

        using var source = new MemoryStream(output.ToArray());
        using var validationTemplate = new MemoryStream(original);
        var result = new ClosedXmlExcelImporter().ImportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(validationTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Filled", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体异步往返和预取消行为。
    /// </summary>
    [Fact]
    public async Task EntityAsync_ShouldRoundTripAndHonorCancellation()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportEntityAsync(
            new EntityInvoice { Title = "Async" }, layout, output);
        Assert.NotEmpty(output.ToArray());

        using var input = new MemoryStream(output.ToArray());
        var result = await new ClosedXmlExcelImporter().ImportEntityAsync(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Async", result.Entity.Title);

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var canceled = new MemoryStream();
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportEntityAsync(
                new EntityInvoice { Title = "Canceled" }, layout, canceled, cancellation.Token));
        Assert.Empty(canceled.ToArray());
    }

    /// <summary>
    /// 验证实体固定单元格使用命名转换器完成双向转换。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseNamedConverter()
    {
        var converter = new EntityTitleConverter();
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title, converter.Name));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter(valueConverters: new[] { converter }).ExportEntity(
            new EntityInvoice { Title = "Sales" }, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("entity:Sales", workbook.Worksheet("Invoice").Cell("A1").GetString());

        using var source = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter(valueConverters: new[] { converter })
            .ImportEntity(source, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Sales", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体固定单元格在双向流程中复用映射值映射。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseMappingValueMapInBothDirections()
    {
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = nameof(EntityInvoice.Title),
            ValueMappings = new List<ExcelValueMappingConfiguration>
            {
                new ExcelValueMappingConfiguration { Text = "Display", Value = "Internal" }
            }
        });
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title, cell => cell.Mapping(mapping)));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice { Title = "Internal" }, layout, exported);

        exported.Position = 0;
        using (var workbook = new XLWorkbook(exported))
            Assert.Equal("Display", workbook.Worksheet("Invoice").Cell("A1").GetString());

        using var source = new MemoryStream(exported.ToArray());
        var result = new ClosedXmlExcelImporter().ImportEntity(source, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Internal", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体固定单元格校验错误包含结构化定位信息。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldReturnStructuredValidationError()
    {
        var layout = ExcelEntity.Layout<RequiredEntity>(builder => builder
            .Cell("Invoice", "C4", item => item.Title));
        using var input = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("C4").Value = string.Empty;
            workbook.SaveAs(input);
        }
        input.Position = 0;

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("Invoice", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RequiredEntity.Title), error.PropertyName);
    }

    /// <summary>
    /// 验证实体列表导入覆盖内置校验矩阵并保留完整坐标。
    /// </summary>
    [Fact]
    public void EntityLayout_ValidationMatrix_ShouldReportCompleteCoordinates()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));

        void AssertFailure(EntityValidationRoot source, string header, int rowIndex,
            Action<IXLCell> mutate, ExcelImportErrorCode code, string propertyName)
        {
            using var exported = new MemoryStream();
            new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);
            exported.Position = 0;
            byte[] inputBytes;
            using (var workbook = new XLWorkbook(exported))
            {
                var sheet = workbook.Worksheet("验证");
                var headerCell = sheet.Row(1).CellsUsed().Single(cell => cell.GetString() == header);
                mutate(sheet.Cell(rowIndex, headerCell.Address.ColumnNumber));
                using var input = new MemoryStream();
                workbook.SaveAs(input);
                inputBytes = input.ToArray();
            }

            using var sourceStream = new MemoryStream(inputBytes, writable: false);
            var result = new ClosedXmlExcelImporter().ImportEntity(sourceStream, layout);
            Assert.False(result.IsSuccess);
            var error = result.Errors.FirstOrDefault(item => item.Code == code
                && item.RowIndex == rowIndex && item.PropertyName == propertyName);
            Assert.NotNull(error);
            Assert.Equal(code, error.Code);
            Assert.Equal("验证", error.SheetName);
            Assert.Equal(rowIndex, error.RowIndex);
            Assert.Equal(propertyName, error.PropertyName);
            Assert.True(error.ColumnIndex > 0);
        }

        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Required), 2,
            cell => cell.Value = string.Empty, ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Required));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Name), 2,
            cell => cell.Value = "超长名称", ExcelImportErrorCode.MaxLength,
            nameof(EntityValidationLine.Name));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Quantity), 2,
            cell => cell.Value = 99d, ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Quantity));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Code), 2,
            cell => cell.Value = "BAD", ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Code));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.DateValue), 2,
            cell => cell.Value = "2024/01/01", ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.DateValue));
        AssertFailure(CreateValidationRoot(new EntityValidationLine { UniqueCode = "U1" },
                new EntityValidationLine { UniqueCode = "U2" }), nameof(EntityValidationLine.UniqueCode), 3,
            cell => cell.Value = "U1", ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.UniqueCode));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Number), 2,
            cell => cell.Value = "不是数字", ExcelImportErrorCode.ValueConversion,
            nameof(EntityValidationLine.Number));
    }

    /// <summary>
    /// 验证实体列表唯一值跟踪达到资源上限时返回结构化错误。
    /// </summary>
    [Fact]
    public void EntityLayout_UniqueResourceLimit_ShouldReportStructuredError()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder => builder
            .ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateValidationRoot(
            new EntityValidationLine { UniqueCode = "U1" },
            new EntityValidationLine { UniqueCode = "U2" }), layout, exported);
        exported.Position = 0;

        var result = new ClosedXmlExcelImporter().ImportEntity(exported, layout,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 }));

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit
            && error.SheetName == "验证"
            && error.PropertyName == nameof(EntityValidationLine.UniqueCode));
    }

    /// <summary>
    /// 验证无效行回滚已预留的唯一值，后续有效行可以复用该值。
    /// </summary>
    [Fact]
    public void EntityLayout_UniqueValue_ShouldRollbackWhenAnotherColumnFails()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder => builder
            .ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(CreateValidationRoot(
            new EntityValidationLine { UniqueCode = "U1", Number = 1 },
            new EntityValidationLine { UniqueCode = "U1", Number = 2 }), layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = new XLWorkbook(exported))
        {
            var sheet = workbook.Worksheet("验证");
            var numberColumn = sheet.Row(1).CellsUsed().Single(cell => cell.GetString() == "Number")
                .Address.ColumnNumber;
            sheet.Cell(2, numberColumn).Value = "不是数字";
            using var input = new MemoryStream();
            workbook.SaveAs(input);
            inputBytes = input.ToArray();
        }

        using var source = new MemoryStream(inputBytes, writable: false);
        var result = new ClosedXmlExcelImporter().ImportEntity(source, layout);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ValueConversion
            && error.RowIndex == 2 && error.PropertyName == nameof(EntityValidationLine.Number));
        Assert.DoesNotContain(result.Errors, error => error.RowIndex == 3
            && error.PropertyName == nameof(EntityValidationLine.UniqueCode));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("U1", line.UniqueCode);
        Assert.Equal(2, line.Number);
    }

    /// <summary>
    /// 验证实体布局的动态列唯一性和大小写比较策略。
    /// </summary>
    [Fact]
    public void EntityLayout_DynamicUniqueColumn_ShouldReportDuplicateValue()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityDynamicLine.Name), Title = "Item" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "code", Title = "Code", DataTypeName = "string",
                    ValidationRules = new List<ExcelMappingDynamicValidationConfiguration>
                    {
                        new() { Name = "unique", IgnoreEmpty = false }
                    }
                }
            }
        };
        var layout = ExcelEntity.Layout<EntityDynamicRoot>(builder => builder
            .ListRegion("动态", "A1", root => root.Lines, region => region.Mapping(mapping)));
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new EntityDynamicRoot
        {
            Lines = new List<EntityDynamicLine>
            {
                new() { Name = "A", CustomFields = new Dictionary<string, object> { ["code"] = "Abc" } },
                new() { Name = "B", CustomFields = new Dictionary<string, object> { ["code"] = "abc" } }
            }
        }, layout, exported);
        exported.Position = 0;

        var result = new ClosedXmlExcelImporter().ImportEntity(exported, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors.Where(item => item.PropertyName == "code"));
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal("动态", error.SheetName);
        Assert.Single(result.Entity.Lines);
        Assert.Equal("Abc", result.Entity.Lines[0].CustomFields["code"]);
    }

    /// <summary>
    /// 验证动态列可以完成 ClosedXML 导出导入往返。
    /// </summary>
    [Fact]
    public void DynamicColumns_ShouldRoundTripThroughClosedXml()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        };
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<DynamicWorkbook>(workbook =>
            workbook.Sheet("Dynamic", root => root.Rows,
                sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal("Alice", row.Name);
        Assert.Equal("east", row.CustomFields["region"]);
    }

    /// <summary>
    /// 验证关系绑定使用配置的 comparer 连接父子数据。
    /// </summary>
    [Fact]
    public void Relations_ShouldUseConfiguredComparerAndBindChildren()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Parents", new[] { new RelationParent { Id = "P-1" } })
            .AddSheet("Children", new[] { new RelationChild { ParentId = "p-1", Name = "Line" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;

        var import = ExcelImport.Workbook<RelationWorkbook>(workbook => workbook
            .Sheet<RelationParent>("Parents", root => root.Parents)
            .Sheet<RelationChild>("Children", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var result = new ClosedXmlExcelImporter().Import(stream, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var parent = Assert.Single(result.Workbook.Parents);
        var child = Assert.Single(parent.Children);
        Assert.Equal("Line", child.Name);
    }

    /// <summary>
    /// 验证唯一约束和唯一资源限制错误均被完整报告。
    /// </summary>
    [Fact]
    public void Import_ShouldReportUniqueAndUniqueResourceLimitErrors()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Codes",
            new[] { new UniqueRow { Code = "A" }, new UniqueRow { Code = "A" } }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;
        var request = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .Sheet<UniqueRow>("Codes", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(stream, request);
        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(3, error.RowIndex);

        using var limitedSource = new MemoryStream();
        new ClosedXmlExcelExporter().Export(ExcelExport.Workbook(workbook => workbook.AddSheet("Codes",
            new[] { new UniqueRow { Code = "A" }, new UniqueRow { Code = "B" } })), limitedSource);
        limitedSource.Position = 0;
        var limited = ExcelImport.Workbook<UniqueWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxTrackedUniqueValues = 1 })
            .Sheet<UniqueRow>("Codes", root => root.Rows));
        var limitedResult = new ClosedXmlExcelImporter().Import(limitedSource, limited);
        Assert.False(limitedResult.IsSuccess);
        Assert.Contains(limitedResult.Errors, item => item.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证达到最大错误数后导入错误集合被截断。
    /// </summary>
    [Fact]
    public void Import_MaxErrors_ShouldTruncateStructuredErrors()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "error-limit" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxErrors = 1 })
            .Sheet<Person>("MissingOne", root => root.People)
            .Sheet<Person>("MissingTwo", root => root.People));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(result.ErrorsTruncated);
        Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, result.Errors[0].Code);
    }

    /// <summary>
    /// 验证属性级合并声明生成真实合并区域。
    /// </summary>
    [Fact]
    public void MergeColumns_ShouldCreateRealMergedRanges()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Merged",
            new[]
            {
                new MergeRow { Group = "A", Name = "One" },
                new MergeRow { Group = "A", Name = "Two" },
                new MergeRow { Group = "B", Name = "Three" }
            }));
        using var stream = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, stream);
        stream.Position = 0;
        using var workbook = new XLWorkbook(stream);

        var ranges = workbook.Worksheet("Merged").MergedRanges.ToArray();
        var range = Assert.Single(ranges);
        Assert.Equal("A2:A3", range.RangeAddress.ToString());
    }

    /// <summary>
    /// 验证 XLS 输入格式在写出前被明确拒绝。
    /// </summary>
    [Fact]
    public void UnsupportedXls_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xls)
            .AddSheet("People", new[] { new Person { Name = "invalid" } }));
        using var destination = new MemoryStream();
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, destination));
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证新增 XLSB 枚举在 ClosedXML 中也不会静默降级为 XLSX。
    /// </summary>
    [Fact]
    public void UnsupportedXlsb_ShouldFailBeforeWriting()
    {
        var request = ExcelExport.Workbook(workbook => workbook.Format(ExcelFormat.Xlsb)
            .AddSheet("People", new[] { new Person { Name = "invalid" } }));
        using var destination = new MemoryStream();
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, destination));
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证资源限制在 ClosedXML 加载前拒绝输入。
    /// </summary>
    [Fact]
    public void ResourceLimit_ShouldRejectBeforeClosedXmlLoad()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "large" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 1 })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证 ZIP 和 XML 预检限制在 DOM 创建前生效。
    /// </summary>
    [Fact]
    public void XlsxPreflight_ShouldRejectConfiguredZipAndXmlResourceLimits()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "resource" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();
        var sharedStringsBytes = ReplaceZipEntry(bytes, "xl/sharedStrings.xml",
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + new string('x', 128) + "</sst>");
        var stylesBytes = ReplaceZipEntry(bytes, "xl/styles.xml",
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
            + new string('x', 128) + "</styleSheet>");
        var scenarios = new[]
        {
            (Name: "entries", Limits: new ExcelResourceLimits { MaxZipEntries = 1 }, Bytes: bytes),
            (Name: "entry-size", Limits: new ExcelResourceLimits { MaxZipEntryUncompressedBytes = 1 }, Bytes: bytes),
            (Name: "total-size", Limits: new ExcelResourceLimits { MaxZipTotalUncompressedBytes = 1 }, Bytes: bytes),
            (Name: "compression-ratio", Limits: new ExcelResourceLimits { MaxZipCompressionRatio = 1 }, Bytes: bytes),
            (Name: "shared-strings", Limits: new ExcelResourceLimits { MaxSharedStringsBytes = 1 }, Bytes: sharedStringsBytes),
            (Name: "styles", Limits: new ExcelResourceLimits { MaxStylesBytes = 1 }, Bytes: stylesBytes),
            (Name: "worksheet", Limits: new ExcelResourceLimits { MaxWorksheetBytes = 1 }, Bytes: bytes),
            (Name: "total-worksheet", Limits: new ExcelResourceLimits { MaxTotalWorksheetBytes = 1 }, Bytes: bytes),
            (Name: "xml-depth", Limits: new ExcelResourceLimits { MaxXmlDepth = 1 }, Bytes: bytes),
            (Name: "xml-characters", Limits: new ExcelResourceLimits { MaxXmlCharacters = 1 }, Bytes: bytes)
        };
        foreach (var scenario in scenarios)
        {
            using var source = new MemoryStream(scenario.Bytes, writable: false);
            var limitedRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
                .ResourceLimits(scenario.Limits)
                .Sheet<Person>("People", root => root.People));
            BingOfficesResourceLimitException exception = null;
            try
            {
                new ClosedXmlExcelImporter().Import(source, limitedRequest);
            }
            catch (BingOfficesResourceLimitException caught)
            {
                exception = caught;
            }

            Assert.NotNull(exception);
            Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
            Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        }
    }

    /// <summary>
    /// 验证 AnnotatedOriginal 失败工作簿保留输入内容并写入错误汇总和批注。
    /// </summary>
    [Fact]
    public void FailureWorkbook_AnnotatedOriginal_ShouldPreserveInputAndWriteDiagnostics()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var data = workbook.Worksheets.Add("Data");
            data.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            data.Cell("A2").Value = "invalid";
            workbook.Worksheets.Add("Preserved").Cell("C3").Value = "keep";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var failureOutput = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureOutput
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        failureOutput.Position = 0;
        using var failureWorkbook = new XLWorkbook(failureOutput);
        Assert.Equal("invalid", failureWorkbook.Worksheet("Data").Cell("A2").GetString());
        Assert.Equal("keep", failureWorkbook.Worksheet("Preserved").Cell("C3").GetString());
        Assert.True(failureWorkbook.Worksheet("Data").Cell("A2").HasComment);
        Assert.Contains(result.Errors.Single().Message,
            failureWorkbook.Worksheet("Data").Cell("A2").GetComment().ToString(),
            StringComparison.OrdinalIgnoreCase);
        Assert.Equal("Bing.Offices", failureWorkbook.Worksheet("Data").Cell("A2").GetComment().Author);
        Assert.Equal("Code", failureWorkbook.Worksheet("_ImportErrors").Cell("A1").GetString());
        Assert.Equal(2, failureWorkbook.Worksheet("_ImportErrors").Cell("D2").GetValue<int>());
    }

    /// <summary>
    /// 验证失败批注冲突策略符合公共契约。
    /// </summary>
    /// <param name="policy">待验证的冲突或不支持功能处理策略。</param>
    /// <param name="expectedText">预期保留的批注文本或错误文本标记。</param>
    /// <param name="expectedAuthor">预期批注作者。</param>
    /// <param name="shouldThrow">是否预期抛出配置异常。</param>
    [Theory]
    [InlineData(ExcelImportCommentConflictPolicy.Preserve, "existing", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Append, "invalid", "source", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Replace, "invalid", "Bing.Offices", false)]
    [InlineData(ExcelImportCommentConflictPolicy.Fail, null, null, true)]
    public void FailureWorkbook_AnnotatedOriginal_ShouldHonorCommentConflictPolicy(
        ExcelImportCommentConflictPolicy policy, string expectedText, string expectedAuthor, bool shouldThrow)
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            var comment = sheet.Cell("A2").CreateComment();
            comment.Author = "source";
            comment.AddText("existing");
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                CommentConflictPolicy = policy
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        ExcelWorkbookImportResult<ErrorNumberWorkbook> result = null;
        Action action = () => result = new ClosedXmlExcelImporter().Import(source, request);

        if (shouldThrow)
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(action);
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Equal(0, destination.Length);
            return;
        }
        action();
        destination.Position = 0;
        using var output = new XLWorkbook(destination);
        var outputComment = output.Worksheet("Data").Cell("A2").GetComment();
        Assert.Contains(expectedText == "invalid" ? result.Errors.Single().Message : expectedText,
            outputComment.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(expectedAuthor, outputComment.Author);
    }

    /// <summary>
    /// 验证 ErrorRowsOnly 仅复制表头和失败行，并附加来源与错误列。
    /// </summary>
    [Fact]
    public void FailureWorkbook_ErrorRowsOnly_ShouldCopyOnlyFailureRows()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = 42;
            sheet.Cell("A3").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = new ClosedXmlExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        destination.Position = 0;
        using var output = new XLWorkbook(destination);
        var outputSheet = output.Worksheet("Data");
        Assert.Equal(nameof(ErrorNumberRow.Amount), outputSheet.Cell("A1").GetString());
        Assert.Equal("invalid", outputSheet.Cell("A2").GetString());
        Assert.Equal("__SourceSheet", outputSheet.Cell("B1").GetString());
        Assert.Equal(3, outputSheet.Cell("C2").GetValue<int>());
        Assert.Equal("Data", outputSheet.Cell("B2").GetString());
        Assert.True(outputSheet.Cell("A3").IsEmpty());
        Assert.True(output.TryGetWorksheet("_ImportErrors", out _));
    }

    /// <summary>
    /// 验证异步导入通过真实异步写入提交失败工作簿，且不关闭调用方流。
    /// </summary>
    [Fact]
    public async Task FailureWorkbook_ImportAsync_ShouldUseAsyncDestinationAndKeepStreamsOpen()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        await using var destination = new AsyncOnlyWriteStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var result = await new ClosedXmlExcelImporter().ImportAsync(source, request);

        Assert.False(result.IsSuccess);
        Assert.True(source.CanRead);
        Assert.True(destination.CanWrite);
        using var output = new XLWorkbook(new MemoryStream(destination.ToArray(), writable: false));
        Assert.True(output.TryGetWorksheet("_ImportErrors", out _));
    }

    /// <summary>
    /// 验证失败工作簿序列化超限不会污染调用方目标流。
    /// </summary>
    [Fact]
    public void FailureWorkbook_MaxSerializedBytes_ShouldRejectBeforeDestinationCopy()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = nameof(ErrorNumberRow.Amount);
            sheet.Cell("A2").Value = "invalid";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ErrorNumberWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination,
                MaxSerializedBytes = 1
            })
            .Sheet<ErrorNumberRow>("Data", root => root.Rows));

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Serialize, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证模板异步导入对不支持部件报告 Import 操作上下文。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateAsync_UnsupportedPart_ShouldReportImportOperation()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.Cell("Invoice", "A1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("A1").Value = "Template";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/chart1.xml");

        using var source = new MemoryStream(template.ToArray(), writable: false);
        template.Position = 0;
        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelImporter().ImportForTemplateAsync(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true)));

        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证异常观察器只接收一次通知且观察器异常不覆盖主异常。
    /// </summary>
    [Fact]
    public void ExceptionObserver_ShouldReceiveClosedXmlFailureOnceAndKeepObserverFailureOutOfBand()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "input" } }));
        using var source = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, source);
        source.Position = 0;
        using (var archive = new ZipArchive(source, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/failure-workbook-chart.xml");
        source.Position = 0;
        using var failureOutput = new MemoryStream();
        var observer = new RecordingObserver(throwOnObserve: true);
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failureOutput
            })
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelImporter(exceptionObservers: new[] { observer }).Import(source, request));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.True(exception.Data.Contains(BingOfficesExceptionDispatcher.ObserverFailureKey));
        Assert.Empty(failureOutput.ToArray());
    }

    /// <summary>
    /// 验证异步资源限制失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveAsyncResourceFailureOnce()
    {
        var export = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "async-resource" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var observer = new RecordingObserver();
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxInputBytes = 32 })
            .Sheet<Person>("People", root => root.People));

        var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter(exceptionObservers: new[] { observer })
                .ImportAsync(new TrackingNonSeekableReadStream(generated.ToArray()), request));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
    }

    /// <summary>
    /// 验证异步不支持能力失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveAsyncUnsupportedFailureOnce()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(ExcelFormat.Xls)
            .AddSheet("People", new[] { new Person { Name = "unsupported" } }));
        var observer = new RecordingObserver();
        using var destination = new MemoryStream();

        var exception = await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter(exceptionObservers: new[] { observer })
                .ExportAsync(request, destination));

        Assert.Equal(1, observer.Count);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 验证文件提交失败只触发一次异常观察通知。
    /// </summary>
    [Fact]
    public async Task ExceptionObserver_ShouldReceiveFileCommitFailureOnce()
    {
        var observer = new RecordingObserver();
        var commitException = new BingOfficesFileCommitException("commit failed", provider: "ClosedXML");
        var exporter = new ClosedXmlExcelExporter(
            exceptionObservers: new[] { observer },
            fileExportCommitter: new ThrowingCommitter(commitException));
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "commit" } }));

        var exception = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
            exporter.ExportToFileAsync(request, Path.Combine(Path.GetTempPath(), "closedxml-observer.xlsx")));

        Assert.Equal(1, observer.Count);
        Assert.Same(commitException, exception);
        Assert.Same(exception, observer.LastException);
        Assert.Equal(BingOfficesOperation.FileCommit, exception.Operation);
    }

    /// <summary>
    /// 验证异步导入导出成功后保持调用方流打开。
    /// </summary>
    [Fact]
    public async Task StreamOwnership_ShouldKeepCallerStreamsOpenForAsyncImportAndExport()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "owned" } }));
        await using var destination = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportAsync(request, destination);
        Assert.True(destination.CanWrite);

        destination.Position = 0;
        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = await new ClosedXmlExcelImporter().ImportAsync(destination, import);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 验证异步非定位输入可以往返并保持源流打开。
    /// </summary>
    [Fact]
    public async Task ImportAsync_NonSeekableInput_ShouldRoundTripAndKeepSourceOpen()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "non-seekable" } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, generated);

        using var source = new TrackingNonSeekableReadStream(generated.ToArray());
        var import = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));
        var result = await new ClosedXmlExcelImporter().ImportAsync(source, import);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("non-seekable", Assert.Single(result.Workbook.People).Name);
        Assert.True(source.CanRead);
        Assert.Equal(generated.Length, source.BytesRead);
    }

    /// <summary>
    /// 验证独立工作簿可以并行读写且不共享 DOM 状态。
    /// </summary>
    [Fact]
    public async Task IndependentWorkbooks_ShouldRoundTripInParallelWithoutDomSharing()
    {
        var names = await Task.WhenAll(Enumerable.Range(0, 4).Select(async index =>
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
                new[] { new Person { Name = $"parallel-{index}" } }));
            await using var output = new MemoryStream();
            await new ClosedXmlExcelExporter().ExportAsync(request, output);
            output.Position = 0;
            using var workbook = new XLWorkbook(output);
            return workbook.Worksheet("People").Cell(2, 1).GetString();
        }));

        Assert.Equal(new[] { "parallel-0", "parallel-1", "parallel-2", "parallel-3" }, names);
    }

    /// <summary>
    /// 验证各原生比较运算符的通过与失败边界。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    [Theory]
    [InlineData(XLAllowedValues.WholeNumber)]
    [InlineData(XLAllowedValues.Decimal)]
    [InlineData(XLAllowedValues.TextLength)]
    [InlineData(XLAllowedValues.Date)]
    [InlineData(XLAllowedValues.Time)]
    public void WorkbookValidationComparisonOperators_ShouldMatchClosedXmlRuleSemantics(
        XLAllowedValues allowedValues)
    {
        foreach (var operation in new[]
                 {
                     XLOperator.Between, XLOperator.NotBetween, XLOperator.EqualTo,
                     XLOperator.NotEqualTo, XLOperator.GreaterThan, XLOperator.LessThan,
                     XLOperator.EqualOrGreaterThan, XLOperator.EqualOrLessThan
                 })
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Data");
            var validation = worksheet.Range("A2").CreateDataValidation();
            ConfigureValidation(validation, allowedValues, operation);
            var (validRaw, validText, invalidRaw, invalidText) = GetComparisonValues(allowedValues, operation);

            var valid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), validRaw,
                validText, CultureInfo.InvariantCulture, isDate1904: allowedValues == XLAllowedValues.Date,
                CancellationToken.None);
            Assert.True(valid.IsValid, $"{allowedValues}/{operation}: {valid.Message}");

            var invalid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), invalidRaw,
                invalidText, CultureInfo.InvariantCulture, isDate1904: allowedValues == XLAllowedValues.Date,
                CancellationToken.None);
            Assert.False(invalid.IsValid, $"{allowedValues}/{operation} must reject its invalid boundary.");
            Assert.False(invalid.IsUnsupported);
        }
    }

    /// <summary>
    /// 验证空值策略和列表区域解析均使用明确的 Workbook 规则边界。
    /// </summary>
    [Fact]
    public void WorkbookValidationBlankAndListRules_ShouldUseExplicitPolicies()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "A";
        lookup.Cell("A2").Value = "B";
        worksheet.Cell("C1").Value = "C";
        worksheet.Cell("C2").Value = "D";

        var blankValidation = worksheet.Range("A2").CreateDataValidation();
        blankValidation.AllowedValues = XLAllowedValues.WholeNumber;
        blankValidation.Operator = XLOperator.EqualTo;
        blankValidation.Value = "1";
        blankValidation.IgnoreBlanks = true;
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), null,
            string.Empty, CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);
        blankValidation.IgnoreBlanks = false;
        var blankRejected = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), null,
            string.Empty, CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.False(blankRejected.IsValid);
        Assert.Equal("不允许 Workbook 校验目标为空。", blankRejected.Message);

        blankValidation.ClearRanges();
        var listValidation = worksheet.Range("A2").CreateDataValidation();
        listValidation.AllowedValues = XLAllowedValues.List;
        listValidation.Value = "Lookup!$A$1:$A$2";
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "B", "B",
            CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);

        listValidation.ClearRanges();
        var localListValidation = worksheet.Range("A2").CreateDataValidation();
        localListValidation.AllowedValues = XLAllowedValues.List;
        localListValidation.Value = "$C$1:$C$2";
        Assert.True(ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "D", "D",
            CultureInfo.InvariantCulture, false, CancellationToken.None).IsValid);
    }

    /// <summary>
    /// 验证无法解析的引用返回不支持结果。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="expression">待验证的原生引用表达式。</param>
    [Theory]
    [InlineData(XLAllowedValues.List, "KnownValues")]
    [InlineData(XLAllowedValues.List, "[External.xlsx]Lookup!A1:A2")]
    [InlineData(XLAllowedValues.List, "MissingSheet!A1:A2")]
    [InlineData(XLAllowedValues.Custom, "MissingName=1")]
    [InlineData(XLAllowedValues.Custom, "MissingSheet!A2=1")]
    public void WorkbookValidationUnresolvableReferences_ShouldBeExplicitlyUnsupported(
        XLAllowedValues allowedValues, string expression)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var validation = worksheet.Range("A2").CreateDataValidation();
        validation.AllowedValues = allowedValues;
        validation.Value = expression;

        var result = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "1", "1",
            CultureInfo.InvariantCulture, false, CancellationToken.None);

        Assert.False(result.IsValid);
        Assert.True(result.IsUnsupported);
        Assert.Equal("Workbook Data Validation 规则类型或公式暂不支持。", result.Message);
    }

    /// <summary>
    /// 验证单单元格比较公式保留为受支持的 Custom 规则，并且不依赖当前值回退。
    /// </summary>
    [Fact]
    public void WorkbookValidationSimpleCustomComparison_ShouldValidateResolvedCellValue()
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Data");
        var validation = worksheet.Range("A2").CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.Custom;
        validation.Value = "$A$2>=5";
        worksheet.Cell("A2").Value = "5";

        var valid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "5", "5",
            CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.True(valid.IsValid, valid.Message);

        worksheet.Cell("A2").Value = "4";
        var invalid = ClosedXmlWorkbookValidationPipeline.Validate(worksheet, worksheet.Cell("A2"), "4", "4",
            CultureInfo.InvariantCulture, false, CancellationToken.None);
        Assert.False(invalid.IsValid);
        Assert.False(invalid.IsUnsupported);
    }

    /// <summary>
    /// 验证模板 leaveOpen 语义及既有样式保留。
    /// </summary>
    [Fact]
    public void TemplateLeaveOpen_ShouldPreserveTemplateStreamAndExistingStyle()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell(1, 1).Value = "Template Header";
            sheet.Cell(1, 1).Style.Font.Italic = true;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Templated" } }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        Assert.True(template.CanRead);
        output.Position = 0;
        using var result = new XLWorkbook(output);
        Assert.Equal("Templated", result.Worksheet("People").Cell(2, 1).GetString());
        Assert.True(result.Worksheet("People").Cell(1, 1).Style.Font.Italic);
    }

    /// <summary>
    /// 验证模板中的公式、批注、验证和条件格式可以保留。
    /// </summary>
    [Fact]
    public void Template_ShouldPreserveFormulaCommentValidationAndConditionalFormatting()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sheet = sourceWorkbook.Worksheets.Add("People");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("D1").FormulaA1 = "=1+1";
            sheet.Cell("D1").CreateComment().AddText("template comment");
            var validation = sheet.Range("D2:D10").CreateDataValidation();
            validation.WholeNumber.Between(1, 10);
            var conditional = sheet.Range("E2:E10").AddConditionalFormat();
            conditional.WhenGreaterThan(5).Fill.SetBackgroundColor(XLColor.Yellow);
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "Templated" } }));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var outputSheet = workbook.Worksheet("People");
        Assert.Equal("1+1", outputSheet.Cell("D1").FormulaA1);
        Assert.True(outputSheet.Cell("D1").HasComment);
        Assert.Equal("template comment", outputSheet.Cell("D1").GetComment().ToString());
        Assert.Single(outputSheet.DataValidations);
        Assert.Single(outputSheet.ConditionalFormats);
    }

    /// <summary>
    /// 验证 leaveOpen=false 时模板流在导出成功后被释放。
    /// </summary>
    [Fact]
    public void TemplateLeaveOpenFalse_ShouldDisposeTemplateAfterExport()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("People").Cell("A1").Value = "Name";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: false)
            .AddSheet("People", new[] { new Person { Name = "disposed" } }));
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, output);

        Assert.False(template.CanRead);
    }

    /// <summary>
    /// 验证预取消的模板操作仍按 leaveOpen=false 释放模板流。
    /// </summary>
    [Fact]
    public async Task TemplateLeaveOpenFalse_ShouldDisposeOnPreCancelledTemplateOperations()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder =>
            builder.Cell("Invoice", "A1", item => item.Title));
        using var exportTemplate = new MemoryStream(new byte[] { 1, 2, 3 });
        using var exportDestination = new MemoryStream();
        using var importTemplate = new MemoryStream(new byte[] { 4, 5, 6 });
        using var source = new MemoryStream(new byte[] { 7, 8, 9 });
        using var workbookTemplate = new MemoryStream(new byte[] { 10, 11, 12 });
        using var workbookDestination = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplateAsync(
                new EntityInvoice { Title = "cancelled" }, layout,
                new ExcelEntityTemplateOptions(exportTemplate, leaveOpen: false), exportDestination,
                cancellation.Token));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().ImportForTemplateAsync(
                source, layout, new ExcelEntityTemplateOptions(importTemplate, leaveOpen: false),
                cancellation.Token));
        var workbookRequest = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(workbookTemplate, leaveOpen: false)
            .AddSheet("People", new[] { new Person { Name = "cancelled" } }));
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelExporter().ExportAsync(workbookRequest, workbookDestination,
                cancellation.Token));

        Assert.False(exportTemplate.CanRead);
        Assert.False(importTemplate.CanRead);
        Assert.False(workbookTemplate.CanRead);
    }

    /// <summary>
    /// 验证包含 Chart 部件的模板在 ClosedXML 加载前失败。
    /// </summary>
    [Fact]
    public void TemplateWithChartPart_ShouldFailBeforeClosedXmlLoad()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            sourceWorkbook.Worksheets.Add("People").Cell(1, 1).Value = "Name";
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry("xl/charts/chart1.xml");
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "rejected" } }));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Empty(output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证不支持的模板部件按名称在输出前 fail-fast。
    /// </summary>
    /// <param name="entryName">要注入模板的 ZIP 部件名称。</param>
    [Theory]
    [InlineData("xl/pivotTables/pivotTable1.xml")]
    [InlineData("xl/pivotCache/pivotCacheDefinition1.xml")]
    [InlineData("xl/media/image1.png")]
    [InlineData("xl/tables/table1.xml")]
    [InlineData("xl/vbaProject.bin")]
    [InlineData("xl/vbaProjectSignature.bin")]
    [InlineData("xl/externalLinks/externalLink1.xml")]
    public void UnsupportedTemplateParts_ShouldFailBeforeClosedXmlLoad(string entryName)
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            sourceWorkbook.Worksheets.Add("People").Cell("A1").Value = "Name";
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;
        using (var archive = new ZipArchive(template, ZipArchiveMode.Update, leaveOpen: true))
            archive.CreateEntry(entryName);
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "rejected" } }));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Empty(output.ToArray());
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 Provider 使用映射计划工厂并隔离导入导出方向。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_ShouldBeUsedByProviderAndKeepDirectionIsolation()
    {
        var inner = ExcelMappingPlanFactoryProvider.CreateDefault();
        var counting = new CountingMappingPlanFactory(inner);
        var document = new ExcelMappingDocument { UseConventionFallback = true };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "cached", Age = 1, Score = 1m } },
            sheet => sheet.Mapping(document)));
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        var exporter = new ClosedXmlExcelExporter(mappingPlanFactory: counting);

        exporter.Export(request, first);
        exporter.Export(request, second);

        Assert.True(counting.CreateCount >= 2);
        var exportPlan = inner.Create<Person>(document, MappingDirection.Export);
        Assert.Same(exportPlan, inner.Create<Person>(document, MappingDirection.Export));
        Assert.NotSame(exportPlan, inner.Create<Person>(document, MappingDirection.Import));
    }

    /// <summary>
    /// 验证映射计划按类型、配置和动态键隔离，并在工厂失败后恢复。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_ShouldIsolateTypeConfigurationAndDynamicKeysAndRecoverAfterFailure()
    {
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault();
        var document = new ExcelMappingDocument { UseConventionFallback = true };
        var basePlan = factory.Create<Person>(document, MappingDirection.Export);
        Assert.Same(basePlan, factory.Create<Person>(document, MappingDirection.Export));
        Assert.NotSame(basePlan, factory.Create<DynamicRow>(document, MappingDirection.Export));

        var configured = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(Person.Name), Title = "Display Name" }
            }
        };
        var configuredPlan = factory.Create<Person>(document, configured, MappingDirection.Export);
        Assert.NotSame(basePlan, configuredPlan);
        Assert.Equal("Display Name", configuredPlan.Columns.Single(column =>
            column.Name == nameof(Person.Name)).Title);

        var dynamicConfiguration = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "region", Title = "Region", DataTypeName = "string" }
            }
        };
        var dynamicPlan = factory.Create<DynamicRow>(document, dynamicConfiguration,
            MappingDirection.Export);
        Assert.Contains(dynamicPlan.DynamicColumns, column => column.Key == "region");
        Assert.NotSame(dynamicPlan, factory.Create<DynamicRow>(document, MappingDirection.Export));

        var invalid = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "", Title = "invalid" }
            }
        };
        Assert.Throws<ArgumentException>(() => factory.Create<Person>(document, invalid,
            MappingDirection.Export));
        var recovered = factory.Create<Person>(document, MappingDirection.Export);
        Assert.Same(basePlan, recovered);
    }

    /// <summary>
    /// 验证模板批注冲突策略产生对应的保留、覆盖或忽略结果。
    /// </summary>
    /// <param name="policy">要验证的批注冲突策略。</param>
    /// <param name="expected">预期读回的批注文本。</param>
    [Theory]
    [InlineData(ExcelCommentConflictPolicy.Preserve, "existing")]
    [InlineData(ExcelCommentConflictPolicy.Replace, "generated")]
    public void TemplateCommentConflict_ShouldHonorPolicy(ExcelCommentConflictPolicy policy, string expected)
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated", Age = 1, Score = 1m } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated", "test"))
                        })
                    })
                    .CommentConflicts(policy)));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var result = new XLWorkbook(output);
        Assert.Equal(expected, result.Worksheet("People").Cell("A1").GetComment().ToString());
    }

    /// <summary>
    /// 验证批注冲突 fail 策略在输出前报告结构化错误。
    /// </summary>
    [Fact]
    public void TemplateCommentConflictFail_ShouldRejectBeforeOutput()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated" } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated"))
                        })
                    })
                    .CommentConflicts(ExcelCommentConflictPolicy.Fail)));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证模板覆盖策略可以清理既有批注和样式。
    /// </summary>
    [Fact]
    public void TemplateOverwrite_ShouldClearTemplateCommentAndStyle()
    {
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").Style.Font.Bold = true;
            sheet.Cell("A1").CreateComment().AddText("existing");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("People", new[] { new Person { Name = "generated" } },
                sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3)
                    .TemplateCellOverwrite(ExcelTemplateCellOverwritePolicy.ReplaceTemplate)
                    .HeaderRows(new[]
                    {
                        new ExcelHeaderRow(0, new[]
                        {
                            new ExcelHeaderCell(0, "Name", comment: new ExcelComment("generated"))
                        })
                    })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var result = new XLWorkbook(output);
        var cell = result.Worksheet("People").Cell("A1");
        Assert.False(cell.Style.Font.Bold);
        Assert.Equal("generated", cell.GetComment().ToString());
    }

    /// <summary>
    /// 验证重叠的实体列表区域在布局构建阶段被拒绝。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectOverlappingListRegions()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("Lines", "A1", item => item.Lines)
            .ListRegion("Lines", "A2", item => item.Lines)));
    }

    /// <summary>
    /// 验证可变行数和尾部导致的列表区域运行时重叠在写入前返回配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectRuntimeOverlapBetweenListRegions()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("Invoice", "A1", item => item.Lines,
                region => region.Footer("合计"))
            .ListRegion("Invoice", "A3", item => item.Lines,
                region => region.End("B4")));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice
            {
                Lines = new List<EntityLine> { new() { Code = "A", Quantity = 1 } }
            }, layout, output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("运行时范围", exception.Message, StringComparison.Ordinal);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证 Footer 聚合委托异常时不应提交部分工作簿。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterValueFactoryFailure_ShouldNotWritePartialWorkbook()
    {
        var expected = new InvalidOperationException("合计失败");
        Func<IReadOnlyList<EntityLine>, string> failure = _ => throw expected;
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .ListRegion("Invoice", "A1", item => item.Lines,
                region => region.Footer("合计", footer => footer.Cell("B1", failure))));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            new ClosedXmlExcelExporter().ExportEntity(new EntityInvoice
            {
                Lines = new List<EntityLine> { new() { Code = "A", Quantity = 1 } }
            }, layout, output));

        Assert.Same(expected, exception.InnerException);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 验证实体模板缺少工作表或合并区域时被拒绝。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldRejectMissingSheetAndMerge()
    {
        var missingSheetLayout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var missingSheetTemplate = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Other");
            workbook.SaveAs(missingSheetTemplate);
        }
        missingSheetTemplate.Position = 0;
        using var missingSheetOutput = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice(), missingSheetLayout,
                new ExcelEntityTemplateOptions(missingSheetTemplate, leaveOpen: true), missingSheetOutput));

        var missingMergeLayout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Merge("Invoice", "A1:B1")
            .Cell("Invoice", "A1", item => item.Title));
        using var missingMergeTemplate = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice");
            workbook.SaveAs(missingMergeTemplate);
        }
        missingMergeTemplate.Position = 0;
        using var missingMergeOutput = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().ExportForTemplate(new EntityInvoice(), missingMergeLayout,
                new ExcelEntityTemplateOptions(missingMergeTemplate, leaveOpen: true), missingMergeOutput));
    }

    /// <summary>
    /// 验证实体模板异步往返并执行模板结构校验。
    /// </summary>
    [Fact]
    public async Task EntityTemplateAsync_ShouldRoundTripAndValidateTemplate()
    {
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Invoice").Cell("A1").Value = "Template";
            workbook.SaveAs(template);
        }
        var templateBytes = template.ToArray();
        template.Position = 0;
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportForTemplateAsync(
            new EntityInvoice { Title = "Async template" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var source = new MemoryStream(output.ToArray());
        using var validationTemplate = new MemoryStream(templateBytes);
        var result = await new ClosedXmlExcelImporter().ImportForTemplateAsync(source, layout,
            new ExcelEntityTemplateOptions(validationTemplate));

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Async template", result.Entity.Title);
    }

    /// <summary>
    /// 验证实体异步文件导出提交后可以重新读取工作簿。
    /// </summary>
    [Fact]
    public async Task EntityExportToFileAsync_ShouldCommitReadableWorkbook()
    {
        var path = Path.Combine(Path.GetTempPath(), $"closedxml-entity-{Guid.NewGuid():N}.xlsx");
        try
        {
            var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
                .Cell("Invoice", "A1", item => item.Title));
            await new ClosedXmlExcelExporter().ExportEntityToFileAsync(
                new EntityInvoice { Title = "file entity" }, layout, path);

            using var workbook = new XLWorkbook(path);
            Assert.Equal("file entity", workbook.Worksheet("Invoice").Cell("A1").GetString());
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证实体文件导出预取消时保留既有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task EntityFileExport_PreCanceled_ShouldPreserveTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-offices-closedxml-entity-test-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(path, original);
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new ClosedXmlExcelExporter().ExportEntityToFileAsync(
                    new EntityInvoice { Title = "取消" }, layout, path, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证实体文件导出写入中取消时保留既有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task EntityFileExport_MidFlightCancellation_ShouldPreserveTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-offices-closedxml-entity-test-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7, 6 };
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstEntityConverter(cancellation);
        var layout = ExcelEntity.Layout<EntityInvoice>(builder => builder
            .Cell("Invoice", "A1", item => item.Title, converter.Name));

        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new ClosedXmlExcelExporter(valueConverters: new[] { converter })
                    .ExportEntityToFileAsync(new EntityInvoice { Title = "中途取消" }, layout,
                        path, cancellation.Token));
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "entity.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证非法 XLSX 输入的预检异常带有 ClosedXML Provider 上下文。
    /// </summary>
    [Fact]
    public void InvalidXlsx_ShouldUseClosedXmlProviderInPreflightError()
    {
        using var source = new MemoryStream(new byte[] { 1, 2, 3, 4 });
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal("ClosedXML", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("ClosedXML", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 表示基础导入测试的工作簿模型。
    /// </summary>
    private sealed class PeopleWorkbook
    {
        /// <summary>
        /// 获取人员数据行集合。
        /// </summary>
        public List<Person> People { get; } = new();
    }

    /// <summary>
    /// 表示边界导入测试的工作簿模型。
    /// </summary>
    private sealed class BoundaryWorkbook
    {
        /// <summary>
        /// 获取边界数据行集合。
        /// </summary>
        public List<BoundaryRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示单元格边界导入测试的数据行。
    /// </summary>
    private sealed class BoundaryRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置显式空字符串字段。
        /// </summary>
        public string Empty { get; set; }

        /// <summary>
        /// 获取或设置错误单元格文本字段。
        /// </summary>
        public string Error { get; set; }

        /// <summary>
        /// 获取或设置合并锚点字段。
        /// </summary>
        public string Merged { get; set; }
    }

    /// <summary>
    /// 表示错误数值单元格测试的工作簿模型。
    /// </summary>
    private sealed class ErrorNumberWorkbook
    {
        /// <summary>
        /// 获取错误数值数据行集合。
        /// </summary>
        public List<ErrorNumberRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示错误数值单元格测试的数据行。
    /// </summary>
    private sealed class ErrorNumberRow
    {
        /// <summary>
        /// 获取或设置金额字段。
        /// </summary>
        public int Amount { get; set; }
    }

    /// <summary>
    /// 验证动态列布局、行高和样式优先级在真实 XLSX 中生效。
    /// </summary>
    [Fact]
    public void DynamicLayoutAndRowHeight_ShouldUsePhysicalColumnAndStylePrecedence()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string),
            Placement = ExcelColumnPlacement.Before(nameof(DynamicRow.Name)),
            HeaderStyle = new ExcelCellStyle { Bold = true },
            BodyStyle = new ExcelCellStyle { NumberFormat = "@" }
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet
                .DynamicColumns(row => row.CustomFields, new[] { definition })
                .HeaderStyle(new ExcelCellStyle { Italic = true })
                .BodyStyle(new ExcelCellStyle { NumberFormat = "0.00" })
                .RowHeight(new ExcelRowHeightOptions { HeaderHeight = 22, BodyHeight = 18 })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("Dynamic");
        Assert.Equal("Region", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal("east", sheet.Cell("A2").GetString());
        Assert.True(sheet.Cell("A1").Style.Font.Bold);
        Assert.True(sheet.Cell("A1").Style.Font.Italic);
        Assert.Equal("@", sheet.Cell("A2").Style.NumberFormat.Format);
        Assert.Equal(22, sheet.Row(1).Height);
        Assert.Equal(18, sheet.Row(2).Height);
    }

    /// <summary>
    /// 验证动态列物理索引保留最终列位置并拒绝冲突。
    /// </summary>
    [Fact]
    public void DynamicPhysicalIndex_ShouldUseFinalColumnAndRejectConflicts()
    {
        var definition = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            PhysicalColumnIndex = 3
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice", CustomFields = new Dictionary<string, object>
            {
                ["region"] = "east"
            } } }, sheet => sheet.DynamicColumns(row => row.CustomFields, new[] { definition })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            var sheet = workbook.Worksheet("Dynamic");
            Assert.Equal("Name", sheet.Cell("A1").GetString());
            Assert.True(sheet.Cell("B1").IsEmpty());
            Assert.True(sheet.Cell("C1").IsEmpty());
            Assert.Equal("Region", sheet.Cell("D1").GetString());
            Assert.Equal("Alice", sheet.Cell("A2").GetString());
            Assert.True(sheet.Cell("B2").IsEmpty());
            Assert.True(sheet.Cell("C2").IsEmpty());
            Assert.Equal("east", sheet.Cell("D2").GetString());
        }

        var conflicting = ExcelExport.Workbook(workbook => workbook.AddSheet("Dynamic",
            new[] { new DynamicRow { Name = "Alice" } }, sheet => sheet.DynamicColumns(row => row.CustomFields,
                new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "first", Title = "First", PhysicalColumnIndex = 2
                    },
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "second", Title = "Second", PhysicalColumnIndex = 2
                    }
                })));
        using var conflictingOutput = new MemoryStream();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ClosedXmlExcelExporter().Export(conflicting, conflictingOutput));
        Assert.Empty(conflictingOutput.ToArray());
    }

    /// <summary>
    /// 验证映射中的固定列索引保持声明的物理列顺序。
    /// </summary>
    [Fact]
    public void MappingColumnIndex_ShouldPreserveConfiguredFixedColumnOrder()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "Alice", Age = 30 } }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(Person.Age), ColumnIndex = 0 },
                        new() { PropertyName = nameof(Person.Name), ColumnIndex = 1 }
                    }
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Age", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal(30, sheet.Cell("A2").GetDouble());
        Assert.Equal("Alice", sheet.Cell("B2").GetString());
    }

    /// <summary>
    /// 验证固定列、动态列和稀疏列宽使用统一的物理布局。
    /// </summary>
    [Fact]
    public void MixedFixedColumnIndex_ShouldFillRemainingSlotsAndSkipSparseWidthGaps()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new SparseFixedRow { Name = "Alice", Age = 30 } }, sheet => sheet
                .Mapping(new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(SparseFixedRow.Name), ColumnIndex = 1 }
                    }
                })
                .DynamicColumns(row => CreateRegionValues(row), new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "region", Title = "Region", PhysicalColumnIndex = 3
                    }
                })
                .ColumnWidth(new ExcelColumnWidthOptions
                {
                    Mode = ExcelColumnWidthMode.Fixed,
                    FixedWidth = 20
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("People");
        Assert.Equal("Age", sheet.Cell("A1").GetString());
        Assert.Equal("Name", sheet.Cell("B1").GetString());
        Assert.Equal(string.Empty, sheet.Cell("C1").GetString());
        Assert.Equal("Region", sheet.Cell("D1").GetString());
        Assert.Equal(20, sheet.Column("A").Width);
        Assert.Equal(20, sheet.Column("B").Width);
        Assert.NotEqual(20, sheet.Column("C").Width);
        Assert.Equal(20, sheet.Column("D").Width);
    }

    /// <summary>
    /// 创建稀疏固定列测试使用的动态列值。
    /// </summary>
    /// <param name="_">当前数据行；该测试不读取行内容。</param>
    /// <returns>动态列键和值。</returns>
    private static IDictionary<string, object> CreateRegionValues(SparseFixedRow _)
    {
        return new Dictionary<string, object>
        {
            ["region"] = "east"
        };
    }

    /// <summary>
    /// 验证规范化后的重复表头在行物化前被拒绝。
    /// </summary>
    [Fact]
    public void Import_DuplicateNormalizedHeaders_ShouldFailBeforeMaterializingRows()
    {
        using var source = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("People");
            sheet.Cell("A1").Value = "Name";
            sheet.Cell("B1").Value = " name ";
            sheet.Cell("A2").Value = "first";
            workbook.SaveAs(source);
        }
        source.Position = 0;
        var request = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelImporter().Import(source, request));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
    }

    /// <summary>
    /// 验证工作表、列数和物理单元格资源限制的精确边界。
    /// </summary>
    [Fact]
    public void ResourceLimits_ShouldEnforceSheetColumnAndPhysicalCellCounts()
    {
        var export = ExcelExport.Workbook(workbook => workbook
            .AddSheet("One", new[] { new Person { Name = "A", Age = 1 } })
            .AddSheet("Two", new[] { new Person { Name = "B", Age = 2 } }));
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().Export(export, generated);
        var bytes = generated.ToArray();

        using var sheetSource = new MemoryStream(bytes, writable: false);
        var sheetRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxSheets = 1 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(sheetSource, sheetRequest));

        using var columnSource = new MemoryStream(bytes, writable: false);
        var columnRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 3 })
            .Sheet<Person>("One", root => root.People));
        var columnResult = new ClosedXmlExcelImporter().Import(columnSource, columnRequest);
        Assert.True(columnResult.IsSuccess, string.Join(";", columnResult.Errors.Select(error => error.Message)));

        using var overColumnSource = new MemoryStream(bytes, writable: false);
        var overColumnRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxColumnsPerSheet = 2 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(overColumnSource, overColumnRequest));

        var physicalCells = CountPhysicalCells(bytes);
        using var exactCellSource = new MemoryStream(bytes, writable: false);
        var exactCellRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxCells = physicalCells })
            .Sheet<Person>("One", root => root.People));
        var exactCellResult = new ClosedXmlExcelImporter().Import(exactCellSource, exactCellRequest);
        Assert.True(exactCellResult.IsSuccess,
            string.Join(";", exactCellResult.Errors.Select(error => error.Message)));

        using var overCellSource = new MemoryStream(bytes, writable: false);
        var overCellRequest = ExcelImport.Workbook<PeopleWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxCells = physicalCells - 1 })
            .Sheet<Person>("One", root => root.People));
        Assert.Throws<BingOfficesResourceLimitException>(() =>
            new ClosedXmlExcelImporter().Import(overCellSource, overCellRequest));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxSheets = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxCells = -1 }.Validate());
    }

    /// <summary>
    /// 验证同步工作簿准入在队列已满时于 DOM 创建前拒绝请求。
    /// </summary>
    [Fact]
    public void WorkbookAdmission_ShouldRejectWhenQueueCapacityIsExceeded()
    {
        var admission = new ClosedXmlWorkbookAdmission(new ClosedXmlProviderOptions
        {
            MaxConcurrentWorkbooks = 1,
            MaxQueuedOperations = 0
        });
        using var held = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            admission.Acquire(CancellationToken.None, BingOfficesOperation.Import));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
    }

    /// <summary>
    /// 验证异步准入的等待、取消、释放和并发统计行为。
    /// </summary>
    [Fact]
    public async Task WorkbookAdmissionAsync_ShouldWaitCancelAndRelease()
    {
        var admission = new ClosedXmlWorkbookAdmission(new ClosedXmlProviderOptions
        {
            MaxConcurrentWorkbooks = 1,
            MaxQueuedOperations = 1
        });
        using var held = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);
        var waiting = admission.AcquireAsync(CancellationToken.None, BingOfficesOperation.Import).AsTask();
        await Task.Yield();
        Assert.False(waiting.IsCompleted);

        held.Dispose();
        await using (await waiting)
        {
        }

        using var heldAgain = admission.Acquire(CancellationToken.None, BingOfficesOperation.Export);
        using var cancellation = new CancellationTokenSource();
        var canceled = admission.AcquireAsync(cancellation.Token, BingOfficesOperation.Import).AsTask();
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() => canceled);
        heldAgain.Dispose();

        using var afterCancellation = admission.Acquire(CancellationToken.None, BingOfficesOperation.Import);
        Assert.Equal(1, admission.RequestedConcurrency);
        Assert.Equal(1, admission.MaxQueuedOperations);
        Assert.Equal(1, admission.MaxActualParallelism);
        Assert.True(admission.MaxObservedQueuedOperations >= 1);
    }

    /// <summary>
    /// 验证导出在目标流阻塞时仍能释放 DOM 准入供后续请求使用。
    /// </summary>
    [Fact]
    public async Task ExportAsync_ShouldReleaseAdmissionBeforeBlockedDestinationCopy()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesClosedXml(options =>
            {
                options.MaxConcurrentWorkbooks = 1;
                options.MaxQueuedOperations = 1;
            })
            .BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
            new[] { new Person { Name = "blocked" } }));
        await using var blockedDestination = new BlockingWriteStream();

        var first = exporter.ExportAsync(request, blockedDestination);
        await blockedDestination.WriteStarted.Task.WaitAsync(TimeSpan.FromSeconds(5));

        await using var secondDestination = new MemoryStream();
        var second = exporter.ExportAsync(request, secondDestination);
        await second.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(secondDestination.Length > 0);

        blockedDestination.Release();
        await first;
    }

    /// <summary>
    /// 验证多个实体列表区域导入后按关系配置绑定子集合。
    /// </summary>
    [Fact]
    public void EntityRelations_ShouldBindMultipleListRegionsAfterImport()
    {
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var entity = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new EntityRelationParent { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new EntityRelationChild { ParentId = "p-1", Name = "child" }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, output);
        using var input = new MemoryStream(output.ToArray());

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var parent = Assert.Single(result.Entity.Parents);
        Assert.Equal("child", Assert.Single(parent.Children).Name);
    }

    /// <summary>
    /// 验证实体关系绑定不会丢失子项动态列值。
    /// </summary>
    [Fact]
    public void EntityRelationsWithDynamicColumns_ShouldRoundTripDynamicValues()
    {
        var childMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(EntityRelationChild.ParentId), Title = "ParentId" },
                new() { PropertyName = nameof(EntityRelationChild.Name), Title = "Name" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string" }
            }
        };
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children,
                region => region.Mapping(childMapping))
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var entity = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new() { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new()
                {
                    ParentId = "P-1", Name = "child",
                    CustomFields = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["zone"] = "North"
                    }
                }
            }
        };
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var child = Assert.Single(Assert.Single(result.Entity.Parents).Children);
        Assert.Equal("child", child.Name);
        Assert.Equal("North", child.CustomFields["zone"]);
    }

    /// <summary>
    /// 验证关系缺失父项报告结构化错误并遵守异步取消。
    /// </summary>
    [Fact]
    public async Task EntityRelations_ShouldReportMissingParentAndHonorAsyncCancellation()
    {
        var layout = ExcelEntity.Layout<EntityRelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var source = new EntityRelationRoot
        {
            Parents = new List<EntityRelationParent> { new() { Id = "P-1" } },
            Children = new List<EntityRelationChild>
            {
                new() { ParentId = "P-2", Name = "orphan" }
            }
        };
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        var result = new ClosedXmlExcelImporter().ImportEntity(exported, layout);

        Assert.False(result.IsSuccess);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.Relationship);
        Assert.Empty(Assert.Single(result.Entity.Parents).Children);

        using var canceledSource = new MemoryStream(exported.ToArray());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAsync<OperationCanceledException>(() =>
            new ClosedXmlExcelImporter().ImportEntityAsync(canceledSource, layout, cancellation.Token));
    }

    /// <summary>
    /// 验证提交失败时保留目标目录并清理原子提交临时文件。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldKeepExistingDirectoryAndCleanTemporaryFileOnCommitFailure()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"closedxml-commit-{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("People",
                new[] { new Person { Name = "commit-failure" } }));
            var exception = Assert.Throws<BingOfficesFileCommitException>(() =>
                new ClosedXmlExcelExporter().ExportToFile(request, directory));

            Assert.Equal(BingOfficesStage.Commit, exception.Stage);
            Assert.True(Directory.Exists(directory));
            Assert.Empty(Directory.GetFiles(Path.GetDirectoryName(directory),
                Path.GetFileName(directory) + ".*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 表示基础 ClosedXML Provider 测试中的人员数据行。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }

        /// <summary>
        /// 获取或设置分数。
        /// </summary>
        public decimal Score { get; set; }
    }

    /// <summary>
    /// 表示稀疏固定列布局测试的数据行。
    /// </summary>
    private sealed class SparseFixedRow
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 表示公式导出测试的数据行。
    /// </summary>
    private sealed class FormulaRow
    {
        /// <summary>
        /// 获取或设置公式文本。
        /// </summary>
        public string Formula { get; set; }
    }

    /// <summary>
    /// 表示公式导出测试的工作簿模型。
    /// </summary>
    private sealed class FormulaWorkbook
    {
        /// <summary>
        /// 获取公式数据行集合。
        /// </summary>
        public List<FormulaRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示日期系统测试的工作簿模型。
    /// </summary>
    private sealed class DateWorkbook
    {
        /// <summary>
        /// 获取日期数据行集合。
        /// </summary>
        public List<DateRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示日期系统测试的数据行。
    /// </summary>
    private sealed class DateRow
    {
        /// <summary>
        /// 获取或设置日期值。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 表示动态列测试的工作簿模型。
    /// </summary>
    private sealed class DynamicWorkbook
    {
        /// <summary>
        /// 获取动态列数据行集合。
        /// </summary>
        public List<DynamicRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示关系测试的工作簿模型。
    /// </summary>
    private sealed class RelationWorkbook
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示关系测试中的父项。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关系绑定后的子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示关系测试中的子项。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示唯一性测试的工作簿模型。
    /// </summary>
    private sealed class UniqueWorkbook
    {
        /// <summary>
        /// 获取唯一性测试的数据行集合。
        /// </summary>
        public List<UniqueRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示唯一性测试的数据行。
    /// </summary>
    private sealed class UniqueRow
    {
        /// <summary>
        /// 获取或设置唯一编码。
        /// </summary>
        [ExcelUnique]
        public string Code { get; set; }
    }

    /// <summary>
    /// 表示动态列测试的数据行。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置姓名。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置动态列值集合。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; } =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 表示合并列测试的数据行。
    /// </summary>
    private sealed class MergeRow
    {
        /// <summary>
        /// 获取或设置分组名称。
        /// </summary>
        [MergeColumns]
        public string Group { get; set; }

        /// <summary>
        /// 获取或设置行名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 表示实体导出测试的根模型。
    /// </summary>
    private sealed class EntityInvoice
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取或设置客户名称。
        /// </summary>
        public string Customer { get; set; }

        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<EntityLine> Lines { get; set; } = new List<EntityLine>();
    }

    /// <summary>
    /// 表示包含只读固定属性的导入测试模型。
    /// </summary>
    private sealed class ReadOnlyEntity
    {
        /// <summary>
        /// 获取单据标题。
        /// </summary>
        public string Title => "只读";
    }

    /// <summary>
    /// 表示动态列表区域测试的根模型。
    /// </summary>
    private sealed class EntityDynamicRoot
    {
        /// <summary>
        /// 获取或设置动态明细行集合。
        /// </summary>
        public List<EntityDynamicLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示显式动态列校验测试的根模型。
    /// </summary>
    private sealed class ExplicitDynamicRoot
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<ExplicitDynamicLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示显式动态列校验测试的明细模型。
    /// </summary>
    private sealed class ExplicitDynamicLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段字典。
        /// </summary>
        public IDictionary<string, object> Fields { get; set; }
    }

    /// <summary>
    /// 表示包含固定列和动态字典列的实体明细行。
    /// </summary>
    private sealed class EntityDynamicLine
    {
        /// <summary>
        /// 获取或设置项目名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置动态列值。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; }
    }

    /// <summary>
    /// 表示属性式单据布局测试的根模型。
    /// </summary>
    private sealed class EntityAttributeRoot
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        [ExcelEntityCell("采购单", "B2")]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置采购明细集合。
        /// </summary>
        public List<EntityGroupedLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示包含多个动态字段字典的采购明细。
    /// </summary>
    private sealed class EntityGroupedLine
    {
        /// <summary>
        /// 获取或设置商品名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置商品字段字典。
        /// </summary>
        public IDictionary<string, object> Goods { get; set; }

        /// <summary>
        /// 获取或设置产品字段字典。
        /// </summary>
        public IDictionary<string, object> Product { get; set; }
    }

    /// <summary>
    /// 连续分组小计测试根模型。
    /// </summary>
    private sealed class GroupSubtotalDocument
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<GroupSubtotalLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 连续分组小计测试明细模型。
    /// </summary>
    private sealed class GroupSubtotalLine
    {
        /// <summary>
        /// 获取或设置分组键。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置明细金额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 商品档案测试的根模型。
    /// </summary>
    private sealed class ProductArchive
    {
        /// <summary>
        /// 获取或设置商品条目。
        /// </summary>
        public List<ProductArchiveLine> Items { get; set; } = new();
    }

    /// <summary>
    /// 商品档案测试的明细模型。
    /// </summary>
    private sealed class ProductArchiveLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置商品名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置商品字段。
        /// </summary>
        public IDictionary<string, object> GoodsFields { get; set; }

        /// <summary>
        /// 获取或设置产品字段。
        /// </summary>
        public IDictionary<string, object> ProductFields { get; set; }

        /// <summary>
        /// 获取或设置扩展字段。
        /// </summary>
        public IDictionary<string, object> ExtensionFields { get; set; }
    }

    /// <summary>
    /// 赠品单测试的根模型。
    /// </summary>
    private sealed class GiftOrder
    {
        /// <summary>
        /// 获取或设置赠品明细。
        /// </summary>
        public List<GiftLine> Items { get; set; } = new();
    }

    /// <summary>
    /// 赠品单测试的明细模型。
    /// </summary>
    private sealed class GiftLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置赠送数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 盘点单测试的根模型。
    /// </summary>
    private sealed class InventoryOrder
    {
        /// <summary>
        /// 获取或设置差异数量。
        /// </summary>
        public int Difference { get; set; }

        /// <summary>
        /// 获取或设置库存明细。
        /// </summary>
        public List<InventoryStock> Stocks { get; set; } = new();

        /// <summary>
        /// 获取或设置调整明细。
        /// </summary>
        public List<InventoryAdjustment> Adjustments { get; set; } = new();
    }

    /// <summary>
    /// 盘点库存明细测试模型。
    /// </summary>
    private sealed class InventoryStock
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置库存数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 盘点调整明细测试模型。
    /// </summary>
    private sealed class InventoryAdjustment
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置调整数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 可自动创建动态字典的测试根模型。
    /// </summary>
    private sealed class WritableDynamicRoot
    {
        /// <summary>
        /// 获取或设置动态明细。
        /// </summary>
        public List<WritableDynamicLine> Items { get; set; } = new();
    }

    /// <summary>
    /// 可自动创建动态字典的测试明细。
    /// </summary>
    private sealed class WritableDynamicLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段。
        /// </summary>
        public IDictionary<string, object> Values { get; set; }
    }

    /// <summary>
    /// 动态字典属性只读的测试根模型。
    /// </summary>
    private sealed class ReadOnlyDynamicRoot
    {
        /// <summary>
        /// 获取动态明细。
        /// </summary>
        public List<ReadOnlyDynamicLine> Items { get; } = new();
    }

    /// <summary>
    /// 动态字典属性只读的测试明细。
    /// </summary>
    private sealed class ReadOnlyDynamicLine
    {
        /// <summary>
        /// 获取明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取动态字段。
        /// </summary>
        public IDictionary<string, object> Values => null;
    }

    /// <summary>
    /// 动态字典具体类型无法默认构造的测试根模型。
    /// </summary>
    private sealed class NoConstructorDynamicRoot
    {
        /// <summary>
        /// 获取动态明细。
        /// </summary>
        public List<NoConstructorDynamicLine> Items { get; } = new();
    }

    /// <summary>
    /// 动态字典具体类型无法默认构造的测试明细。
    /// </summary>
    private sealed class NoConstructorDynamicLine
    {
        /// <summary>
        /// 获取明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段。
        /// </summary>
        public NoConstructorDictionary Values { get; set; }
    }

    /// <summary>
    /// 动态字典接口无法由默认字典实现的测试根模型。
    /// </summary>
    private sealed class IncompatibleDynamicRoot
    {
        /// <summary>
        /// 获取动态明细。
        /// </summary>
        public List<IncompatibleDynamicLine> Items { get; } = new();
    }

    /// <summary>
    /// 动态字典接口无法由默认字典实现的测试明细。
    /// </summary>
    private sealed class IncompatibleDynamicLine
    {
        /// <summary>
        /// 获取明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段。
        /// </summary>
        public IBadDynamicDictionary Values { get; set; }
    }

    /// <summary>
    /// 没有无参构造函数的动态字典类型。
    /// </summary>
    private sealed class NoConstructorDictionary : Dictionary<string, object>
    {
        /// <summary>
        /// 初始化一个 <see cref="NoConstructorDictionary" /> 类型的实例。
        /// </summary>
        /// <param name="capacity">初始容量。</param>
        public NoConstructorDictionary(int capacity) : base(capacity) { }
    }

    /// <summary>
    /// 默认字典未实现的动态字典接口。
    /// </summary>
    private interface IBadDynamicDictionary : IDictionary<string, object>
    {
    }

    /// <summary>
    /// 表示实体关系测试的根模型。
    /// </summary>
    private sealed class EntityRelationRoot
    {
        /// <summary>
        /// 获取或设置父项列表区域。
        /// </summary>
        public List<EntityRelationParent> Parents { get; set; } = new();

        /// <summary>
        /// 获取或设置子项列表区域。
        /// </summary>
        public List<EntityRelationChild> Children { get; set; } = new();
    }

    /// <summary>
    /// 表示实体关系测试中的父项。
    /// </summary>
    private sealed class EntityRelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关系绑定后的子项集合。
        /// </summary>
        public List<EntityRelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 表示实体关系测试中的子项。
    /// </summary>
    private sealed class EntityRelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置子项动态列字典。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> CustomFields { get; set; } =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// 表示实体列表区域中的明细行。
    /// </summary>
    private sealed class EntityLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 表示显式 Footer 公式测试单据。
    /// </summary>
    private sealed class FooterFormulaDocument
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<FooterFormulaLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 表示显式 Footer 公式测试明细。
    /// </summary>
    private sealed class FooterFormulaLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 表示必填实体字段测试模型。
    /// </summary>
    private sealed class RequiredEntity
    {
        /// <summary>
        /// 获取或设置必填标题。
        /// </summary>
        [ExcelRequired]
        public string Title { get; set; }
    }

    /// <summary>
    /// 创建实体导入校验矩阵使用的有效根实体。
    /// </summary>
    private static EntityValidationRoot CreateValidationRoot(params EntityValidationLine[] lines) =>
        new EntityValidationRoot
        {
            Lines = lines.Length == 0
                ? new List<EntityValidationLine> { new() }
                : lines.ToList()
        };

    /// <summary>
    /// 实体列表导入校验矩阵的根模型。
    /// </summary>
    private sealed class EntityValidationRoot
    {
        /// <summary>
        /// 获取或设置校验明细集合。
        /// </summary>
        public List<EntityValidationLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 实体列表导入校验矩阵的明细模型。
    /// </summary>
    private sealed class EntityValidationLine
    {
        /// <summary>
        /// 获取或设置必填字段。
        /// </summary>
        [ExcelRequired]
        public string Required { get; set; } = "必填";

        /// <summary>
        /// 获取或设置长度受限字段。
        /// </summary>
        [ExcelMaxLength(3)]
        public string Name { get; set; } = "名称";

        /// <summary>
        /// 获取或设置范围受限数量。
        /// </summary>
        [ExcelRange(1, 10)]
        public int Quantity { get; set; } = 5;

        /// <summary>
        /// 获取或设置正则校验字段。
        /// </summary>
        [ExcelRegex("^OK-")]
        public string Code { get; set; } = "OK-001";

        /// <summary>
        /// 获取或设置日期字段。
        /// </summary>
        [ExcelDate("yyyy-MM-dd")]
        public DateTime DateValue { get; set; } = new DateTime(2024, 1, 1);

        /// <summary>
        /// 获取或设置唯一字段。
        /// </summary>
        [ExcelUnique]
        public string UniqueCode { get; set; } = "U1";

        /// <summary>
        /// 获取或设置数值转换字段。
        /// </summary>
        public int Number { get; set; } = 1;
    }

    /// <summary>
    /// 提供实体标题的双向命名转换。
    /// </summary>
    private sealed class EntityTitleConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "entity-title";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            var text = context.Value?.ToString() ?? string.Empty;
            value = text.StartsWith("entity:", StringComparison.Ordinal)
                ? text.Substring("entity:".Length) : text;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = "entity:" + (context.Value?.ToString() ?? string.Empty);
            return true;
        }
    }

    /// <summary>
    /// 在第一次实体转换后触发取消的测试转换器。
    /// </summary>
    private sealed class CancelAfterFirstEntityConverter : INamedExcelValueConverter
    {
        /// <summary>
        /// 保存要触发取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 记录已执行的转换次数。
        /// </summary>
        private int _calls;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstEntityConverter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">用于发出取消信号的令牌源。</param>
        public CancelAfterFirstEntityConverter(CancellationTokenSource cancellation)
        {
            _cancellation = cancellation;
        }

        /// <inheritdoc />
        public string Name => "entity-cancel";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value?.ToString();
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value?.ToString();
            if (Interlocked.Increment(ref _calls) == 1)
                _cancellation.Cancel();
            return true;
        }
    }

    /// <summary>
    /// 记录异常通知次数并可选择模拟观察器失败。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 指示观察器是否在收到通知后抛出异常。
        /// </summary>
        private readonly bool _throwOnObserve;

        /// <summary>
        /// 初始化一个 <see cref="RecordingObserver"/> 类型的实例。
        /// </summary>
        /// <param name="throwOnObserve">是否在通知后故意抛出异常。</param>
        public RecordingObserver(bool throwOnObserve = false) => _throwOnObserve = throwOnObserve;

        /// <summary>
        /// 获取已收到的通知次数。
        /// </summary>
        public int Count { get; private set; }

        /// <summary>
        /// 获取最近一次收到的异常。
        /// </summary>
        public BingOfficesException LastException { get; private set; }

        /// <inheritdoc />
        public void Observe(BingOfficesException exception)
        {
            Count++;
            LastException = exception;
            if (_throwOnObserve)
                throw new InvalidOperationException("observer failed");
        }
    }

    /// <summary>
    /// 模拟文件提交失败的提交器。
    /// </summary>
    private sealed class ThrowingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 保存预设的提交异常。
        /// </summary>
        private readonly BingOfficesFileCommitException _exception;

        /// <summary>
        /// 初始化一个 <see cref="ThrowingCommitter"/> 类型的实例。
        /// </summary>
        /// <param name="exception">提交时要抛出的异常。</param>
        public ThrowingCommitter(BingOfficesFileCommitException exception) => _exception = exception;

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
            => throw _exception;

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format) => Task.FromException(_exception);
    }

    /// <summary>
    /// 包装映射计划工厂并记录创建次数。
    /// </summary>
    private sealed class CountingMappingPlanFactory : IExcelMappingPlanFactory
    {
        /// <summary>
        /// 保存实际执行计划创建的内部工厂。
        /// </summary>
        private readonly IExcelMappingPlanFactory _inner;

        /// <summary>
        /// 初始化一个 <see cref="CountingMappingPlanFactory"/> 类型的实例。
        /// </summary>
        /// <param name="inner">实际映射计划工厂。</param>
        public CountingMappingPlanFactory(IExcelMappingPlanFactory inner) =>
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));

        /// <summary>
        /// 获取已调用的计划创建次数。
        /// </summary>
        public int CreateCount { get; private set; }

        /// <inheritdoc />
        public IExcelMappingPlan Create<T>(ExcelMappingDocument document, MappingDirection direction)
            where T : class, new()
        {
            CreateCount++;
            return _inner.Create<T>(document, direction);
        }

        /// <inheritdoc />
        public IExcelMappingPlan Create<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
            where T : class, new()
        {
            CreateCount++;
            return _inner.Create<T>(document, requestConfiguration, direction);
        }

        /// <inheritdoc />
        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new() =>
            _inner.CreateWorkbook<T>(document, direction, sheetNames);

        /// <inheritdoc />
        public IExcelMappingWorkbookPlan CreateWorkbook<T>(ExcelMappingDocument document,
            ExcelMappingConfiguration requestConfiguration, MappingDirection direction,
            IReadOnlyList<string> sheetNames) where T : class, new() =>
            _inner.CreateWorkbook<T>(document, requestConfiguration, direction, sheetNames);
    }

    /// <summary>
    /// 提供不可定位读取流并记录实际读取字节数。
    /// </summary>
    private sealed class TrackingNonSeekableReadStream : Stream
    {
        /// <summary>
        /// 保存底层只读内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="TrackingNonSeekableReadStream"/> 类型的实例。
        /// </summary>
        /// <param name="bytes">要读取的输入字节。</param>
        public TrackingNonSeekableReadStream(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

        /// <summary>
        /// 获取已读取的字节数。
        /// </summary>
        public long BytesRead { get; private set; }

        /// <inheritdoc />
        public override bool CanRead => true;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            var read = _inner.Read(buffer.Span);
            BytesRead += read;
            return ValueTask.FromResult(read);
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => Task.FromResult(Read(buffer, offset, count));

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Flush() { }
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    /// <summary>
    /// 只允许异步写入的内存流，用于拒绝异步 API 内的同步写入回退。
    /// </summary>
    private sealed class AsyncOnlyWriteStream : Stream
    {
        /// <summary>
        /// 保存异步写入内容的内存流，由测试包装流负责释放。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 获取已写入内容的字节副本。
        /// </summary>
        /// <returns>当前写入内容的独立字节数组。</returns>
        internal byte[] ToArray() => _inner.ToArray();

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => throw new NotSupportedException();

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);

        /// <inheritdoc />
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
            => _inner.WriteAsync(buffer, cancellationToken);

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
            => _inner.WriteAsync(buffer, offset, count, cancellationToken);

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
            => throw new NotSupportedException("不允许同步写入。");

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value)
            => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 提供可控阻塞的写入流以测试 DOM 准入释放时机。
    /// </summary>
    private sealed class BlockingWriteStream : Stream
    {
        /// <summary>
        /// 保存底层内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <summary>
        /// 获取首次写入已开始的通知。
        /// </summary>
        public TaskCompletionSource<object> WriteStarted { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 获取允许写入继续执行的内部信号。
        /// </summary>
        private TaskCompletionSource<object> ReleaseSignal { get; } =
            new(TaskCreationOptions.RunContinuationsAsynchronously);

        /// <summary>
        /// 释放阻塞的写入操作。
        /// </summary>
        public void Release() => ReleaseSignal.TrySetResult(null);

        /// <inheritdoc />
        public override bool CanRead => false;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => true;
        /// <inheritdoc />
        public override long Length => _inner.Length;
        /// <inheritdoc />
        public override long Position
        {
            get => _inner.Position;
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override Task FlushAsync(CancellationToken cancellationToken)
            => _inner.FlushAsync(cancellationToken);

        /// <inheritdoc />
        public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            WriteStarted.TrySetResult(null);
            await ReleaseSignal.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
            await _inner.WriteAsync(buffer, cancellationToken).ConfigureAwait(false);
        }

        /// <inheritdoc />
        public override Task WriteAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
            => WriteAsync(buffer.AsMemory(offset, count), cancellationToken).AsTask();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
            => _inner.Write(buffer, offset, count);

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin)
            => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => _inner.SetLength(value);

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 将工作簿日期系统标记为 1904 系统并保持输入流可继续使用。
    /// </summary>
    /// <param name="stream">待修改的 XLSX 内存流。</param>
    private static void MarkWorkbookDate1904(MemoryStream stream)
    {
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/workbook.xml");
            Assert.NotNull(entry);
            XDocument document;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, true))
                document = XDocument.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
            var root = document.Root;
            Assert.NotNull(root);
            var workbookProperties = root.Element(root.Name.Namespace + "workbookPr");
            if (workbookProperties == null)
            {
                workbookProperties = new XElement(root.Name.Namespace + "workbookPr");
                root.AddFirst(workbookProperties);
            }
            workbookProperties.SetAttributeValue("date1904", "1");
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("xl/workbook.xml").Open(),
                new UTF8Encoding(false));
            document.Save(writer, System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        stream.Position = 0;
    }

    /// <summary>
    /// 为管线直接测试配置与 ClosedXML 保存格式一致的原生比较规则。
    /// </summary>
    /// <param name="validation">待配置的原生校验规则。</param>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="operation">待验证的比较运算符。</param>
    private static void ConfigureValidation(IXLDataValidation validation, XLAllowedValues allowedValues,
        XLOperator operation)
    {
        validation.AllowedValues = allowedValues;
        validation.Operator = operation;
        validation.MinValue = allowedValues == XLAllowedValues.Time ? "0.5" : "5";
        validation.MaxValue = allowedValues == XLAllowedValues.Time ? "0.75" : "10";
        validation.Value = validation.MinValue;
        validation.IgnoreBlanks = false;
    }

    /// <summary>
    /// 创建指定比较规则的通过值和失败值。
    /// </summary>
    /// <param name="allowedValues">待验证的原生校验类型。</param>
    /// <param name="operation">待验证的比较运算符。</param>
    /// <returns>分别用于通过与失败场景的原始值及文本表示。</returns>
    private static (object ValidRaw, string ValidText, object InvalidRaw, string InvalidText) GetComparisonValues(
        XLAllowedValues allowedValues, XLOperator operation)
    {
        var (valid, invalid) = operation switch
        {
            XLOperator.Between => (5m, 11m),
            XLOperator.NotBetween => (11m, 5m),
            XLOperator.EqualTo => (5m, 6m),
            XLOperator.NotEqualTo => (6m, 5m),
            XLOperator.GreaterThan => (6m, 5m),
            XLOperator.LessThan => (4m, 5m),
            XLOperator.EqualOrGreaterThan => (5m, 4m),
            XLOperator.EqualOrLessThan => (5m, 6m),
            _ => throw new ArgumentOutOfRangeException(nameof(operation), operation, null)
        };
        if (allowedValues == XLAllowedValues.Time)
        {
            var timeValid = DateTime.MinValue.Add(TimeSpan.FromDays((double)(valid / 10m)));
            var timeInvalid = DateTime.MinValue.Add(TimeSpan.FromDays((double)(invalid / 10m)));
            return (timeValid, timeValid.ToString("HH:mm:ss", CultureInfo.InvariantCulture), timeInvalid,
                timeInvalid.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
        }
        if (allowedValues == XLAllowedValues.Date)
        {
            var dateValid = new DateTime(1904, 1, 1).AddDays((double)valid);
            var dateInvalid = new DateTime(1904, 1, 1).AddDays((double)invalid);
            return (dateValid, dateValid.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), dateInvalid,
                dateInvalid.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        }
        if (allowedValues == XLAllowedValues.TextLength)
        {
            var validText = new string('a', (int)valid);
            var invalidText = new string('a', (int)invalid);
            return (validText, validText, invalidText, invalidText);
        }
        return (valid, valid.ToString(CultureInfo.InvariantCulture), invalid,
            invalid.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// 替换 XLSX ZIP 中指定名称的 XML 部件。
    /// </summary>
    /// <param name="source">原始 XLSX 字节。</param>
    /// <param name="entryName">要替换的 ZIP 部件名称。</param>
    /// <param name="content">新的 UTF-8 部件内容。</param>
    /// <returns>替换部件后的 XLSX 字节。</returns>
    private static byte[] ReplaceZipEntry(byte[] source, string entryName, string content)
    {
        using var stream = new MemoryStream();
        stream.Write(source, 0, source.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            archive.GetEntry(entryName)?.Delete();
            var entry = archive.CreateEntry(entryName);
            using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
            writer.Write(content);
        }
        return stream.ToArray();
    }

    /// <summary>
    /// 读取 XLSX 工作簿中指定名称的局部工作表索引。
    /// </summary>
    /// <param name="source">XLSX 工作簿字节。</param>
    /// <param name="name">定义名称。</param>
    /// <returns>可空局部工作表索引。</returns>
    private static string ReadDefinedNameLocalSheetId(byte[] source, string name)
    {
        using var stream = new MemoryStream(source, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        using var workbookStream = archive.GetEntry("xl/workbook.xml")?.Open()
            ?? throw new InvalidOperationException("XLSX 缺少工作簿定义部件。");
        XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var workbook = XDocument.Load(workbookStream);
        var definedName = Assert.Single(workbook.Root?
            .Element(spreadsheet + "definedNames")?
            .Elements(spreadsheet + "definedName")
            .Where(element => (string)element.Attribute("name") == name)
            ?? Enumerable.Empty<XElement>());
        return (string)definedName.Attribute("localSheetId");
    }

    /// <summary>
    /// 流式统计 XLSX 工作表 XML 中的物理单元格元素数量。
    /// </summary>
    /// <param name="source">待统计的 XLSX 字节。</param>
    /// <returns>所有工作表中物理单元格元素的总数。</returns>
    private static long CountPhysicalCells(byte[] source)
    {
        using var stream = new MemoryStream(source, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        long count = 0;
        foreach (var entry in archive.Entries.Where(item =>
                     item.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                     && item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            using var reader = System.Xml.XmlReader.Create(entry.Open(), new System.Xml.XmlReaderSettings
            {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit,
                XmlResolver = null
            });
            while (reader.Read())
            {
                if (reader.NodeType == System.Xml.XmlNodeType.Element
                    && string.Equals(reader.LocalName, "c", StringComparison.Ordinal))
                    count++;
            }
        }
        return count;
    }
}
