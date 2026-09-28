using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Exports;
using Bing.Offices.Styles;
using Bing.Offices.Validations;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 单实体布局和 Provider 能力边界测试。
/// </summary>
public sealed class EntityLayoutProviderTest
{
    /// <summary>
    /// 测试 - NPOI 实体布局应按模板命名锚点写入并读回固定单元格和列表区域。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_NamedAnchors_ShouldRoundTripTemplate()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .CellNamed("单据", "InvoiceCustomer", item => item.Customer)
            .ListRegionNamed("明细", "InvoiceLines", item => item.Lines));
        var source = new Invoice
        {
            Customer = "客户 A",
            Lines = new List<InvoiceLine>
            {
                new InvoiceLine { Code = "A", Quantity = 2 },
                new InvoiceLine { Code = "B", Quantity = 3 }
            }
        };

        using var template = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var invoice = workbook.CreateSheet("单据");
            var lines = workbook.CreateSheet("明细");
            var customer = workbook.CreateName();
            customer.NameName = "InvoiceCustomer";
            customer.SheetIndex = workbook.GetSheetIndex(invoice);
            customer.RefersToFormula = "'单据'!$B$2";
            var list = workbook.CreateName();
            list.NameName = "InvoiceLines";
            list.SheetIndex = workbook.GetSheetIndex(lines);
            list.RefersToFormula = "'明细'!$A$5";
            workbook.Write(template, true);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal("客户 A", workbook.GetSheet("单据").GetRow(1).GetCell(1).StringCellValue);
            Assert.Equal("Code", workbook.GetSheet("明细").GetRow(4).GetCell(0).StringCellValue);
            Assert.Equal("B", workbook.GetSheet("明细").GetRow(6).GetCell(0).StringCellValue);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("客户 A", result.Entity.Customer);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
    }

    /// <summary>
    /// 验证 HSSF 与 XSSF 按命名列表的实际地址校验固定单元格冲突。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="conflict">命名列表是否覆盖固定单元格。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public void Npoi_EntityLayout_NamedListFixedCell_ShouldResolveActualAddress(bool xls, bool conflict)
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("明细", "B2", item => item.Title)
            .ListRegionNamed("明细", "InvoiceLines", item => item.Lines));
        var source = new Invoice
        {
            Title = "采购单",
            Lines = new List<InvoiceLine> { new InvoiceLine { Code = "A", Quantity = 2 } }
        };
        using var template = new MemoryStream();
        using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("明细");
            var name = workbook.CreateName();
            name.NameName = "InvoiceLines";
            name.SheetIndex = workbook.GetSheetIndex(sheet);
            name.RefersToFormula = conflict ? "'明细'!$A$1" : "'明细'!$D$5";
            workbook.Write(template, true);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        if (conflict)
        {
            output.WriteByte(42);
            var originalPosition = output.Position;
            var exportError = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelExporter().ExportForTemplate(source, layout,
                    new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
            Assert.Equal(BingOfficesStage.Plan, exportError.Stage);
            Assert.Equal(new byte[] { 42 }, output.ToArray());
            Assert.Equal(originalPosition, output.Position);
            template.Position = 0;
            var importError = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(template, layout));
            Assert.Equal(BingOfficesStage.Plan, importError.Stage);
            return;
        }

        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("采购单", result.Entity.Title);
        Assert.Equal("A", Assert.Single(result.Entity.Lines).Code);
    }

    /// <summary>
    /// 验证最终 Footer 命名锚点在新工作簿导出后指向真实 marker，且导入仍按 marker 截断明细。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FooterNamed_ShouldCreateNameAndRoundTrip()
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
            new NpoiExcelExporter().ExportEntity(source, layout, output);
            using (var read = new MemoryStream(output.ToArray(), writable: false))
            using (var workbook = WorkbookFactory.Create(read))
            {
                var name = Assert.Single(workbook.GetAllNames().Where(item => item.NameName == "OrderFooter"));
                var expectedAddress = count == 0 ? "$A$2" : "$A$6";
                Assert.Equal("明细!" + expectedAddress,
                    name.RefersToFormula.Replace("'明细'!", "明细!", StringComparison.Ordinal));
                Assert.Equal(-1, name.SheetIndex);
                Assert.Equal("订单合计", workbook.GetSheet("明细")
                    .GetRow(count == 0 ? 1 : 5).GetCell(0).StringCellValue);
                Assert.DoesNotContain(workbook.GetAllNames(), item => item.NameName == "页小计");
            }

            using var input = new MemoryStream(output.ToArray(), writable: false);
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(source.Lines.Select(item => item.Category), result.Entity.Lines.Select(item => item.Category));
            Assert.Equal(source.Lines.Select(item => item.Amount), result.Entity.Lines.Select(item => item.Amount));
        }
    }

    /// <summary>
    /// 验证模板中的 global/local 单格名称会随最终 Footer marker 重定位。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="localScope">名称是否限定在明细工作表。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Npoi_EntityLayout_FooterNamed_ShouldRelocateTemplateName(bool xls, bool localScope)
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
            using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
            {
                workbook.CreateSheet("说明");
                var sheet = workbook.CreateSheet("明细");
                var name = workbook.CreateName();
                name.NameName = "OrderFooter";
                name.SheetIndex = localScope ? workbook.GetSheetIndex(sheet) : -1;
                name.RefersToFormula = "'明细'!$Z$20";
                workbook.Write(template, true);
            }

            template.Position = 0;
            using var output = new MemoryStream();
            new NpoiExcelExporter().ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            using (var read = new MemoryStream(output.ToArray(), writable: false))
            using (var workbook = WorkbookFactory.Create(read))
            {
                var name = Assert.Single(workbook.GetAllNames().Where(item => item.NameName == "OrderFooter"));
                var expectedAddress = count == 0 ? "$A$2" : "$A$6";
                Assert.Equal("明细!" + expectedAddress,
                    name.RefersToFormula.Replace("'明细'!", "明细!", StringComparison.Ordinal));
                Assert.Equal(localScope ? 1 : -1, name.SheetIndex);
            }

            using var input = new MemoryStream(output.ToArray(), writable: false);
            var result = new NpoiExcelImporter().ImportEntity(input, layout);
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
    public void Npoi_EntityLayout_FooterNamed_ShouldRejectDuplicateAnchorsAndClearOnFooter()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder => builder
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
        new NpoiExcelExporter().ExportEntity(new GroupSubtotalDocument(), layout, output);
        using var workbook = WorkbookFactory.Create(new MemoryStream(output.ToArray(), writable: false));
        Assert.DoesNotContain(workbook.GetAllNames(), item => item.NameName == "DiscardedFooter");
        Assert.Equal("最终 marker", workbook.GetSheet("明细").GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 验证最终 Footer 名称冲突和预取消都不会改写既有目标流。
    /// </summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Npoi_EntityLayout_FooterNamed_ShouldPreserveOutputOnConflictAndCancellation(
        bool xls, bool localScope)
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
        using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
        {
            workbook.CreateSheet("说明");
            var sheet = workbook.CreateSheet("明细");
            var name = workbook.CreateName();
            name.NameName = "OrderFooter";
            name.SheetIndex = localScope ? workbook.GetSheetIndex(sheet) : -1;
            name.RefersToFormula = "'明细'!$Z$20:$Z$21";
            workbook.Write(template, true);
            template.Position = 0;
            using var output = new MemoryStream();
            output.Write(original, 0, original.Length);
            output.Position = 0;
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelExporter().ExportForTemplate(source, layout,
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
            new NpoiExcelExporter().ExportEntityAsync(source, layout, canceledOutput, cancellation.Token));
        Assert.Equal(original, canceledOutput.ToArray());
    }

    /// <summary>
    /// 测试 - NPOI 实体布局缺少或指向多单元格命名锚点时应在计划阶段失败。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_NamedAnchor_ShouldRejectInvalidTemplateRange()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder =>
            builder.CellNamed("单据", "InvoiceCustomer", item => item.Customer));
        using var template = new MemoryStream();
        using (var workbook = new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("单据");
            var name = workbook.CreateName();
            name.NameName = "InvoiceCustomer";
            name.SheetIndex = workbook.GetSheetIndex(sheet);
            name.RefersToFormula = "'单据'!$B$2:$C$2";
            workbook.Write(template, true);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(new Invoice(), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Contains("单个单元格", exception.Message);
    }

    /// <summary>
    /// 测试 - NPOI 实体布局应拒绝重复命名锚点和命名列表的绝对结束地址。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_NamedAnchor_ShouldRejectDuplicateAndAbsoluteEnd()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder => builder
            .CellNamed("单据", "InvoiceCustomer", item => item.Customer)
            .CellNamed("单据", "InvoiceCustomer", item => item.Title)));

        var exception = Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder =>
            builder.ListRegionNamed("明细", "InvoiceLines", item => item.Lines,
                region => region.End("D20"))));
        Assert.Contains("绝对 End", exception.Message);
    }

    /// <summary>
    /// 测试 - NPOI 实体布局计算列应接收行上下文、保持物理顺序并在导入时跳过。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_CalculatedColumns_ShouldUseRowContextAndSkipOnImport()
    {
        var contexts = new List<(int Index, int RowNumber, int ColumnNumber, int ItemCount, string Sheet)>();
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
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
        var source = new Invoice
        {
            Lines = new List<InvoiceLine>
            {
                new InvoiceLine { Code = "A", Quantity = 2 },
                new InvoiceLine { Code = "B", Quantity = 3 }
            }
        };

        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, exported);
        Assert.Equal(new[] { (0, 3, 2, 2, "明细"), (1, 4, 2, 2, "明细") }, contexts);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("序号", sheet.GetRow(1).GetCell(1).StringCellValue);
            Assert.Equal("Code", sheet.GetRow(1).GetCell(2).StringCellValue);
            Assert.Equal("金额", sheet.GetRow(1).GetCell(4).StringCellValue);
            Assert.Equal(1d, sheet.GetRow(2).GetCell(1).NumericCellValue);
            Assert.Equal(20d, sheet.GetRow(2).GetCell(4).NumericCellValue);
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(item => item.Quantity));
        Assert.Equal(2, contexts.Count);
    }

    /// <summary>
    /// 测试 - NPOI 计算列委托异常应包含列坐标并保留取消语义。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_CalculatedColumnFailure_ShouldReportCoordinates()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .CalculatedColumn<decimal>("amount", "金额", _ => throw new InvalidOperationException("计算失败"))));

        using var output = new MemoryStream();
        var exception = Assert.Throws<BingOfficesExportException>(() =>
            new NpoiExcelExporter().ExportEntity(new Invoice
            {
                Lines = new List<InvoiceLine> { new InvoiceLine { Code = "A", Quantity = 1 } }
            }, layout, output));
        Assert.Equal("明细", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(3, exception.ColumnIndex);
        Assert.Equal("amount", exception.PropertyName);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 测试 - NPOI 应支持固定单元格、合并锚点和列表区域的导出与导入。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ShouldRoundTripFixedCellsAndListRegion()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "B2", item => item.Number)
            .Cell("单据", "B3", item => item.Customer)
            .Merge("单据", "A1:B1")
            .Cell("单据", "B1", item => item.Title)
            .ListRegion("明细", "A1", item => item.Lines));
        var source = new Invoice
        {
            Number = 1001,
            Customer = "客户 A",
            Title = "销售单",
            Lines = new List<InvoiceLine>
            {
                new InvoiceLine { Code = "A", Quantity = 2 },
                new InvoiceLine { Code = "B", Quantity = 3 }
            }
        };

        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var invoiceSheet = workbook.GetSheet("单据");
            Assert.Equal("销售单", invoiceSheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal(1001d, invoiceSheet.GetRow(1).GetCell(1).NumericCellValue);
            Assert.Equal(1, invoiceSheet.NumMergedRegions);
            var listSheet = workbook.GetSheet("明细");
            Assert.Equal("Code", listSheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal("B", listSheet.GetRow(2).GetCell(0).StringCellValue);
        }

        using var importSource = new MemoryStream(exported.ToArray());
        var result = new NpoiExcelImporter().ImportEntity(importSource, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal(1001, result.Entity.Number);
        Assert.Equal("客户 A", result.Entity.Customer);
        Assert.Equal("销售单", result.Entity.Title);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
        Assert.Equal(new[] { 2, 3 }, result.Entity.Lines.Select(item => item.Quantity));
    }

    /// <summary>
    /// 测试 - NPOI 实体列表应保持既有单字典动态列的读写能力。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ShouldRoundTripLegacyDynamicDictionary()
    {
        var layout = ExcelEntity.Layout<LegacyDynamicDocument>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region.Mapping(
                new ExcelMappingConfiguration
                {
                    DynamicColumns =
                    {
                        new ExcelMappingDynamicColumnConfiguration
                        {
                            Key = "color", Title = "颜色", DataTypeName = "string"
                        }
                    }
                })));
        var source = new LegacyDynamicDocument
        {
            Lines = new List<LegacyDynamicLine>
            {
                new LegacyDynamicLine
                {
                    Code = "A",
                    Values = new Dictionary<string, object> { ["color"] = "红" }
                }
            }
        };

        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("颜色", sheet.GetRow(0).GetCell(1).StringCellValue);
            Assert.Equal("红", sheet.GetRow(1).GetCell(1).StringCellValue);
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("红", result.Entity.Lines[0].Values["color"]);
    }

    /// <summary>
    /// 测试 - NPOI 实体显式动态列组应执行内置校验规则。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ExplicitDynamicValidation_ShouldReportStructuredError()
    {
        var layout = ExcelEntity.Layout<ExplicitDynamicDocument>(builder => builder
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
        new NpoiExcelExporter().ExportEntity(new ExplicitDynamicDocument
        {
            Lines = new List<ExplicitDynamicLine>
            {
                new() { Code = "A", Fields = new Dictionary<string, object>
                    { ["weight"] = 5m } }
            }
        }, layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            workbook.GetSheet("明细").GetRow(1).GetCell(1).SetCellValue(99d);
            using var input = new MemoryStream();
            workbook.Write(input, false);
            inputBytes = input.ToArray();
        }

        using var source = new MemoryStream(inputBytes, writable: false);
        var result = new NpoiExcelImporter().ImportEntity(source, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors.Where(item =>
            item.PropertyName == nameof(ExplicitDynamicLine.Fields)));
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
    }

    /// <summary>
    /// 测试 - NPOI 实体显式动态列应覆盖全部内置校验规则和转换失败。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ExplicitDynamicValidation_ShouldCoverAllRules()
    {
        void AssertFailure(ExcelDynamicColumnDefinition definition, IReadOnlyList<object> values,
            Action<ICell> mutate, ExcelImportErrorCode code, int rowIndex = 2)
        {
            var layout = ExcelEntity.Layout<ExplicitDynamicDocument>(builder => builder
                .ListRegion("动态验证", "A1", item => item.Lines, region => region
                    .DynamicColumnGroup("商品", item => item.Fields, new[] { definition })));
            using var exported = new MemoryStream();
            new NpoiExcelExporter().ExportEntity(new ExplicitDynamicDocument
            {
                Lines = values.Select((value, index) => new ExplicitDynamicLine
                {
                    Code = $"L{index + 1}",
                    Fields = new Dictionary<string, object> { [definition.Key] = value }
                }).ToList()
            }, layout, exported);
            exported.Position = 0;
            byte[] inputBytes;
            using (var workbook = WorkbookFactory.Create(exported))
            {
                var sheet = workbook.GetSheet("动态验证");
                var columnIndex = Enumerable.Range(0, sheet.GetRow(0).LastCellNum)
                    .Single(index => sheet.GetRow(0).GetCell(index).StringCellValue == definition.Title);
                mutate(sheet.GetRow(rowIndex - 1).GetCell(columnIndex));
                using var input = new MemoryStream();
                workbook.Write(input, false);
                inputBytes = input.ToArray();
            }

            using var source = new MemoryStream(inputBytes, writable: false);
            var result = new NpoiExcelImporter().ImportEntity(source, layout);
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
        }, new object[] { "有效" }, cell => cell.SetCellValue(string.Empty),
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "length", Title = "长度", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "maxLength", MaxLength = 3 }
            }
        }, new object[] { "有效" }, cell => cell.SetCellValue("超长数据"),
            ExcelImportErrorCode.MaxLength);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "regex", Title = "正则", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "regex", Pattern = "^OK-" }
            }
        }, new object[] { "OK-001" }, cell => cell.SetCellValue("BAD"),
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "date", Title = "日期", DataType = typeof(DateTime),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "date", Format = "yyyy-MM-dd" }
            }
        }, new object[] { new DateTime(2024, 1, 1) }, cell => cell.SetCellValue("2024/01/01"),
            ExcelImportErrorCode.Validation);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "max", Title = "最大值", DataType = typeof(decimal),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "maxValue", MaxValue = 10 }
            }
        }, new object[] { 5m }, cell => cell.SetCellValue(11d), ExcelImportErrorCode.MaxValue);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "unique", Title = "唯一", DataType = typeof(string),
            ValidationRules = new[]
            {
                new ExcelMappingDynamicValidationConfiguration { Name = "unique", IgnoreEmpty = false }
            }
        }, new object[] { "U1", "U2" }, cell => cell.SetCellValue("U1"),
            ExcelImportErrorCode.Validation, rowIndex: 3);
        AssertFailure(new ExcelDynamicColumnDefinition
        {
            Key = "number", Title = "数值", DataType = typeof(decimal)
        }, new object[] { 5m }, cell => cell.SetCellValue("不是数字"),
            ExcelImportErrorCode.ValueConversion);
    }

    /// <summary>
    /// 测试 - NPOI 实体导入可选择收集当前行的全部校验错误。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ValidationFailureMode_ShouldControlRowErrors()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(CreateValidationRoot(), layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("验证");
            sheet.GetRow(1).GetCell(0).SetCellValue(string.Empty);
            sheet.GetRow(1).GetCell(3).SetCellValue("BAD");
            using var input = new MemoryStream();
            workbook.Write(input, false);
            inputBytes = input.ToArray();
        }

        var stopResult = new NpoiExcelImporter().ImportEntity(
            new MemoryStream(inputBytes, writable: false), layout);
        var continueResult = new NpoiExcelImporter().ImportEntity(
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
    /// 测试 - NPOI 实体异步和模板异步导入应遵守继续收集行校验错误的策略。
    /// </summary>
    [Fact]
    public async Task Npoi_EntityLayout_ValidationFailureMode_ShouldContinueForAsyncAndTemplate()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(CreateValidationRoot(), layout, exported);
        exported.Position = 0;
        byte[] inputBytes;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("验证");
            sheet.GetRow(1).GetCell(0).SetCellValue(string.Empty);
            sheet.GetRow(1).GetCell(3).SetCellValue("BAD");
            using var input = new MemoryStream();
            workbook.Write(input, false);
            inputBytes = input.ToArray();
        }

        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue);
        var asyncResult = await new NpoiExcelImporter().ImportEntityAsync(
            new MemoryStream(inputBytes, writable: false), layout, options);
        Assert.True(asyncResult.Errors.Count(error => error.RowIndex == 2) >= 2);
        Assert.Contains(asyncResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Required));
        Assert.Contains(asyncResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Code));

        using var template = new MemoryStream(inputBytes, writable: false);
        var templateResult = await new NpoiExcelImporter().ImportForTemplateAsync(
            new MemoryStream(inputBytes, writable: false), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), options);
        Assert.True(templateResult.Errors.Count(error => error.RowIndex == 2) >= 2);
        Assert.Contains(templateResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Required));
        Assert.Contains(templateResult.Errors, error => error.RowIndex == 2
            && error.PropertyName == nameof(EntityValidationLine.Code));
    }

    /// <summary>
    /// 测试 - NPOI 实体列表应支持多个动态列组和相对尾部聚合。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ShouldRoundTripDynamicGroupsAndFooter()
    {
        var layout = ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .DynamicColumnGroup("商品", item => item.ProductFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "color", Title = "颜色", DataType = typeof(string) }
                })
                .DynamicColumnGroup("扩展", item => item.ExtensionFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "weight", Title = "重量", DataType = typeof(decimal) }
                })
                .UnknownDynamicValues(ExcelUnknownDynamicValuePolicy.Fail)
                .Footer("合计", footer => footer
                    .Cell("D1", items => items.Sum(item => item.Amount), numberFormat: "0.00")
                    .Merge("A1:C1"))));
        var source = new DynamicDocument
        {
            Lines = new List<DynamicLine>
            {
                new DynamicLine
                {
                    Code = "A",
                    Amount = 1.25m,
                    ProductFields = new Dictionary<string, object> { ["color"] = "红" },
                    ExtensionFields = new Dictionary<string, object> { ["weight"] = 2.5m }
                },
                new DynamicLine
                {
                    Code = "B",
                    Amount = 2.25m,
                    ProductFields = new Dictionary<string, object> { ["color"] = "蓝" },
                    ExtensionFields = new Dictionary<string, object> { ["weight"] = 3.5m }
                }
            }
        };

        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(new[] { "Code", "Amount", "颜色", "重量" },
                Enumerable.Range(0, 4).Select(index => sheet.GetRow(0).GetCell(index).StringCellValue));
            Assert.Equal("红", sheet.GetRow(1).GetCell(2).StringCellValue);
            Assert.Equal(2.5d, sheet.GetRow(1).GetCell(3).NumericCellValue);
            Assert.Equal("合计", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.Equal(3.5d, sheet.GetRow(3).GetCell(3).NumericCellValue, 2);
            Assert.Equal(1, sheet.NumMergedRegions);
        }

        using var input = new MemoryStream(exported.ToArray());
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Code));
        Assert.Equal("红", result.Entity.Lines[0].ProductFields["color"]);
        Assert.Equal(2.5m, Convert.ToDecimal(result.Entity.Lines[0].ExtensionFields["weight"]));
    }

    /// <summary>
    /// 测试 - NPOI 实体列表应按连续分组写入小计，并在导入时跳过小计行。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_GroupSubtotal_ShouldRoundTripAndSkipSubtotalRows()
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("小计", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.Equal(4d, sheet.GetRow(3).GetCell(1).NumericCellValue, 2);
            Assert.Equal("小计", sheet.GetRow(5).GetCell(0).StringCellValue);
            Assert.Equal(4d, sheet.GetRow(5).GetCell(1).NumericCellValue, 2);
            Assert.Equal("总计", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal(8d, sheet.GetRow(6).GetCell(1).NumericCellValue, 2);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "A", "B" }, result.Entity.Lines.Select(item => item.Category));
        Assert.Equal(new[] { 1.25m, 2.75m, 4m }, result.Entity.Lines.Select(item => item.Amount));
    }

    /// <summary>
    /// 测试 - NPOI 应支持属性式固定单元格、多个动态字典组和尾部聚合。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_AttributesGroupsAndFooter_ShouldRoundTrip()
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
                    .Cell("B1", items => items.Sum(item => item.Quantity), numberFormat: "0")
                    .Cell("C1", items => items.Sum(item => item.Amount), numberFormat: "0.00")
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("采购单");
            Assert.Equal("PO-001", sheet.GetRow(1).GetCell(1).StringCellValue);
            Assert.Equal(new[] { "Name", "Quantity", "Amount", "商品字段", "产品字段" },
                Enumerable.Range(0, 5).Select(index => sheet.GetRow(3).GetCell(index).StringCellValue));
            Assert.Equal("合计", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal(5d, sheet.GetRow(6).GetCell(1).NumericCellValue);
            Assert.Equal(20d, sheet.GetRow(6).GetCell(2).NumericCellValue);
            Assert.Contains(sheet.MergedRegions, range => range.FormatAsString() == "D7:E7");
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("PO-001", result.Entity.Code);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(item => item.Name));
        Assert.Equal("G2", result.Entity.Lines[1].Goods["goods"]);
        Assert.Equal("P", result.Entity.Lines[0].Product["product"]);
    }

    /// <summary>
    /// 测试 - NPOI XSSF 实体列表应按明细行数写入水平分页符。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_PageBreak_ShouldWriteXssfRowBreaks()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageBreak(2)));
        var source = new Invoice
        {
            Lines = Enumerable.Range(1, 5)
                .Select(index => new InvoiceLine { Code = $"C{index}", Quantity = index })
                .ToList()
        };

        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);

        Assert.Equal(new[] { 2, 4 }, workbook.GetSheet("明细").RowBreaks);
    }

    /// <summary>
    /// 测试 - NPOI HSSF 模板实体列表应保留分页语义并写入水平分页符。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_PageBreak_ShouldWriteHssfTemplateRowBreaks()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageBreak(2)));
        var source = new Invoice
        {
            Lines = Enumerable.Range(1, 5)
                .Select(index => new InvoiceLine { Code = $"C{index}", Quantity = index })
                .ToList()
        };

        using var template = new MemoryStream();
        using (var workbook = new HSSFWorkbook())
        {
            workbook.CreateSheet("明细");
            workbook.Write(template, true);
        }

        template.Position = 0;
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using var workbookOutput = Assert.IsType<HSSFWorkbook>(WorkbookFactory.Create(output));

        Assert.Equal(new[] { 2, 4 }, workbookOutput.GetSheet("明细").RowBreaks);
    }

    /// <summary>
    /// 验证签字区推移末页明细后仍在写入目标流前拒绝越界。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="physical">是否验证工作表物理边界。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PageSignatures_ShouldRejectShiftedDetailOverflow(bool xls, bool physical)
    {
        var start = physical ? (xls ? "A65531" : "A1048571") : "A1";
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
        using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
        {
            workbook.CreateSheet("明细");
            workbook.Write(template, true);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        var original = new byte[] { 1, 2, 3, 4 };
        output.Write(original, 0, original.Length);
        output.Position = 0;
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(source, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证分页签字区的完整内容、合并、模板保留和导入边界。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 模板。</param>
    /// <param name="header">是否包含表头。</param>
    /// <param name="asynchronous">是否通过异步入口执行。</param>
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public async Task PageSignatures_ShouldRoundTripTemplate(bool xls, bool header, bool asynchronous)
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
            using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("明细");
                sheet.SetColumnWidth(0, 24 * 256);
                foreach (var row in signatureRows)
                {
                    sheet.CreateRow(row + headerRows).HeightInPoints = 29;
                    sheet.AddMergedRegion(new CellRangeAddress(row + headerRows, row + headerRows + 1, 0, 1));
                }
                workbook.Write(template, true);
            }
            template.Position = 0;
            using var output = new MemoryStream();
            var exporter = new NpoiExcelExporter();
            if (asynchronous)
                await exporter.ExportForTemplateAsync(source, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            else
                exporter.ExportForTemplate(source, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
            Assert.True(template.CanRead);
            Assert.True(output.CanWrite);
            using (var read = new MemoryStream(output.ToArray(), writable: false))
            using (var workbook = WorkbookFactory.Create(read))
            {
                Assert.Equal(xls, workbook is HSSFWorkbook);
                var sheet = workbook.GetSheet("明细");
                var actual = Enumerable.Range(headerRows, expected.Length).Select(row =>
                {
                    var cells = sheet.GetRow(row);
                    var left = cells?.GetCell(0)?.ToString() ?? string.Empty;
                    var right = cells?.GetCell(1);
                    return left + "|" + (right?.CellType == CellType.Numeric
                        ? right.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
                        : right?.ToString() ?? string.Empty);
                }).ToArray();
                Assert.Equal(expected, actual);
                Assert.Equal(signatureRows.Length, sheet.NumMergedRegions);
                Assert.Equal(24 * 256, sheet.GetColumnWidth(0));
                foreach (var row in signatureRows)
                {
                    Assert.Contains(Enumerable.Range(0, sheet.NumMergedRegions).Select(sheet.GetMergedRegion),
                        merge => merge.FirstRow == row + headerRows && merge.LastRow == row + headerRows + 1
                            && merge.FirstColumn == 0 && merge.LastColumn == 1);
                    Assert.Equal(29f, sheet.GetRow(row + headerRows).HeightInPoints);
                    Assert.True(workbook.GetFontAt(sheet.GetRow(row + headerRows).GetCell(0).CellStyle.FontIndex).IsBold);
                }
                Assert.Equal(count == 5 ? new[] { 5 + headerRows, 11 + headerRows } : Array.Empty<int>(), sheet.RowBreaks);
            }
            using var input = new MemoryStream(output.ToArray(), writable: false);
            var importer = new NpoiExcelImporter();
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
    /// 测试 - NPOI 分页小计应按页聚合、写入分页符并在导入时跳过小计行。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_PageSubtotal_ShouldRoundTripAndBreakAfterSubtotal()
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("页小计", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.Equal(3d, sheet.GetRow(3).GetCell(1).NumericCellValue, 2);
            Assert.Equal("总计", sheet.GetRow(6).GetCell(0).StringCellValue);
            Assert.Equal(10d, sheet.GetRow(6).GetCell(1).NumericCellValue, 2);
            Assert.Equal(new[] { 3 }, sheet.RowBreaks);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { 1m, 2m, 3m, 4m }, result.Entity.Lines.Select(item => item.Amount));
    }

    /// <summary>
    /// 测试 - 分页小计 marker 缺失或位置错误时应返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_PageSubtotalMarkerErrors_ShouldBeStructured()
    {
        var layout = ExcelEntity.Layout<GroupSubtotalDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("页小计")
                .Footer("总计")));

        static byte[] CreateWorkbook(params string[] rows)
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("明细");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Category");
            sheet.GetRow(0).CreateCell(1).SetCellValue("Amount");
            for (var index = 0; index < rows.Length; index++)
            {
                var row = sheet.CreateRow(index + 1);
                row.CreateCell(0).SetCellValue(rows[index]);
                row.CreateCell(1).SetCellValue(index + 1);
            }
            using var stream = new MemoryStream();
            workbook.Write(stream, false);
            return stream.ToArray();
        }

        using (var missing = new MemoryStream(CreateWorkbook("A", "B", "C", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(missing, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("缺少分页小计标记", exception.Message);
        }
        using (var terminal = new MemoryStream(CreateWorkbook("A", "B", "页小计", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(terminal, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("分页小计之后缺少明细", exception.Message);
        }
        using (var invalid = new MemoryStream(CreateWorkbook("A", "页小计", "B", "C", "总计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(invalid, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("分页小计标记位置无效", exception.Message);
        }
    }

    /// <summary>
    /// 测试 - 分页行数必须为正数且不能与连续分组小计同时配置。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_PageBreak_ShouldRejectInvalidCombinations()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => ExcelEntity.Layout<Invoice>(builder =>
            builder.ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageBreak(0))));

        Assert.Throws<InvalidOperationException>(() => ExcelEntity.Layout<GroupSubtotalDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .GroupSubtotal(item => item.Category, "小计"))));

        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder =>
            builder.ListRegion("明细", "A1", invoice => invoice.Lines,
                region => region.PageSubtotal("页小计"))));
    }

    /// <summary>
    /// 测试 - 商品档案应支持三个动态字典组、空字典和未知键忽略。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ProductArchive_ShouldRoundTripThreeDynamicGroups()
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
        new NpoiExcelExporter().ExportEntity(source, layout, exported);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = workbook.GetSheet("商品档案");
            Assert.Equal(new[] { "Code", "Name", "品牌", "重量", "启用" },
                Enumerable.Range(0, 5).Select(index => sheet.GetRow(0).GetCell(index).StringCellValue));
            Assert.Equal("Bing", sheet.GetRow(1).GetCell(2).StringCellValue);
            Assert.Equal("2.5", sheet.GetRow(1).GetCell(3).StringCellValue);
            Assert.Equal("true", sheet.GetRow(1).GetCell(4).StringCellValue);
            Assert.Null(sheet.GetRow(1).GetCell(5));
        }

        using var input = new MemoryStream(exported.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
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
    /// 测试 - 赠品单无表头且无明细时应只输出尾部标记并成功读回。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_GiftOrderWithoutHeaderAndEmptyDetails_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("赠品单", "A1", item => item.Items, region => region
                .Header(false)
                .Footer("合计", footer => footer.Cell("B1", items => items.Count))));
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(new GiftOrder(), layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("赠品单");
            Assert.Equal("合计", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal(0d, sheet.GetRow(0).GetCell(1).NumericCellValue);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Empty(result.Entity.Items);
    }

    /// <summary>
    /// 测试 - HSSF 与 XSSF 模板的 Footer 公式应写为公式单元格且不参与明细导入。
    /// </summary>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(true, 2)]
    [InlineData(false, 0)]
    [InlineData(false, 2)]
    public void Npoi_EntityLayout_FooterFormula_ShouldRoundTripForHssfAndXssf(bool useHssf, int lineCount)
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Header(false)
                .Footer("合计", footer => footer
                    .Formula("C1", "=SUM(B1:B2)", numberFormat: "0.00"))));
        var templateBytes = CreateTemplate();
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };

        using var template = new MemoryStream(templateBytes, writable: false);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal(useHssf, workbook is HSSFWorkbook);
            var sheet = workbook.GetSheet("明细");
            var footerRowIndex = lineCount;
            var footerRow = sheet.GetRow(footerRowIndex);
            Assert.Equal("合计", footerRow.GetCell(0).StringCellValue);
            var formulaCell = footerRow.GetCell(2);
            Assert.Equal(CellType.Formula, formulaCell.CellType);
            Assert.Equal("SUM(B1:B2)", formulaCell.CellFormula);
            Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat(formulaCell.CellStyle.DataFormat));
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        using var importTemplate = new MemoryStream(templateBytes, writable: false);
        var result = new NpoiExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(importTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
        {
            Assert.Equal("L1", result.Entity.Lines[0].Code);
            Assert.Equal(1.25m, result.Entity.Lines[0].Amount);
        }

        byte[] CreateTemplate()
        {
            IWorkbook workbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
            using (workbook)
            using (var stream = new MemoryStream())
            {
                workbook.CreateSheet("明细");
                workbook.Write(stream);
                return stream.ToArray();
            }
        }
    }

    /// <summary>
    /// 测试 - Footer 公式必须在布局构建阶段以等号开头并包含公式内容。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FooterFormula_ShouldRejectInvalidFormulaAtBuild()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Footer("合计", footer => footer.Formula("C1", "SUM(B1:B2)")))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Footer("合计", footer => footer.Formula("C1", "")))));
    }

    /// <summary>
    /// 测试 - HSSF 与 XSSF 拒绝无效公式时应保留调用方流内容。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Npoi_EntityLayout_FooterFormula_ShouldPreserveOutputOnInvalidFormula(bool useHssf)
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .Footer("合计", footer => footer.Formula("C1", "=SUM("))));
        IWorkbook workbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
        byte[] templateBytes;
        using (workbook)
        using (var templateOutput = new MemoryStream())
        {
            workbook.CreateSheet("明细");
            workbook.Write(templateOutput);
            templateBytes = templateOutput.ToArray();
        }
        using var template = new MemoryStream(templateBytes, writable: false);
        var original = new byte[] { 1, 2, 3, 4 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);

        Assert.ThrowsAny<Exception>(() => new NpoiExcelExporter().ExportForTemplate(
            new FooterFormulaDocument(), layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
    }

    /// <summary>
    /// 测试 - HSSF 与 XSSF 的连续明细求和公式应按空行和公式相对行定位。
    /// </summary>
    /// <param name="useHssf">是否使用 HSSF 模板。</param>
    /// <param name="lineCount">明细行数。</param>
    /// <param name="setGapBeforeFormula">是否在公式前设置空行数。</param>
    [Theory]
    [InlineData(true, 0, true)]
    [InlineData(true, 0, false)]
    [InlineData(true, 3, true)]
    [InlineData(true, 3, false)]
    [InlineData(false, 0, true)]
    [InlineData(false, 0, false)]
    [InlineData(false, 3, true)]
    [InlineData(false, 3, false)]
    public void Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldRoundTrip(
        bool useHssf, int lineCount, bool setGapBeforeFormula)
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
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        IWorkbook templateWorkbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
        byte[] templateBytes;
        using (templateWorkbook)
        using (var stream = new MemoryStream())
        {
            templateWorkbook.CreateSheet("明细");
            templateWorkbook.Write(stream);
            templateBytes = stream.ToArray();
        }

        using var template = new MemoryStream(templateBytes, writable: false);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal(useHssf, workbook is HSSFWorkbook);
            var sheet = workbook.GetSheet("明细");
            var markerRow = lineCount + 1;
            var formulaRow = markerRow + 1;
            Assert.Equal("合计", sheet.GetRow(markerRow).GetCell(0).StringCellValue);
            var formula = sheet.GetRow(formulaRow).GetCell(2);
            Assert.Equal(CellType.Formula, formula.CellType);
            Assert.Equal(lineCount == 0
                ? "0"
                : $"SUM(INDEX(B:B,ROW()-{lineCount + 2}):INDEX(B:B,ROW()-3))",
                formula.CellFormula);
            Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat(formula.CellStyle.DataFormat));
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        using var importTemplate = new MemoryStream(templateBytes, writable: false);
        var result = new NpoiExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(importTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
        {
            Assert.Equal("L1", result.Entity.Lines[0].Code);
            Assert.Equal(1.25m, result.Entity.Lines[0].Amount);
        }
    }

    /// <summary>
    /// 测试 - 分页小计可使用连续明细求和公式，且导入会跳过小计行。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForPageSubtotal()
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("页小计", sheet.GetRow(2).GetCell(0).StringCellValue);
            Assert.Equal("SUM(INDEX(B:B,ROW()-2):INDEX(B:B,ROW()-1))",
                sheet.GetRow(2).GetCell(2).CellFormula);
            Assert.Equal("总计", sheet.GetRow(4).GetCell(0).StringCellValue);
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "L1", "L2", "L3" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 测试 - 分组小计可为每组写入连续明细求和公式并保持导入边界。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldWorkForGroupSubtotal()
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal("小计", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.Equal("SUM(INDEX(B:B,ROW()-2):INDEX(B:B,ROW()-1))",
                sheet.GetRow(3).GetCell(2).CellFormula);
            Assert.Equal("小计", sheet.GetRow(5).GetCell(0).StringCellValue);
            Assert.Equal("SUM(INDEX(B:B,ROW()-1):INDEX(B:B,ROW()-1))",
                sheet.GetRow(5).GetCell(2).CellFormula);
            Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat(sheet.GetRow(3).GetCell(2).CellStyle.DataFormat));
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "A", "B" }, result.Entity.Lines.Select(item => item.Category));
    }

    /// <summary>
    /// 测试 - 非法求和列及最终 Footer 与中间小计冲突时应在布局阶段拒绝。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldRejectInvalidColumnsAndFinalFooterConflicts()
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
    /// 测试 - 连续求和公式尾部越过声明边界时应保留调用方输出流。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Npoi_EntityLayout_FormulaSumContiguousRowsAbove_ShouldPreserveOutputOnPlanFailure(bool useHssf)
    {
        var layout = ExcelEntity.Layout<FooterFormulaDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .End("B10")
                .Footer("合计", footer => footer.FormulaSumContiguousRowsAbove("C2", "B"))));
        IWorkbook workbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
        byte[] templateBytes;
        using (workbook)
        using (var stream = new MemoryStream())
        {
            workbook.CreateSheet("明细");
            workbook.Write(stream);
            templateBytes = stream.ToArray();
        }
        using var template = new MemoryStream(templateBytes, writable: false);
        var original = new byte[] { 4, 3, 2, 1 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(new FooterFormulaDocument(), layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(original, output.ToArray());
    }

    /// <summary>
    /// 测试 - 最终 Footer 求和公式应跳过分页小计的间隔行和多行内容。
    /// </summary>
    /// <param name="useHssf">是否使用 HSSF 模板。</param>
    /// <param name="lineCount">明细行数。</param>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(true, 4)]
    [InlineData(false, 0)]
    [InlineData(false, 4)]
    public void Npoi_EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipPageSubtotals(
        bool useHssf, int lineCount)
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
        var source = new FooterFormulaDocument
        {
            Lines = Enumerable.Range(1, lineCount)
                .Select(index => new FooterFormulaLine { Code = $"L{index}", Amount = index + 0.25m })
                .ToList()
        };
        IWorkbook templateWorkbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
        byte[] templateBytes;
        using (templateWorkbook)
        using (var stream = new MemoryStream())
        {
            templateWorkbook.CreateSheet("明细");
            templateWorkbook.Write(stream);
            templateBytes = stream.ToArray();
        }

        using var template = new MemoryStream(templateBytes, writable: false);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            if (lineCount == 0)
            {
                Assert.Equal("总计", sheet.GetRow(1).GetCell(0).StringCellValue);
                Assert.Equal("0", sheet.GetRow(2).GetCell(2).CellFormula);
                Assert.Equal("复核签字", sheet.GetRow(2).GetCell(3).StringCellValue);
            }
            else
            {
                Assert.Equal("页小计", sheet.GetRow(3).GetCell(0).StringCellValue);
                Assert.Equal("经办人签字", sheet.GetRow(4).GetCell(3).StringCellValue);
                Assert.Equal("总计", sheet.GetRow(8).GetCell(0).StringCellValue);
                var formula = sheet.GetRow(9).GetCell(2);
                Assert.Equal(CellType.Formula, formula.CellType);
                Assert.Equal("SUM(B1:B2,B6:B7)", formula.CellFormula);
                Assert.Equal("复核签字", sheet.GetRow(9).GetCell(3).StringCellValue);
                Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat(formula.CellStyle.DataFormat));
            }
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        using var importTemplate = new MemoryStream(templateBytes, writable: false);
        var result = new NpoiExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(importTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
            Assert.Equal(new[] { "L1", "L2", "L3", "L4" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 测试 - 最终 Footer 求和公式应只汇总分组明细并跳过多行分组小计。
    /// </summary>
    /// <param name="useHssf">是否使用 HSSF 模板。</param>
    /// <param name="lineCount">明细行数。</param>
    [Theory]
    [InlineData(true, 0)]
    [InlineData(true, 4)]
    [InlineData(false, 0)]
    [InlineData(false, 4)]
    public void Npoi_EntityLayout_FormulaSumDetailRowsAbove_ShouldSkipGroupSubtotals(
        bool useHssf, int lineCount)
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
        IWorkbook templateWorkbook = useHssf ? new HSSFWorkbook() : new XSSFWorkbook();
        byte[] templateBytes;
        using (templateWorkbook)
        using (var stream = new MemoryStream())
        {
            templateWorkbook.CreateSheet("明细");
            templateWorkbook.Write(stream);
            templateBytes = stream.ToArray();
        }

        using var template = new MemoryStream(templateBytes, writable: false);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            if (lineCount == 0)
            {
                Assert.Equal("总计", sheet.GetRow(1).GetCell(0).StringCellValue);
                Assert.Equal("0", sheet.GetRow(2).GetCell(2).CellFormula);
            }
            else
            {
                Assert.Equal("分组小计", sheet.GetRow(3).GetCell(0).StringCellValue);
                Assert.Equal("组内签字", sheet.GetRow(4).GetCell(3).StringCellValue);
                Assert.Equal("分组小计", sheet.GetRow(8).GetCell(0).StringCellValue);
                Assert.Equal("组内签字", sheet.GetRow(9).GetCell(3).StringCellValue);
                Assert.Equal("总计", sheet.GetRow(11).GetCell(0).StringCellValue);
                var formula = sheet.GetRow(12).GetCell(2);
                Assert.Equal(CellType.Formula, formula.CellType);
                Assert.Equal("SUM(B1:B2,B6:B7)", formula.CellFormula);
                Assert.Equal("0.00", workbook.CreateDataFormat().GetFormat(formula.CellStyle.DataFormat));
            }
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        using var importTemplate = new MemoryStream(templateBytes, writable: false);
        var result = new NpoiExcelImporter().ImportForTemplate(input, layout,
            new ExcelEntityTemplateOptions(importTemplate));
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(lineCount, result.Entity.Lines.Count);
        if (lineCount > 0)
            Assert.Equal(new[] { "A", "A", "B", "B" }, result.Entity.Lines.Select(line => line.Category));
    }

    /// <summary>
    /// 测试 - 明细段求和公式不能配置在分页或分组中间小计。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FormulaSumDetailRowsAbove_ShouldRejectIntermediateSubtotalAtBuild()
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
    /// 测试 - 盘点单应支持多工作表、多列表区域和固定汇总单元格。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_Inventory_ShouldRoundTripMultipleSheetsAndRegions()
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
        new NpoiExcelExporter().ExportEntity(source, layout, output);
        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(-3, result.Entity.Difference);
        Assert.Equal("A", Assert.Single(result.Entity.Stocks).Code);
        Assert.Equal(-3, Assert.Single(result.Entity.Adjustments).Quantity);
    }

    /// <summary>
    /// 测试 - Footer marker 缺失、重复或早于 GapRows 时应返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FooterMarkerErrors_ShouldBeStructured()
    {
        var layout = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("赠品单", "A1", item => item.Items, region => region
                .Footer("合计", footer => footer.GapRows(1))));
        static byte[] CreateWorkbook(params string[] markerRows)
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("赠品单");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            for (var index = 0; index < markerRows.Length; index++)
                sheet.CreateRow(1 + index).CreateCell(0).SetCellValue(markerRows[index]);
            using var stream = new MemoryStream();
            workbook.Write(stream, false);
            return stream.ToArray();
        }

        using (var missing = new MemoryStream(CreateWorkbook("项目"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(missing, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("缺少尾部标记", exception.Message);
        }
        using (var duplicate = new MemoryStream(CreateWorkbook("合计", "合计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(duplicate, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("尾部标记重复", exception.Message);
        }
        using (var early = new MemoryStream(CreateWorkbook("合计"), writable: false))
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelImporter().ImportEntity(early, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("起始位置之前", exception.Message);
        }
    }

    /// <summary>
    /// 测试 - 动态字典为空时应自动创建，无法创建时统一返回 Plan 配置错误。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_DynamicDictionaryCreationErrors_ShouldBeStructured()
    {
        static byte[] CreateInput()
        {
            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("动态");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
            sheet.GetRow(0).CreateCell(1).SetCellValue("值");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("A");
            sheet.GetRow(1).CreateCell(1).SetCellValue("v");
            using var stream = new MemoryStream();
            workbook.Write(stream, false);
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
            var result = new NpoiExcelImporter().ImportEntity(writable, writableLayout);
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
                new NpoiExcelImporter().ImportEntity(source, layout));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Contains("动态列属性", exception.Message);
        }
        AssertCreationFailure(readOnlyLayout);
        AssertCreationFailure(noConstructorLayout);
        AssertCreationFailure(incompatibleLayout);
    }

    /// <summary>
    /// 测试 - XLS/HSSF 与 XLSX/XSSF 均应在写入前拒绝 Footer 物理行列溢出。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_FooterPhysicalBounds_ShouldBePreflightedForHssfAndXssf()
    {
        static byte[] CreateHssfTemplate(string sheetName)
        {
            using var workbook = new HSSFWorkbook();
            workbook.CreateSheet(sheetName);
            using var stream = new MemoryStream();
            workbook.Write(stream);
            return stream.ToArray();
        }

        var hssfLayout = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "A65536", item => item.Items, region => region
                .Header(false).Footer("合计")));
        using (var template = new MemoryStream(CreateHssfTemplate("边界"), writable: false))
        using (var output = new MemoryStream())
        {
            new NpoiExcelExporter().ExportForTemplate(new GiftOrder(), hssfLayout,
                new ExcelEntityTemplateOptions(template), output);
            Assert.NotEmpty(output.ToArray());
        }

        var hssfOverflow = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "A65536", item => item.Items, region => region
                .Header(false).Footer("合计", footer => footer.GapRows(1))));
        using (var template = new MemoryStream(CreateHssfTemplate("边界"), writable: false))
        using (var output = new MemoryStream())
        {
            var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
                new NpoiExcelExporter().ExportForTemplate(new GiftOrder(), hssfOverflow,
                    new ExcelEntityTemplateOptions(template), output));
            Assert.Equal(BingOfficesStage.Plan, exception.Stage);
            Assert.Empty(output.ToArray());
        }

        var xssfOverflow = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "A1048576", item => item.Items, region => region
                .Header(false).Footer("合计", footer => footer.GapRows(1))));
        using var xssfOutput = new MemoryStream();
        var xssfException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportEntity(new GiftOrder(), xssfOverflow, xssfOutput));
        Assert.Equal(BingOfficesStage.Plan, xssfException.Stage);
        Assert.Empty(xssfOutput.ToArray());

        var hssfColumnOverflow = ExcelEntity.Layout<GiftOrder>(builder => builder
            .ListRegion("边界", "IV1", item => item.Items, region => region
                .Header(false).Footer("合计", footer => footer.Cell("B1", "越界"))));
        using var columnTemplate = new MemoryStream(CreateHssfTemplate("边界"), writable: false);
        using var columnOutput = new MemoryStream();
        var columnException = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(new GiftOrder(), hssfColumnOverflow,
                new ExcelEntityTemplateOptions(columnTemplate), columnOutput));
        Assert.Equal(BingOfficesStage.Plan, columnException.Stage);
        Assert.Empty(columnOutput.ToArray());
    }

    /// <summary>
    /// 测试 - 实体导入选项应在 NPOI 打开工作簿前执行输入字节限制。
    /// </summary>
    [Fact]
    public void Npoi_EntityImportOptions_MaxInputBytes_ShouldRejectBeforeWorkbookOpen()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title));
        using var generated = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(new Invoice { Title = "资源限制" }, layout, generated);
        using var source = new MemoryStream(generated.ToArray());
        IExcelEntityResourceImporter importer = new NpoiExcelImporter();
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxInputBytes = 1 });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntity(source, layout, options));

        Assert.Equal(BingOfficesErrorCode.ResourceLimitExceeded, exception.Code);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal("NPOI", exception.Provider);
        Assert.Equal(BingOfficesStage.Open, exception.Stage);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - 布局构建后应与构建器后续修改隔离。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldSnapshotBuilderAndMapping()
    {
        var builder = new ExcelEntityLayoutBuilder<Invoice>();
        builder.Cell("单据", "A1", item => item.Number);
        var layout = builder.Build();
        builder.Cell("单据", "A2", item => item.Customer);
        Assert.Single(layout.Cells);
    }

    /// <summary>
    /// 测试 - 模板实体导出应复用已有合并区域，模板导入应校验合并结构。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldPreserveMergeAndValidateTemplateShape()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B1")
            .Cell("单据", "B1", item => item.Title));
        var templateBytes = CreateTemplate(withMerge: true);
        var source = new Invoice { Title = "模板单据" };
        using var template = new MemoryStream(templateBytes);
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal("模板单据", workbook.GetSheet("单据").GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal(1, workbook.GetSheet("单据").NumMergedRegions);
        }

        using var importSource = new MemoryStream(output.ToArray());
        using var importTemplate = new MemoryStream(templateBytes);
        var result = new NpoiExcelImporter().ImportForTemplate(importSource, layout,
            new ExcelEntityTemplateOptions(importTemplate));
        Assert.True(result.IsSuccess);
        Assert.Equal("模板单据", result.Entity.Title);

        using var missingMergeTemplate = new MemoryStream(CreateTemplate(withMerge: false));
        using var missingMergeSource = new MemoryStream(output.ToArray());
        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelImporter().ImportForTemplate(
            missingMergeSource, layout, new ExcelEntityTemplateOptions(missingMergeTemplate)));
    }

    /// <summary>
    /// 测试 - NPOI HSSF 模板应支持实体固定单元格和合并区域。
    /// </summary>
    [Fact]
    public void EntityTemplate_Hssf_ShouldRoundTrip()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B1")
            .Cell("单据", "A1", item => item.Title));
        var templateBytes = CreateHssfTemplate();
        using var template = new MemoryStream(templateBytes);
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportForTemplate(new Invoice { Title = "HSSF" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        using var read = new MemoryStream(output.ToArray());
        using var workbook = WorkbookFactory.Create(read);
        Assert.IsType<HSSFWorkbook>(workbook);
        Assert.Equal("HSSF", workbook.GetSheet("单据").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - NPOI HSSF 模板应支持实体动态列组和尾部聚合往返。
    /// </summary>
    [Fact]
    public void EntityTemplate_Hssf_ShouldRoundTripDynamicGroupsAndFooter()
    {
        var layout = ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .DynamicColumnGroup("商品", item => item.ProductFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "color", Title = "颜色", DataType = typeof(string) }
                })
                .DynamicColumnGroup("扩展", item => item.ExtensionFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "weight", Title = "重量", DataType = typeof(decimal) }
                })
                .Footer("合计", footer => footer
                    .GapRows(1)
                    .MarkerStyle(new ExcelCellStyle { Bold = true })
                    .Merge("A1:B1")
                    .Cell("E1", items => items.Sum(item => item.Amount), numberFormat: "0.00"))));
        var source = new DynamicDocument
        {
            Lines = new List<DynamicLine>
            {
                new DynamicLine
                {
                    Code = "A",
                    Amount = 1.25m,
                    ProductFields = new Dictionary<string, object> { ["color"] = "红" },
                    ExtensionFields = new Dictionary<string, object> { ["weight"] = 2.5m }
                }
            }
        };
        byte[] templateBytes;
        using (var template = new HSSFWorkbook())
        {
            template.CreateSheet("明细");
            using var buffer = new MemoryStream();
            template.Write(buffer);
            templateBytes = buffer.ToArray();
        }

        using var templateStream = new MemoryStream(templateBytes);
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(templateStream, leaveOpen: true), output);

        using var exported = new MemoryStream(output.ToArray(), writable: false);
        using (var workbook = WorkbookFactory.Create(exported))
        {
            var sheet = Assert.IsType<HSSFWorkbook>(workbook).GetSheet("明细");
            Assert.Equal("颜色", sheet.GetRow(0).GetCell(2).StringCellValue);
            Assert.Equal("红", sheet.GetRow(1).GetCell(2).StringCellValue);
            Assert.Equal("合计", sheet.GetRow(3).GetCell(0).StringCellValue);
            Assert.True(workbook.GetFontAt(sheet.GetRow(3).GetCell(0).CellStyle.FontIndex).IsBold);
            Assert.Equal(1.25d, sheet.GetRow(3).GetCell(4).NumericCellValue, 2);
            Assert.Contains(Enumerable.Range(0, sheet.NumMergedRegions)
                .Select(index => sheet.GetMergedRegion(index).FormatAsString()), value => value == "A4:B4");
        }

        using var input = new MemoryStream(output.ToArray(), writable: false);
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("红", result.Entity.Lines[0].ProductFields["color"]);
        Assert.Equal(2.5m, Convert.ToDecimal(result.Entity.Lines[0].ExtensionFields["weight"]));
    }

    /// <summary>
    /// 测试 - NPOI 实体导入只读固定属性时应返回结构化配置错误。
    /// </summary>
    [Fact]
    public void Npoi_EntityLayout_ShouldRejectReadOnlyPropertyOnImport()
    {
        var layout = ExcelEntity.Layout<ReadOnlyInvoice>(builder => builder
            .Cell("单据", "A1", item => item.Title));
        byte[] sourceBytes;
        using (var workbook = new XSSFWorkbook())
        {
            workbook.CreateSheet("单据").CreateRow(0).CreateCell(0).SetCellValue("只读");
            using var buffer = new MemoryStream();
            workbook.Write(buffer, false);
            sourceBytes = buffer.ToArray();
        }
        using var source = new MemoryStream(sourceBytes, writable: false);
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(source, layout));
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("不可写入", exception.Message);
    }

    /// <summary>
    /// 测试 - 模板实体导出应保留单元格样式、未映射公式、图片数据及图片锚点，且不修改模板原件。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldPreserveStylePictureAnchorAndOriginalTemplate()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B1")
            .Cell("单据", "B1", item => item.Title));
        var templateBytes = CreateRichTemplate();
        using var template = new MemoryStream(templateBytes.ToArray());
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportForTemplate(new Invoice { Title = "保真单据" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        Assert.Equal(templateBytes, template.ToArray());
        output.Position = 0;
        using var workbook = Assert.IsType<XSSFWorkbook>(WorkbookFactory.Create(output));
        var sheet = workbook.GetSheet("单据");
        var title = sheet.GetRow(0).GetCell(0);
        Assert.Equal("保真单据", title.StringCellValue);
        Assert.Equal(IndexedColors.LightCornflowerBlue.Index, title.CellStyle.FillForegroundColor);
        Assert.Equal(FillPattern.SolidForeground, title.CellStyle.FillPattern);
        Assert.True(workbook.GetFontAt(title.CellStyle.FontIndex).IsBold);
        Assert.Equal("1+1", sheet.GetRow(0).GetCell(2).CellFormula);
        Assert.Equal(18 * 256, sheet.GetColumnWidth(0));

        var pictureData = Assert.IsAssignableFrom<IPictureData>(Assert.Single(workbook.GetAllPictures()));
        Assert.Equal(TemplatePng, pictureData.Data);
        var drawing = Assert.IsType<XSSFDrawing>(sheet.CreateDrawingPatriarch());
        var picture = Assert.IsType<XSSFPicture>(Assert.Single(drawing.GetShapes()));
        Assert.Equal(2, picture.ClientAnchor.Row1);
        Assert.Equal(4, picture.ClientAnchor.Row2);
        Assert.Equal(2, picture.ClientAnchor.Col1);
        Assert.Equal(4, picture.ClientAnchor.Col2);
    }

    /// <summary>
    /// 测试 - 模板已有合并区域与布局部分重叠或包含但不相等时应在写出前失败。
    /// </summary>
    /// <param name="existingRange">模板中的已有合并区域。</param>
    [Theory]
    [InlineData("B2:C3")]
    [InlineData("A1:C3")]
    [InlineData("A1:C1")]
    public void EntityTemplate_ShouldRejectNonExactMergeConflictsBeforeWriting(string existingRange)
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B2")
            .Cell("单据", "B2", item => item.Title));
        var templateBytes = CreateTemplateWithMerge(existingRange);
        using var template = new MemoryStream(templateBytes.ToArray());
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportForTemplate(new Invoice { Title = "冲突" }, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output));

        Assert.Contains("合并区域与模板已有区域冲突", exception.Message);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(output.ToArray());
        Assert.Equal(templateBytes, template.ToArray());
    }

    /// <summary>
    /// 测试 - 一万行实体明细应在精确声明边界内完整导出并导入。
    /// </summary>
    [Fact]
    public void EntityListRegion_LargeDetail_ShouldRoundTripAtExactBoundary()
    {
        const int itemCount = 10000;
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region.End("B10001")));
        var source = new Invoice
        {
            Lines = Enumerable.Range(1, itemCount)
                .Select(index => new InvoiceLine { Code = "L" + index, Quantity = index })
                .ToList()
        };
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportEntity(source, layout, output);

        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(itemCount, sheet.LastRowNum);
            Assert.Equal("L1", sheet.GetRow(1).GetCell(0).StringCellValue);
            Assert.Equal("L10000", sheet.GetRow(itemCount).GetCell(0).StringCellValue);
            Assert.Equal(10000d, sheet.GetRow(itemCount).GetCell(1).NumericCellValue);
        }

        using var input = new MemoryStream(output.ToArray());
        var result = new NpoiExcelImporter().ImportEntity(input, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal(itemCount, result.Entity.Lines.Count);
        Assert.Equal("L1", result.Entity.Lines[0].Code);
        Assert.Equal("L10000", result.Entity.Lines[itemCount - 1].Code);
        Assert.Equal(itemCount, result.Entity.Lines[itemCount - 1].Quantity);

        source.Lines.Add(new InvoiceLine { Code = "越界", Quantity = itemCount + 1 });
        using var overflowOutput = new MemoryStream();
        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportEntity(source, layout, overflowOutput));
        Assert.Contains("列表区域超出声明边界", exception.Message);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(overflowOutput.ToArray());
    }

    /// <summary>
    /// 测试 - 空集合带表头时，映射列宽超出区域应在写入前失败并保持输出为空。
    /// </summary>
    [Fact]
    public void EntityListRegion_EmptyWithHeader_ShouldRejectColumnOverflowBeforeWriting()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region.End("A1")));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportEntity(new Invoice(), layout, output));

        Assert.Contains("列表区域超出声明边界", exception.Message);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 测试 - 无表头空集合应允许精确列边界，并仍拒绝映射列宽超出的声明。
    /// </summary>
    [Fact]
    public void EntityListRegion_EmptyWithoutHeader_ShouldValidateColumnBounds()
    {
        var exactLayout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines,
                region => region.Header(false).End("B1")));
        using var exactOutput = new MemoryStream();

        new NpoiExcelExporter().ExportEntity(new Invoice(), exactLayout, exactOutput);
        Assert.NotEmpty(exactOutput.ToArray());
        exactOutput.Position = 0;
        using (var workbook = WorkbookFactory.Create(exactOutput))
            Assert.Null(workbook.GetSheet("明细").GetRow(0));

        var overflowLayout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines,
                region => region.Header(false).End("A1")));
        using var overflowOutput = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportEntity(new Invoice(), overflowLayout, overflowOutput));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(overflowOutput.ToArray());
    }

    /// <summary>
    /// 测试 - 导入应在复制区域单元格前拒绝超出 End.Column 的映射宽度。
    /// </summary>
    [Fact]
    public void EntityListRegion_ImportShouldRejectColumnOverflowBeforeReading()
    {
        var sourceLayout = ExcelEntity.Layout<Invoice>(builder =>
            builder.ListRegion("明细", "A1", item => item.Lines));
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(new Invoice
        {
            Lines = new List<InvoiceLine> { new InvoiceLine { Code = "A", Quantity = 1 } }
        }, sourceLayout, exported);
        var sourceBytes = exported.ToArray();

        var restrictedLayout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region.End("A2")));
        using var input = new MemoryStream(sourceBytes);

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelImporter().ImportEntity(input, restrictedLayout));

        Assert.Contains("列表区域超出声明边界", exception.Message);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Equal(sourceBytes, input.ToArray());
    }

    /// <summary>
    /// 测试 - 模板 LeaveOpen=false 时应释放调用方交给 Provider 的模板流。
    /// </summary>
    [Fact]
    public void EntityTemplate_ShouldDisposeTemplateWhenLeaveOpenIsFalse()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B1")
            .Cell("单据", "A1", item => item.Title));
        var exporterTemplate = new MemoryStream(CreateTemplate(withMerge: true));
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportForTemplate(new Invoice { Title = "释放模板" }, layout,
            new ExcelEntityTemplateOptions(exporterTemplate, leaveOpen: false), output);

        Assert.False(exporterTemplate.CanRead);
        Assert.NotEmpty(output.ToArray());

        using var source = new MemoryStream(output.ToArray());
        var importerTemplate = new MemoryStream(CreateTemplate(withMerge: true));
        var result = new NpoiExcelImporter().ImportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(importerTemplate, leaveOpen: false));

        Assert.True(result.IsSuccess);
        Assert.False(importerTemplate.CanRead);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - 实体布局应拒绝越界地址及可确定的区域冲突。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectInvalidAddressesAndOverlappingRegions()
    {
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder =>
            builder.Cell("单据", "XFE1", item => item.Title)));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder =>
            builder.Cell("单据", "A1048577", item => item.Title)));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder => builder
            .Merge("单据", "A1:B2")
            .Merge("单据", "B2:C3")
            .Cell("单据", "A1", item => item.Title)));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title)
            .ListRegion("单据", "A1", item => item.Lines)));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("明细", "B2", item => item.Lines, region => region.End("A1"))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .Footer("合计", footer => footer.Merge("A1:B1").Cell("B1", "冲突")))));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("明细", "A1", item => item.Lines, region => region
                .DynamicColumnGroup("goods", item => item.ProductFields, new[]
                {
                    new ExcelDynamicColumnDefinition { Key = "goods", Title = "商品" }
                }))));
    }

    /// <summary>
    /// 测试 - 可变行数和尾部导致的列表区域运行时重叠应在写入前返回配置错误。
    /// </summary>
    [Fact]
    public void EntityLayout_ShouldRejectRuntimeOverlapBetweenListRegions()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("单据", "A1", item => item.Lines,
                region => region.Footer("合计"))
            .ListRegion("单据", "A3", item => item.Lines,
                region => region.End("B4")));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            new NpoiExcelExporter().ExportEntity(new Invoice
            {
                Lines = new List<InvoiceLine> { new() { Code = "A", Quantity = 1 } }
            }, layout, output));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("运行时范围", exception.Message, StringComparison.Ordinal);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 测试 - Footer 聚合委托异常时不应提交部分工作簿。
    /// </summary>
    [Fact]
    public void EntityLayout_FooterValueFactoryFailure_ShouldNotWritePartialWorkbook()
    {
        var expected = new InvalidOperationException("合计失败");
        Func<IReadOnlyList<InvoiceLine>, string> failure = _ => throw expected;
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .ListRegion("单据", "A1", item => item.Lines,
                region => region.Footer("合计", footer => footer.Cell("B1", failure))));
        using var output = new MemoryStream();

        var exception = Assert.Throws<BingOfficesExportException>(() =>
            new NpoiExcelExporter().ExportEntity(new Invoice
            {
                Lines = new List<InvoiceLine> { new() { Code = "A", Quantity = 1 } }
            }, layout, output));

        Assert.Same(expected, exception.InnerException);
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// 测试 - 固定单元格应按布局声明的名称选择值转换器。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseNamedConverter()
    {
        var converter = new EntityTitleConverter();
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title, converter.Name));
        using var exported = new MemoryStream();
        new NpoiExcelExporter(valueConverters: new[] { converter }).ExportEntity(
            new Invoice { Title = "销售单" }, layout, exported);

        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
            Assert.Equal("entity:销售单", workbook.GetSheet("单据").GetRow(0).GetCell(0).StringCellValue);

        using var source = new MemoryStream(exported.ToArray());
        var result = new NpoiExcelImporter(valueConverters: new[] { converter }).ImportEntity(source, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal("销售单", result.Entity.Title);
    }

    /// <summary>
    /// 测试 - 固定单元格应复用映射计划的双向值映射。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldUseMappingPlanValueMapInBothDirections()
    {
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = nameof(Invoice.Title),
            ValueMappings = new List<ExcelValueMappingConfiguration>
            {
                new ExcelValueMappingConfiguration { Text = "显示标题", Value = "内部标题" }
            }
        });
        var layout = ExcelEntity.Layout<Invoice>(builder => builder.Cell("单据", "A1",
            item => item.Title, cell => cell.Mapping(mapping)));

        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(new Invoice { Title = "内部标题" }, layout, exported);
        exported.Position = 0;
        using (var workbook = WorkbookFactory.Create(exported))
            Assert.Equal("显示标题", workbook.GetSheet("单据").GetRow(0).GetCell(0).StringCellValue);

        using var source = new MemoryStream(exported.ToArray());
        var result = new NpoiExcelImporter().ImportEntity(source, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal("内部标题", result.Entity.Title);
    }

    /// <summary>
    /// 测试 - 固定单元格应执行属性特性校验并返回完整错误坐标。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldReturnStructuredValidationError()
    {
        var layout = ExcelEntity.Layout<RequiredInvoice>(builder => builder
            .Cell("单据", "C4", item => item.Title));
        byte[] inputBytes;
        using (var workbook = new XSSFWorkbook())
        {
            using var output = new MemoryStream();
            workbook.CreateSheet("单据").CreateRow(3).CreateCell(2).SetCellValue(string.Empty);
            workbook.Write(output, false);
            inputBytes = output.ToArray();
        }
        using var input = new MemoryStream(inputBytes);

        var result = new NpoiExcelImporter().ImportEntity(input, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("单据", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RequiredInvoice.Title), error.PropertyName);
    }

    /// <summary>
    /// 测试 - 固定单元格应在不可空数值转换前执行必填校验，并跳过转换与 setter。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldValidateRequiredBeforeNumericConversion()
    {
        var layout = ExcelEntity.Layout<RequiredNumberInvoice>(builder => builder
            .Cell("单据", "C4", item => item.Number));
        byte[] inputBytes;
        using (var workbook = new XSSFWorkbook())
        {
            using var output = new MemoryStream();
            workbook.CreateSheet("单据").CreateRow(3).CreateCell(2).SetCellValue(string.Empty);
            workbook.Write(output, false);
            inputBytes = output.ToArray();
        }
        using var input = new MemoryStream(inputBytes);

        var result = new NpoiExcelImporter().ImportEntity(input, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("单据", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(RequiredNumberInvoice.Number), error.PropertyName);
        Assert.Equal(0, result.Entity.Number);
    }

    /// <summary>
    /// 测试 - 请求级命名校验应在固定单元格转换后执行，并在失败时阻止 setter。
    /// </summary>
    [Fact]
    public void EntityCell_ShouldRunRequestValidationAfterConversion()
    {
        var rule = new EntityConvertedNumberValidationRule();
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = nameof(Invoice.Number),
            ValidationRuleNames = new List<string> { rule.Name }
        });
        var layout = ExcelEntity.Layout<Invoice>(builder => builder.Cell("单据", "C4",
            item => item.Number, cell => cell.Mapping(mapping)));
        byte[] inputBytes;
        using (var workbook = new XSSFWorkbook())
        {
            using var output = new MemoryStream();
            workbook.CreateSheet("单据").CreateRow(3).CreateCell(2).SetCellValue("42");
            workbook.Write(output, false);
            inputBytes = output.ToArray();
        }
        using var input = new MemoryStream(inputBytes);

        var result = new NpoiExcelImporter(namedValidationRules: new[] { rule }).ImportEntity(input, layout);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("单据", error.SheetName);
        Assert.Equal(4, error.RowIndex);
        Assert.Equal(3, error.ColumnIndex);
        Assert.Equal(nameof(Invoice.Number), error.PropertyName);
        Assert.Equal("42", rule.LastRawValue);
        Assert.Equal(42, rule.LastConvertedValue);
        Assert.Equal(0, result.Entity.Number);
    }

    /// <summary>
    /// 测试 - 实体列表导入应覆盖内置校验矩阵并保留完整坐标。
    /// </summary>
    [Fact]
    public void EntityLayout_ValidationMatrix_ShouldReportCompleteCoordinates()
    {
        var layout = ExcelEntity.Layout<EntityValidationRoot>(builder =>
            builder.ListRegion("验证", "A1", root => root.Lines));

        void AssertFailure(EntityValidationRoot source, string header, int rowIndex,
            Action<ICell> mutate, ExcelImportErrorCode code, string propertyName)
        {
            using var exported = new MemoryStream();
            new NpoiExcelExporter().ExportEntity(source, layout, exported);
            exported.Position = 0;
            byte[] inputBytes;
            using (var workbook = WorkbookFactory.Create(exported))
            {
                var sheet = workbook.GetSheet("验证");
                var columnIndex = Enumerable.Range(0, sheet.GetRow(0).LastCellNum)
                    .Single(index => sheet.GetRow(0).GetCell(index).StringCellValue == header);
                mutate(sheet.GetRow(rowIndex - 1).GetCell(columnIndex));
                using var input = new MemoryStream();
                workbook.Write(input, false);
                inputBytes = input.ToArray();
            }

            using var sourceStream = new MemoryStream(inputBytes, writable: false);
            var result = new NpoiExcelImporter().ImportEntity(sourceStream, layout);
            Assert.False(result.IsSuccess);
            var error = result.Errors.FirstOrDefault(item => item.Code == code
                && item.RowIndex == rowIndex && item.PropertyName == propertyName);
            Assert.True(error != null, string.Join(";", result.Errors.Select(item =>
                $"{item.Code}:{item.PropertyName}:{item.RowIndex}:{item.ColumnIndex}:{item.Message}")));
            Assert.Equal(code, error.Code);
            Assert.Equal("验证", error.SheetName);
            Assert.Equal(rowIndex, error.RowIndex);
            Assert.Equal(propertyName, error.PropertyName);
            Assert.True(error.ColumnIndex > 0);
        }

        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Required), 2,
            cell => cell.SetCellValue(string.Empty), ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Required));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Name), 2,
            cell => cell.SetCellValue("超长名称"), ExcelImportErrorCode.MaxLength,
            nameof(EntityValidationLine.Name));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Quantity), 2,
            cell => cell.SetCellValue(99d), ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Quantity));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Code), 2,
            cell => cell.SetCellValue("BAD"), ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.Code));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.DateValue), 2,
            cell => cell.SetCellValue("2024/01/01"), ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.DateValue));
        AssertFailure(CreateValidationRoot(new EntityValidationLine { UniqueCode = "U1" },
                new EntityValidationLine { UniqueCode = "U2" }), nameof(EntityValidationLine.UniqueCode), 3,
            cell => cell.SetCellValue("U1"), ExcelImportErrorCode.Validation,
            nameof(EntityValidationLine.UniqueCode));
        AssertFailure(CreateValidationRoot(), nameof(EntityValidationLine.Number), 2,
            cell => cell.SetCellValue("不是数字"), ExcelImportErrorCode.ValueConversion,
            nameof(EntityValidationLine.Number));
    }

    /// <summary>
    /// 测试 - 实体异步入口应完成真实流复制，并在预取消时不写出结果。
    /// </summary>
    [Fact]
    public async Task EntityAsync_ShouldRoundTripAndHonorCancellation()
    {
        var layout = ExcelEntity.Layout<Invoice>(builder => builder.Cell("单据", "A1", item => item.Title));
        using var output = new MemoryStream();
        await new NpoiExcelExporter().ExportEntityAsync(new Invoice { Title = "异步" }, layout, output);
        Assert.NotEmpty(output.ToArray());

        using var input = new MemoryStream(output.ToArray());
        var result = await new NpoiExcelImporter().ImportEntityAsync(input, layout);
        Assert.True(result.IsSuccess);
        Assert.Equal("异步", result.Entity.Title);

        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        using var canceledOutput = new MemoryStream();
        await Assert.ThrowsAsync<OperationCanceledException>(() => new NpoiExcelExporter().ExportEntityAsync(
            new Invoice { Title = "取消" }, layout, canceledOutput, canceled.Token));
        Assert.Empty(canceledOutput.ToArray());
    }

    /// <summary>
    /// 测试 - 实体文件导出预取消时应保留既有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task EntityFileExport_PreCanceled_ShouldPreserveTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-offices-entity-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 1, 2, 3, 4 };
        File.WriteAllBytes(path, original);
        var layout = ExcelEntity.Layout<Invoice>(builder => builder.Cell("单据", "A1", item => item.Title));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new NpoiExcelExporter().ExportEntityToFileAsync(new Invoice { Title = "取消" }, layout,
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
    /// 测试 - 实体文件导出写入中取消时应保留既有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task EntityFileExport_MidFlightCancellation_ShouldPreserveTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-offices-entity-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "entity.xlsx");
        var original = new byte[] { 9, 8, 7, 6 };
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstEntityConverter(cancellation);
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title, converter.Name));

        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                new NpoiExcelExporter(valueConverters: new[] { converter }).ExportEntityToFileAsync(
                    new Invoice { Title = "中途取消" }, layout, path, cancellation.Token));
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
    /// 创建实体模板测试使用的 XSSF 模板。
    /// </summary>
    /// <param name="withMerge">是否创建标题合并区域。</param>
    /// <returns>模板工作簿字节。</returns>
    private static byte[] CreateTemplate(bool withMerge)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("单据");
        var row = sheet.CreateRow(0);
        row.CreateCell(0).SetCellValue("模板标题");
        if (withMerge)
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    /// <summary>
    /// 创建实体模板测试使用的 HSSF 模板。
    /// </summary>
    /// <returns>模板工作簿字节。</returns>
    private static byte[] CreateHssfTemplate()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("单据");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("模板标题");
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        using var stream = new MemoryStream();
        workbook.Write(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// 创建包含样式、公式、合并区域和图片的 XSSF 模板。
    /// </summary>
    /// <returns>模板工作簿字节。</returns>
    private static byte[] CreateRichTemplate()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("单据");
        var row = sheet.CreateRow(0);
        var title = row.CreateCell(0);
        title.SetCellValue("模板标题");
        var font = workbook.CreateFont();
        font.IsBold = true;
        var style = workbook.CreateCellStyle();
        style.SetFont(font);
        style.FillForegroundColor = IndexedColors.LightCornflowerBlue.Index;
        style.FillPattern = FillPattern.SolidForeground;
        title.CellStyle = style;
        row.CreateCell(2).CellFormula = "1+1";
        sheet.SetColumnWidth(0, 18 * 256);
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
        var anchor = workbook.GetCreationHelper().CreateClientAnchor();
        anchor.Row1 = 2;
        anchor.Row2 = 4;
        anchor.Col1 = 2;
        anchor.Col2 = 4;
        var pictureIndex = workbook.AddPicture(TemplatePng, PictureType.PNG);
        sheet.CreateDrawingPatriarch().CreatePicture(anchor, pictureIndex);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    /// <summary>
    /// 创建包含指定合并区域的 XSSF 模板。
    /// </summary>
    /// <param name="range">A1 格式的合并区域。</param>
    /// <returns>模板工作簿字节。</returns>
    private static byte[] CreateTemplateWithMerge(string range)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("单据");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("模板标题");
        sheet.AddMergedRegion(CellRangeAddress.ValueOf(range));
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    /// <summary>
    /// 获取模板保真测试使用的单像素 PNG 数据。
    /// </summary>
    private static readonly byte[] TemplatePng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    /// <summary>
    /// 测试实体。
    /// </summary>
    public sealed class Invoice
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
        public List<InvoiceLine> Lines { get; set; } = new List<InvoiceLine>();
    }

    /// <summary>
    /// 只读固定属性导入测试模型。
    /// </summary>
    private sealed class ReadOnlyInvoice
    {
        /// <summary>
        /// 获取单据标题。
        /// </summary>
        public string Title => "只读";
    }

    /// <summary>
    /// 属性式单元格和动态分组测试的根模型。
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
    /// 包含多个动态字段字典的采购明细。
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
    /// 固定单元格属性校验测试实体。
    /// </summary>
    public sealed class RequiredInvoice
    {
        /// <summary>
        /// 获取或设置必填标题。
        /// </summary>
        [ExcelRequired]
        public string Title { get; set; }
    }

    /// <summary>
    /// 固定单元格不可空数值校验测试实体。
    /// </summary>
    public sealed class RequiredNumberInvoice
    {
        /// <summary>
        /// 获取或设置必填编号。
        /// </summary>
        [ExcelRequired]
        public int Number { get; set; }
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
    /// 测试明细实体。
    /// </summary>
    public sealed class InvoiceLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置明细数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 既有单字典动态列测试使用的聚合实体。
    /// </summary>
    public sealed class LegacyDynamicDocument
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<LegacyDynamicLine> Lines { get; set; } = new List<LegacyDynamicLine>();
    }

    /// <summary>
    /// 既有单字典动态列测试使用的明细实体。
    /// </summary>
    public sealed class LegacyDynamicLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段字典。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 多动态列组测试使用的聚合实体。
    /// </summary>
    public sealed class DynamicDocument
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<DynamicLine> Lines { get; set; } = new List<DynamicLine>();
    }

    /// <summary>
    /// 显式动态列校验测试使用的聚合实体。
    /// </summary>
    private sealed class ExplicitDynamicDocument
    {
        /// <summary>
        /// 获取或设置明细行集合。
        /// </summary>
        public List<ExplicitDynamicLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 显式动态列校验测试使用的明细实体。
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
    /// 多动态列组测试使用的明细实体。
    /// </summary>
    public sealed class DynamicLine
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置商品动态字段。
        /// </summary>
        public IDictionary<string, object> ProductFields { get; set; } =
            new Dictionary<string, object>();

        /// <summary>
        /// 获取或设置扩展动态字段。
        /// </summary>
        public IDictionary<string, object> ExtensionFields { get; set; } =
            new Dictionary<string, object>();
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
    /// 固定单元格测试使用的命名转换器。
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
    /// 记录转换后值并故意失败的请求级命名校验规则。
    /// </summary>
    private sealed class EntityConvertedNumberValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "entity-converted-number";

        /// <inheritdoc />
        public string ErrorMessage => "请求级数值校验失败";

        /// <summary>
        /// 获取最近一次收到的原始文本。
        /// </summary>
        public string LastRawValue { get; private set; }

        /// <summary>
        /// 获取最近一次收到的转换后值。
        /// </summary>
        public object LastConvertedValue { get; private set; }

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context)
        {
            LastRawValue = context.Value;
            LastConvertedValue = context.ConvertedValue;
            return false;
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
        /// 初始化一个 <see cref="CancelAfterFirstEntityConverter"/> 类型的实例。
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
}
