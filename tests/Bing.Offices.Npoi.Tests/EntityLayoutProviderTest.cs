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
        /// 初始化会在首次转换后取消的转换器。
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
