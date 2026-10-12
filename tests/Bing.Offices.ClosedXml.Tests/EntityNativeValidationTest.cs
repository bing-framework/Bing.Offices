using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML 实体导入原生校验测试。
/// </summary>
public sealed class EntityNativeValidationTest
{
    /// <summary>
    /// 验证 Entity 模板生成字典并执行名称下拉校验。
    /// </summary>
    [Fact]
    public void EntityExport_LookupList_ShouldRoundTripTemplate()
    {
        var layout = ExcelEntity.Layout<Order>(builder => builder
            .Cell("Order", "B1", item => item.Status)
            .ListRegion("Order", "A3", item => item.Lines)
            .LookupList("Allowed", new[] { "Open", "Good" })
            .DataValidation("Order", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 0, EndRow = 0, StartColumn = 1, EndColumn = 1
                },
                Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
            }));
        using var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            workbook.Worksheets.Add("Order").Cell("C1").Value = "Kept";
            workbook.SaveAs(template);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportForTemplate(new Order
        {
            Status = "Open", Lines = new List<Line> { new() { Code = "Good" } }
        }, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            Assert.Equal(2, workbook.Worksheets.Count);
            Assert.Equal("Kept", workbook.Worksheet("Order").Cell("C1").GetString());
            var lookup = workbook.Worksheet(2);
            Assert.Equal(XLWorksheetVisibility.Hidden, lookup.Visibility);
            Assert.Equal(new[] { "Open", "Good" }, new[]
            {
                lookup.Cell("A1").GetString(), lookup.Cell("A2").GetString()
            });
            Assert.Equal("Allowed", Assert.Single(workbook.DefinedNames).Name);
            Assert.Single(workbook.Worksheet("Order").DataValidations);
        }
        using var input = new MemoryStream(output.ToArray());
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Open", result.Entity.Status);
        Assert.Equal("Good", Assert.Single(result.Entity.Lines).Code);
    }

    /// <summary>
    /// 验证固定字段和明细按原始坐标执行名称列表校验。
    /// </summary>
    [Fact]
    public void EntityImport_NamedList_ShouldValidateCellsAndKeepDefaultMode()
    {
        var layout = ExcelEntity.Layout<Order>(builder => builder
            .Cell("Order", "B1", item => item.Status)
            .ListRegion("Order", "A3", item => item.Lines, region => region.End("A5")));
        byte[] bytes;
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Order");
            var lookup = workbook.Worksheets.Add("Lookup");
            lookup.Cell("A1").Value = "Open";
            lookup.Cell("A2").Value = "Good";
            workbook.DefinedNames.Add("Allowed", lookup.Range("A1:A2"));
            sheet.Cell("B1").Value = "Bad";
            sheet.Cell("A3").Value = "Code";
            sheet.Cell("A4").Value = "Wrong";
            sheet.Cell("A5").Value = "Good";
            var fixedValidation = sheet.Range("B1").CreateDataValidation();
            fixedValidation.AllowedValues = XLAllowedValues.List;
            fixedValidation.Value = "Allowed";
            var listValidation = sheet.Range("A4:A5").CreateDataValidation();
            listValidation.AllowedValues = XLAllowedValues.List;
            listValidation.Value = "Allowed";
            using var stream = new MemoryStream();
            workbook.SaveAs(stream);
            bytes = stream.ToArray();
        }

        using var legacyInput = new MemoryStream(bytes);
        var legacy = new ClosedXmlExcelImporter().ImportEntity(legacyInput, layout);
        Assert.True(legacy.IsSuccess);
        Assert.Equal("Bad", legacy.Entity.Status);
        Assert.Equal(new[] { "Wrong", "Good" }, legacy.Entity.Lines.Select(item => item.Code));

        using var nativeInput = new MemoryStream(bytes);
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(nativeInput, layout, options);
        Assert.Null(result.Entity.Status);
        Assert.Equal("Good", Assert.Single(result.Entity.Lines).Code);
        Assert.Equal(new[] { (1, 2), (4, 1) }, result.Errors.Select(error =>
            (error.RowIndex, error.ColumnIndex)).OrderBy(item => item.RowIndex).ToArray());
        Assert.All(result.Errors, error => Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code));
    }

    /// <summary>
    /// 验证明细行先完成工作簿规则校验，再执行类型转换。
    /// </summary>
    [Fact]
    public void EntityImport_InvalidNativeValue_ShouldSkipLaterConversionInSameRow()
    {
        var layout = ExcelEntity.Layout<NumericOrder>(builder => builder
            .ListRegion("Order", "A1", item => item.Lines, region => region.End("B2")));
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Order");
        var lookup = workbook.Worksheets.Add("Lookup");
        lookup.Cell("A1").Value = "Allowed";
        workbook.DefinedNames.Add("AllowedValues", lookup.Range("A1"));
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Amount";
        sheet.Cell("A2").Value = "Rejected";
        sheet.Cell("B2").Value = "not-a-number";
        var validation = sheet.Range("A2").CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.List;
        validation.Value = "AllowedValues";
        using var source = new MemoryStream();
        workbook.SaveAs(source);
        source.Position = 0;

        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(source, layout, options);

        Assert.Empty(result.Entity.Lines);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Order", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(NumericLine.Code), error.PropertyName);
    }

    /// <summary>
    /// 验证分页小计后的明细仍使用原始工作表坐标校验。
    /// </summary>
    [Fact]
    public void EntityImport_PageSubtotalNamedList_ShouldReportOriginalRow()
    {
        var layout = ExcelEntity.Layout<Order>(builder => builder
            .ListRegion("Order", "A1", item => item.Lines, region => region
                .PageBreak(2).PageSubtotal("页小计").Footer("总计"))
            .LookupList("Allowed", new[] { "A", "B" }));
        var source = new Order
        {
            Lines = new List<Line>
            {
                new() { Code = "A" }, new() { Code = "B" }, new() { Code = "Rejected" }
            }
        };
        using var generated = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(source, layout, generated);
        generated.Position = 0;
        using var workbook = new XLWorkbook(generated);
        var sheet = workbook.Worksheet("Order");
        Assert.Equal("页小计", sheet.Cell("A4").GetString());
        Assert.Equal("Rejected", sheet.Cell("A5").GetString());
        var validation = sheet.Range("A5").CreateDataValidation();
        validation.AllowedValues = XLAllowedValues.List;
        validation.Value = "Allowed";
        using var input = new MemoryStream();
        workbook.SaveAs(input);
        input.Position = 0;

        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);
        Assert.Equal(new[] { "A", "B" }, result.Entity.Lines.Select(line => line.Code));
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Order", error.SheetName);
        Assert.Equal(5, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
    }

    /// <summary>
    /// 验证采购字段使用三个文本字典，并将显示值映射回业务编码。
    /// </summary>
    [Fact]
    public void PurchaseHeader_LookupLists_ShouldRoundTripValueMappings()
    {
        var layout = ExcelEntity.Layout<PurchaseHeader>(builder => builder
            .Cell("Order", "A1", item => item.Category,
                cell => cell.Mapping(CreateValueMap(nameof(PurchaseHeader.Category), "原料", "RAW")))
            .Cell("Order", "B1", item => item.Warehouse,
                cell => cell.Mapping(CreateValueMap(nameof(PurchaseHeader.Warehouse), "主仓", "MAIN")))
            .Cell("Order", "C1", item => item.Status,
                cell => cell.Mapping(CreateValueMap(nameof(PurchaseHeader.Status), "待审核", "PENDING")))
            .LookupList("Categories", new[] { "原料", "成品" })
            .LookupList("Warehouses", new[] { "主仓", "备仓" })
            .LookupList("Statuses", new[] { "待审核", "已审核" })
            .DataValidation("Order", NamedCellValidation("Categories", 0))
            .DataValidation("Order", NamedCellValidation("Warehouses", 1))
            .DataValidation("Order", NamedCellValidation("Statuses", 2)));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(new PurchaseHeader
        {
            Category = "RAW", Warehouse = "MAIN", Status = "PENDING"
        }, layout, output);
        output.Position = 0;
        using (var workbook = new XLWorkbook(output))
        {
            Assert.Equal(4, workbook.Worksheets.Count);
            Assert.Equal(new[] { "原料", "主仓", "待审核" },
                Enumerable.Range(1, 3).Select(index => workbook.Worksheet("Order").Cell(1, index).GetString()));
            Assert.Equal(new[] { "Categories", "Warehouses", "Statuses" },
                workbook.Worksheet("Order").DataValidations.Select(rule => rule.Value.TrimStart('=')));
            for (var index = 2; index <= 4; index++)
                Assert.Equal(XLWorksheetVisibility.Hidden, workbook.Worksheet(index).Visibility);
        }
        using var input = new MemoryStream(output.ToArray());
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("RAW", result.Entity.Category);
        Assert.Equal("MAIN", result.Entity.Warehouse);
        Assert.Equal("PENDING", result.Entity.Status);
    }

    /// <summary>
    /// 创建固定单元格的双向编码映射。
    /// </summary>
    /// <param name="property">目标属性名称。</param>
    /// <param name="text">工作簿显示文本。</param>
    /// <param name="value">业务编码。</param>
    /// <returns>固定单元格映射配置。</returns>
    private static ExcelMappingConfiguration CreateValueMap(string property, string text, string value)
    {
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = property,
            ValueMappings = new List<ExcelValueMappingConfiguration>
            {
                new() { Text = text, Value = value }
            }
        });
        return mapping;
    }

    /// <summary>
    /// 创建引用字典名称的固定单元格校验。
    /// </summary>
    /// <param name="name">字典名称。</param>
    /// <param name="column">零基列号。</param>
    /// <returns>名称列表校验定义。</returns>
    private static ExcelDataValidationDefinition NamedCellValidation(string name, int column) => new()
    {
        Range = new ExcelRangeDefinition
        {
            StartRow = 0, EndRow = 0, StartColumn = column, EndColumn = column
        },
        Type = ExcelDataValidationType.List,
        ListSourceName = name
    };

    /// <summary>
    /// 接收固定字段和明细的测试实体。
    /// </summary>
    public sealed class Order
    {
        /// <summary>
        /// 获取或设置采购状态。
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 获取或设置明细。
        /// </summary>
        public List<Line> Lines { get; set; } = new();
    }

    /// <summary>
    /// 接收一行明细。
    /// </summary>
    public sealed class Line
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 接收数值明细的测试实体。
    /// </summary>
    public sealed class NumericOrder
    {
        /// <summary>
        /// 获取或设置明细。
        /// </summary>
        public List<NumericLine> Lines { get; set; } = new();
    }

    /// <summary>
    /// 接收包含金额的明细。
    /// </summary>
    public sealed class NumericLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 接收采购单头部编码。
    /// </summary>
    public sealed class PurchaseHeader
    {
        /// <summary>
        /// 获取或设置类别编码。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置仓库编码。
        /// </summary>
        public string Warehouse { get; set; }

        /// <summary>
        /// 获取或设置状态编码。
        /// </summary>
        public string Status { get; set; }
    }
}
