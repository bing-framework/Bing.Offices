using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 实体导入原生校验测试。
/// </summary>
public sealed class EntityNativeValidationTest
{
    /// <summary>
    /// 验证 XLS 与 XLSX 实体模板导出共用字典和名称下拉。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityExport_LookupList_ShouldRoundTripTemplate(bool xls)
    {
        var range = new ExcelRangeDefinition
        {
            StartRow = 0, EndRow = 0, StartColumn = 1, EndColumn = 1
        };
        var layout = ExcelEntity.Layout<Order>(builder => builder
            .Cell("Order", "B1", item => item.Status)
            .ListRegion("Order", "A3", item => item.Lines)
            .LookupList("Allowed", new[] { "Open", "Good" })
            .DataValidation("Order", new ExcelDataValidationDefinition
            {
                Range = range, Type = ExcelDataValidationType.List, ListSourceName = "Allowed"
            }));
        using var template = new MemoryStream();
        using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
        {
            workbook.CreateSheet("Order");
            workbook.Write(template, true);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(new Order
        {
            Status = "Open", Lines = new List<Line> { new() { Code = "Good" } }
        }, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal(2, workbook.NumberOfSheets);
            Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(1));
            Assert.Equal("Open", workbook.GetSheetAt(1).GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal("Good", workbook.GetSheetAt(1).GetRow(1).GetCell(0).StringCellValue);
            Assert.Single(workbook.GetSheet("Order").GetDataValidations());
        }
        using var input = new MemoryStream(output.ToArray());
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal("Open", result.Entity.Status);
        Assert.Equal("Good", Assert.Single(result.Entity.Lines).Code);
    }

    /// <summary>
    /// 验证固定字段和明细按原始坐标执行名称列表校验。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityImport_NamedList_ShouldValidateCellsAndKeepDefaultMode(bool xls)
    {
        var layout = ExcelEntity.Layout<Order>(builder => builder
            .Cell("Order", "B1", item => item.Status)
            .ListRegion("Order", "A3", item => item.Lines, region => region.End("A5")));
        byte[] bytes;
        using (IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook())
        {
            var sheet = workbook.CreateSheet("Order");
            var lookup = workbook.CreateSheet("Lookup");
            lookup.CreateRow(0).CreateCell(0).SetCellValue("Open");
            lookup.CreateRow(1).CreateCell(0).SetCellValue("Good");
            var name = workbook.CreateName();
            name.NameName = "Allowed";
            name.RefersToFormula = "Lookup!$A$1:$A$2";
            sheet.CreateRow(0).CreateCell(1).SetCellValue("Bad");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Code");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("Wrong");
            sheet.CreateRow(4).CreateCell(0).SetCellValue("Good");
            var helper = sheet.GetDataValidationHelper();
            var constraint = helper.CreateFormulaListConstraint("Allowed");
            sheet.AddValidationData(helper.CreateValidation(constraint,
                new CellRangeAddressList(0, 0, 1, 1)));
            sheet.AddValidationData(helper.CreateValidation(constraint,
                new CellRangeAddressList(3, 4, 0, 0)));
            using var stream = new MemoryStream();
            workbook.Write(stream, true);
            bytes = stream.ToArray();
        }

        using var legacyInput = new MemoryStream(bytes);
        var legacy = new NpoiExcelImporter().ImportEntity(legacyInput, layout);
        Assert.True(legacy.IsSuccess);
        Assert.Equal("Bad", legacy.Entity.Status);
        Assert.Equal(new[] { "Wrong", "Good" }, legacy.Entity.Lines.Select(item => item.Code));

        using var nativeInput = new MemoryStream(bytes);
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new NpoiExcelImporter().ImportEntity(nativeInput, layout, options);
        Assert.Null(result.Entity.Status);
        Assert.Equal("Good", Assert.Single(result.Entity.Lines).Code);
        Assert.Equal(new[] { (1, 2), (4, 1) }, result.Errors.Select(error =>
            (error.RowIndex, error.ColumnIndex)).OrderBy(item => item.RowIndex).ToArray());
        Assert.All(result.Errors, error => Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code));
    }

    /// <summary>
    /// 验证分页小计后的明细仍使用原始工作表坐标校验。
    /// </summary>
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityImport_PageSubtotalNamedList_ShouldReportOriginalRow(bool xls)
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
        using IWorkbook templateBook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        templateBook.CreateSheet("Order");
        using var template = new MemoryStream();
        templateBook.Write(template, leaveOpen: true);
        template.Position = 0;
        using var generated = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), generated);
        generated.Position = 0;
        using var workbook = WorkbookFactory.Create(generated);
        var sheet = workbook.GetSheet("Order");
        Assert.Equal("页小计", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal("Rejected", sheet.GetRow(4).GetCell(0).StringCellValue);
        var helper = sheet.GetDataValidationHelper();
        sheet.AddValidationData(helper.CreateValidation(helper.CreateFormulaListConstraint("Allowed"),
            new CellRangeAddressList(4, 4, 0, 0)));
        using var input = new MemoryStream();
        workbook.Write(input, leaveOpen: true);
        input.Position = 0;

        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);
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
    /// <param name="xls">是否使用 HSSF XLS 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void PurchaseHeader_LookupLists_ShouldRoundTripValueMappings(bool xls)
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
        using IWorkbook templateBook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        templateBook.CreateSheet("Order");
        using var template = new MemoryStream();
        templateBook.Write(template, leaveOpen: true);
        template.Position = 0;
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(new PurchaseHeader
        {
            Category = "RAW", Warehouse = "MAIN", Status = "PENDING"
        }, layout, new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        output.Position = 0;
        using (var workbook = WorkbookFactory.Create(output))
        {
            Assert.Equal(4, workbook.NumberOfSheets);
            Assert.Equal(new[] { "原料", "主仓", "待审核" },
                Enumerable.Range(0, 3).Select(index =>
                    workbook.GetSheet("Order").GetRow(0).GetCell(index).StringCellValue));
            Assert.Equal(new[] { "Categories", "Warehouses", "Statuses" },
                workbook.GetSheet("Order").GetDataValidations()
                    .Select(rule => rule.ValidationConstraint.Formula1));
            Assert.Equal(3, workbook.GetAllNames().Count);
            for (var index = 1; index <= 3; index++)
                Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(index));
        }
        using var input = new MemoryStream(output.ToArray());
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);
        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);
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
