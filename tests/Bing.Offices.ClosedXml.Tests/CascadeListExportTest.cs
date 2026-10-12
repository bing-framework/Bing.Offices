using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Entities;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML 级联字典导出的工作簿结构测试。
/// </summary>
public sealed class CascadeListExportTest
{
    /// <summary>
    /// 验证级联字典与随行下拉公式的完整结构。
    /// </summary>
    [Fact]
    public void Export_CascadeList_ShouldWriteHiddenMapAndDependentValidation()
    {
        var request = ExcelExport.Workbook(book => book
            .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
            {
                ["产成品"] = new[] { "手机", "电视" },
                ["原材料"] = new[] { "=A1", "金属,件", "中文\n换行" }
            })
            .AddSheet("Products", new[] { new ProductRow { Category = "原材料", Product = "=A1" } },
                sheet => sheet
                    .DataValidation(new ExcelDataValidationDefinition
                    {
                        Range = new ExcelRangeDefinition
                        {
                            StartRow = 1, EndRow = 10, StartColumn = 0, EndColumn = 0
                        },
                        Type = ExcelDataValidationType.List,
                        ListSourceName = "Categories"
                    })
                    .DataValidation(new ExcelDataValidationDefinition
                    {
                        Range = new ExcelRangeDefinition
                        {
                            StartRow = 1, EndRow = 10, StartColumn = 1, EndColumn = 1
                        },
                        Type = ExcelDataValidationType.List,
                        CascadeSourceName = "Categories",
                        ParentColumnIndex = 0
                    })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;

        using var workbook = new XLWorkbook(output);
        Assert.Equal(new[] { "Products", "__BingCascade_001" },
            workbook.Worksheets.Select(sheet => sheet.Name));
        var helper = workbook.Worksheet("__BingCascade_001");
        Assert.Equal(XLWorksheetVisibility.Hidden, helper.Visibility);
        Assert.Equal(new[] { "产成品", "原材料" },
            new[] { helper.Cell("A1").GetString(), helper.Cell("A2").GetString() });
        Assert.Equal(new[] { "手机", "电视", "=A1", "金属,件", "中文\n换行" },
            Enumerable.Range(1, 5).Select(row => helper.Cell(row, 3).GetString()));
        Assert.False(helper.Cell("C3").HasFormula);
        var names = workbook.DefinedNames.ValidNamedRanges().ToDictionary(item => item.Name);
        Assert.Contains("Categories", names.Keys);
        var mapName = helper.Cell("B1").GetString();
        Assert.Equal("__BingCascadeValues_002", mapName);
        Assert.Equal("__BingCascadeValues_003", helper.Cell("B2").GetString());
        Assert.Equal("__BingCascadeMap_001", Assert.Single(names.Keys.Where(name =>
            name.StartsWith("__BingCascadeMap_", StringComparison.Ordinal))));
        Assert.Equal(4, names.Count);
        var mapRange = Assert.Single(names["__BingCascadeMap_001"].Ranges).RangeAddress;
        Assert.Equal((1, 1, 2, 2), (mapRange.FirstAddress.RowNumber,
            mapRange.FirstAddress.ColumnNumber, mapRange.LastAddress.RowNumber,
            mapRange.LastAddress.ColumnNumber));
        var childRange = Assert.Single(names["__BingCascadeValues_003"].Ranges).RangeAddress;
        Assert.Equal((3, 3, 5, 3), (childRange.FirstAddress.RowNumber,
            childRange.FirstAddress.ColumnNumber, childRange.LastAddress.RowNumber,
            childRange.LastAddress.ColumnNumber));

        var data = workbook.Worksheet("Products");
        Assert.Equal("原材料", data.Cell("A2").GetString());
        var validations = data.DataValidations.ToArray();
        Assert.Equal(2, validations.Length);
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') == "Categories");
        Assert.Contains(validations, rule => rule.Value.TrimStart('=') ==
            "INDIRECT(VLOOKUP($A2,__BingCascadeMap_001,2,FALSE))");
    }

    /// <summary>
    /// 验证实体模板保留已有内容并应用级联下拉。
    /// </summary>
    [Fact]
    public async Task ExportEntityAsync_CascadeList_ShouldKeepTemplateAndStreamOpen()
    {
        var layout = ExcelEntity.Layout<PurchaseOrder>(builder => builder
            .Cell("Order", "A2", order => order.Category)
            .Cell("Order", "B2", order => order.Product)
            .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
            {
                ["材料"] = new[] { "钢板", "铜线" }
            })
            .DataValidation("Order", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition
                {
                    StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1
                },
                Type = ExcelDataValidationType.List,
                CascadeSourceName = "Categories",
                ParentColumnIndex = 0
            }));
        using var template = new MemoryStream();
        using (var original = new XLWorkbook())
        {
            original.Worksheets.Add("Order").Cell("C1").Value = "保留";
            original.SaveAs(template);
        }
        template.Position = 0;
        using var output = new MemoryStream();
        await new ClosedXmlExcelExporter().ExportForTemplateAsync(
            new PurchaseOrder { Category = "材料", Product = "钢板" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        Assert.True(output.CanWrite);
        Assert.True(template.CanRead);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal("保留", workbook.Worksheet("Order").Cell("C1").GetString());
        Assert.Equal("材料", workbook.Worksheet("Order").Cell("A2").GetString());
        Assert.Equal("钢板", workbook.Worksheet("Order").Cell("B2").GetString());
        Assert.Equal("INDIRECT(VLOOKUP($A2,__BingCascadeMap_001,2,FALSE))",
            Assert.Single(workbook.Worksheet("Order").DataValidations).Value.TrimStart('='));
    }

    /// <summary>
    /// 验证模板占用级联名称时不写入目标流。
    /// </summary>
    [Fact]
    public void Export_CascadeList_TemplateNameCollision_ShouldLeaveOutputEmpty()
    {
        using var templateBook = new XLWorkbook();
        var sheet = templateBook.Worksheets.Add("Products");
        sheet.Cell("A1").Value = "旧项";
        templateBook.DefinedNames.Add("Categories", sheet.Range("A1"));
        using var template = new MemoryStream();
        templateBook.SaveAs(template);
        template.Position = 0;
        var request = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
            {
                ["材料"] = new[] { "钢板" }
            })
            .AddSheet("Products", new[] { new ProductRow { Category = "材料" } }));

        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));
        Assert.Equal(0, output.Length);
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证原生级联公式导入按不支持策略返回结构化错误。
    /// </summary>
    [Fact]
    public void Import_CascadeFormula_ShouldReportUnsupportedWorkbookRule()
    {
        var request = ExcelExport.Workbook(book => book
            .CascadeList("Categories", new Dictionary<string, IReadOnlyList<string>>
            {
                ["材料"] = new[] { "钢板" }
            })
            .AddSheet("Products", new[] { new ProductRow { Category = "材料", Product = "钢板" } },
                sheet => sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1
                    },
                    Type = ExcelDataValidationType.List,
                    CascadeSourceName = "Categories", ParentColumnIndex = 0
                })));
        using var output = new MemoryStream();
        new ClosedXmlExcelExporter().Export(request, output);
        output.Position = 0;
        var import = ExcelImport.Workbook<ImportRoot>(book => book
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy.Fail)
            .Sheet<ProductRow>("Products", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var result = new ClosedXmlExcelImporter().Import(output, import);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Products", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
    }

    /// <summary>
    /// 承载商品导入结果。
    /// </summary>
    public sealed class ImportRoot
    {
        /// <summary>
        /// 获取商品行。
        /// </summary>
        public List<ProductRow> Rows { get; } = new();
    }

    /// <summary>
    /// 商品填写行。
    /// </summary>
    public sealed class ProductRow
    {
        /// <summary>
        /// 获取或设置商品类别。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置商品名称。
        /// </summary>
        public string Product { get; set; }
    }

    /// <summary>
    /// 采购单据。
    /// </summary>
    public sealed class PurchaseOrder
    {
        /// <summary>
        /// 获取或设置商品类别。
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// 获取或设置商品名称。
        /// </summary>
        public string Product { get; set; }
    }
}
