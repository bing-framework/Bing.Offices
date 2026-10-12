using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI 级联字典模板。
/// </summary>
public sealed class CascadeListExportTest
{
    /// <summary>
    /// 验证两种格式均写入完整文本字典、名称和相对行公式。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_CascadeList_WritesHiddenMapAndFormula(ExcelFormat format)
    {
        var children = new List<string> { " 手机,配件 ", "=原样", "手机,配件" };
        var source = new Dictionary<string, IReadOnlyList<string>>
        {
            [" 电子 "] = children,
            ["食品"] = new[] { "饮料\n冷藏", "零食" }
        };
        var request = ExcelExport.Workbook(book => book.Format(format)
            .CascadeList("Category", source)
            .AddSheet("Products", new[] { new Row { Parent = "电子", Child = "=原样" } }, sheet =>
                sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition
                    {
                        StartRow = 1, EndRow = 20, StartColumn = 1, EndColumn = 1
                    },
                    Type = ExcelDataValidationType.List,
                    CascadeSourceName = "Category",
                    ParentColumnIndex = 0
                })));
        children[0] = "已变更";
        source["食品"] = new[] { "已变更" };

        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        Assert.True(output.CanWrite);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        Assert.Equal(2, workbook.NumberOfSheets);
        var helper = workbook.GetSheetAt(1);
        Assert.StartsWith("__BingCascade_", helper.SheetName, StringComparison.Ordinal);
        Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(1));
        Assert.Equal("电子", helper.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("食品", helper.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("手机,配件", helper.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("=原样", helper.GetRow(1).GetCell(2).StringCellValue);
        Assert.Equal("饮料\n冷藏", helper.GetRow(2).GetCell(2).StringCellValue);
        Assert.Equal(CellType.String, helper.GetRow(1).GetCell(2).CellType);
        var names = workbook.GetAllNames();
        Assert.Equal(4, names.Count);
        Assert.Contains(names, name => name.NameName == "Category"
            && name.RefersToFormula.EndsWith("!$A$1:$A$2", StringComparison.Ordinal));
        Assert.Contains(names, name => name.RefersToFormula.EndsWith("!$C$1:$C$2", StringComparison.Ordinal));
        Assert.Contains(names, name => name.RefersToFormula.EndsWith("!$C$3:$C$4", StringComparison.Ordinal));
        var map = Assert.Single(names, name => name.RefersToFormula.EndsWith("!$A$1:$B$2",
            StringComparison.Ordinal));
        Assert.Equal("INDIRECT(VLOOKUP($A2," + map.NameName + ",2,FALSE))",
            Assert.Single(workbook.GetSheet("Products").GetDataValidations()).ValidationConstraint.Formula1);
    }

    /// <summary>
    /// 验证实体布局也使用级联字典。
    /// </summary>
    [Fact]
    public void ExportEntity_CascadeList_WritesValidation()
    {
        var layout = ExcelEntity.Layout<Row>(builder => builder
            .Cell("Products", "A2", row => row.Parent)
            .CascadeList("Category", new Dictionary<string, IReadOnlyList<string>>
            {
                ["电子"] = new[] { "手机" }
            })
            .DataValidation("Products", new ExcelDataValidationDefinition
            {
                Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 3, StartColumn = 1, EndColumn = 1 },
                Type = ExcelDataValidationType.List, CascadeSourceName = "Category", ParentColumnIndex = 0
            }));
        using var output = new MemoryStream();
        new NpoiExcelExporter().ExportEntity(new Row { Parent = "电子" }, layout, output);
        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        Assert.Equal("电子", workbook.GetSheet("Products").GetRow(1).GetCell(0).StringCellValue);
        Assert.Single(workbook.GetSheet("Products").GetDataValidations());
        Assert.Equal(SheetVisibility.Hidden, workbook.GetSheetVisibility(1));
    }

    /// <summary>
    /// 验证名称冲突在写目标流前失败。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_CascadeNameCollision_LeavesOutputEmpty(ExcelFormat format)
    {
        using IWorkbook book = format == ExcelFormat.Xls
            ? new NPOI.HSSF.UserModel.HSSFWorkbook()
            : new NPOI.XSSF.UserModel.XSSFWorkbook();
        book.CreateSheet("Products");
        var name = book.CreateName();
        name.NameName = "Category";
        name.RefersToFormula = "Products!$A$1";
        using var template = new MemoryStream();
        book.Write(template, leaveOpen: true);
        template.Position = 0;
        var request = ExcelExport.Workbook(builder => builder.Format(format)
            .UseTemplate(template, leaveOpen: true)
            .CascadeList("Category", new Dictionary<string, IReadOnlyList<string>>
            {
                ["电子"] = new[] { "手机" }
            })
            .AddSheet("Products", new[] { new Row() }));
        using var output = new MemoryStream();
        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().Export(request, output));
        Assert.Equal(0, output.Length);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证级联参数互斥、父列和源数据快照。
    /// </summary>
    [Fact]
    public void CascadeModel_RejectsInvalidSourcesAndSnapshotsValues()
    {
        var values = new Dictionary<string, IReadOnlyList<string>> { ["电子"] = new List<string> { " 手机 " } };
        var request = ExcelExport.Workbook(book => book.CascadeList("Category", values)
            .AddSheet("Products", new[] { new Row() }));
        values["电子"] = new[] { "变更" };
        Assert.Equal("手机", Assert.Single(request.CascadeLists[0].Entries[0].Values));
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(book => book
            .LookupList("Category", new[] { "电子" })
            .CascadeList("category", values).AddSheet("Products", new[] { new Row() })));
        var range = new ExcelRangeDefinition { StartRow = 1, EndRow = 2, StartColumn = 1, EndColumn = 1 };
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition
        {
            Type = ExcelDataValidationType.List, Range = range,
            CascadeSourceName = "Category", ParentColumnIndex = 1
        }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition
        {
            Type = ExcelDataValidationType.List, Range = range,
            CascadeSourceName = "Category", ParentColumnIndex = 0,
            ListSourceName = "Other"
        }.Validate());
    }

    /// <summary>
    /// 验证原生级联公式导入按不支持策略返回结构化错误。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Import_CascadeFormula_ReportsUnsupportedWorkbookRule(ExcelFormat format)
    {
        var request = ExcelExport.Workbook(book => book.Format(format)
            .CascadeList("Category", new Dictionary<string, IReadOnlyList<string>>
            {
                ["电子"] = new[] { "手机" }
            })
            .AddSheet("Products", new[] { new Row { Parent = "电子", Child = "手机" } }, sheet =>
                sheet.DataValidation(new ExcelDataValidationDefinition
                {
                    Range = new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1 },
                    Type = ExcelDataValidationType.List, CascadeSourceName = "Category", ParentColumnIndex = 0
                })));
        using var output = new MemoryStream();
        new NpoiExcelExporter().Export(request, output);
        output.Position = 0;
        var import = ExcelImport.Workbook<Root>(book => book
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .UnsupportedFeaturePolicy(ExcelUnsupportedFeaturePolicy.Fail)
            .Sheet<Row>("Products", root => root.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var result = new NpoiExcelImporter().Import(output, import);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Products", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
    }

    /// <summary>
    /// 承载导入结果。
    /// </summary>
    public sealed class Root
    {
        /// <summary>
        /// 获取商品行。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 提供测试明细。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置上级文本。
        /// </summary>
        public string Parent { get; set; }
        /// <summary>
        /// 获取或设置下级文本。
        /// </summary>
        public string Child { get; set; }
    }
}
