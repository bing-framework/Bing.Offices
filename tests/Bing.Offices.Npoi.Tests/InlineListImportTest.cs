using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Entities;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 XLS 与 XLSX 显式列表的导入行为。
/// </summary>
public sealed class InlineListImportTest
{
    /// <summary>
    /// 验证 Workbook 导入接纳合法值并报告非法值坐标。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void WorkbookImport_ExplicitList_ShouldAcceptAndRejectValues(bool xls)
    {
        using var input = new MemoryStream(CreateWorkbook(xls, sheet =>
        {
            sheet.CreateRow(0).CreateCell(0).SetCellValue("Value");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Yes");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Invalid");
            AddExplicitList(sheet, 1, 2, 0);
        }));
        var request = ExcelImport.Workbook<WorkbookModel>(builder => builder
            .ValidationMode(ExcelImportValidationMode.WorkbookRules)
            .Sheet("Data", model => model.Rows,
                sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));

        var result = new NpoiExcelImporter().Import(input, request);

        Assert.Equal("Yes", Assert.Single(result.Workbook.Rows).Value);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Contains("显式列表", error.Message);
    }

    /// <summary>
    /// 验证 Entity 固定单元格和明细使用同一显式列表规则。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityImport_ExplicitList_ShouldValidateFixedAndDetailCells(bool xls)
    {
        using var input = new MemoryStream(CreateWorkbook(xls, sheet =>
        {
            sheet.CreateRow(0).CreateCell(1).SetCellValue("Invalid");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Yes");
            sheet.CreateRow(3).CreateCell(0).SetCellValue("Invalid");
            AddExplicitList(sheet, 0, 0, 1);
            AddExplicitList(sheet, 2, 3, 0);
        }));
        var layout = ExcelEntity.Layout<EntityModel>(builder => builder
            .Cell("Data", "B1", entity => entity.Status)
            .ListRegion("Data", "A2", entity => entity.Rows,
                region => region.End("A4")));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.WorkbookRules, ExcelUnsupportedFeaturePolicy.Fail);

        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);

        Assert.Null(result.Entity.Status);
        Assert.Equal("Yes", Assert.Single(result.Entity.Rows).Value);
        Assert.Equal(new[] { (1, 2), (4, 1) }, result.Errors
            .Select(error => (error.RowIndex, error.ColumnIndex))
            .OrderBy(location => location.RowIndex).ToArray());
        Assert.All(result.Errors, error =>
        {
            Assert.Equal(ExcelImportErrorCode.WorkbookValidation, error.Code);
            Assert.Contains("显式列表", error.Message);
        });
    }

    /// <summary>
    /// 创建包含业务工作表的真实工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(bool xls, Action<ISheet> configure)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        configure(workbook.CreateSheet("Data"));
        using var output = new MemoryStream();
        workbook.Write(output, leaveOpen: true);
        return output.ToArray();
    }

    /// <summary>
    /// 给指定区域添加显式列表规则。
    /// </summary>
    private static void AddExplicitList(ISheet sheet, int firstRow, int lastRow, int column)
    {
        var helper = sheet.GetDataValidationHelper();
        sheet.AddValidationData(helper.CreateValidation(
            helper.CreateExplicitListConstraint(new[] { "Yes", "No" }),
            new CellRangeAddressList(firstRow, lastRow, column, column)));
    }

    /// <summary>
    /// 接收 Workbook 明细。
    /// </summary>
    public sealed class WorkbookModel
    {
        /// <summary>
        /// 获取导入行。
        /// </summary>
        public List<RowModel> Rows { get; } = new();
    }

    /// <summary>
    /// 接收 Entity 固定单元格和明细。
    /// </summary>
    public sealed class EntityModel
    {
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 获取或设置明细行。
        /// </summary>
        public List<RowModel> Rows { get; set; } = new();
    }

    /// <summary>
    /// 接收单行值。
    /// </summary>
    public sealed class RowModel
    {
        /// <summary>
        /// 获取或设置列表值。
        /// </summary>
        public string Value { get; set; }
    }
}
