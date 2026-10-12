using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI Workbook 工作表保护的职责级测试。
/// </summary>
public sealed class WorkbookSheetProtectionTest
{
    /// <summary>
    /// 验证异步保护输出与预取消时目标流保持完整。
    /// </summary>
    [Fact]
    public async Task ExportAsync_Protection_ShouldWriteAndHonorPreCancellation()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Orders", new[] { new ProtectionRow { Name = "First" } },
                sheet => sheet.Protection(new ExcelSheetProtectionDefinition(new[]
                {
                    new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0 }
                }))));
        var exporter = new NpoiExcelExporter();
        using var output = new MemoryStream();
        await exporter.ExportAsync(request, output);
        using (var workbook = new XSSFWorkbook(new MemoryStream(output.ToArray())))
        {
            var sheet = workbook.GetSheet("Orders");
            Assert.True(sheet.Protect);
            Assert.False(sheet.GetRow(1).GetCell(0).CellStyle.IsLocked);
        }

        var original = new byte[] { 4, 5, 6 };
        using var canceledOutput = new MemoryStream(original);
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            exporter.ExportAsync(request, canceledOutput, canceled.Token));
        Assert.Equal(original, canceledOutput.ToArray());
    }

    /// <summary>
    /// 验证 XLS 和 XLSX 的模板样式保留、工作表保护及可填写区域。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void Export_Protection_ShouldLockSheetAndPreserveEditableTemplateStyle(ExcelFormat format)
    {
        using IWorkbook templateWorkbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        var templateSheet = templateWorkbook.CreateSheet("Orders");
        var templateCell = templateSheet.CreateRow(1).CreateCell(0);
        var templateStyle = templateWorkbook.CreateCellStyle();
        templateStyle.FillForegroundColor = IndexedColors.Yellow.Index;
        templateStyle.FillPattern = FillPattern.SolidForeground;
        templateCell.CellStyle = templateStyle;
        using var template = new MemoryStream();
        templateWorkbook.Write(template, leaveOpen: true);
        template.Position = 0;

        var protection = new ExcelSheetProtectionDefinition(new[]
        {
            new ExcelRangeDefinition { StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 0 }
        }, "sheet-password");
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("Orders", new[]
            {
                new ProtectionRow { Name = "First", Quantity = 1 },
                new ProtectionRow { Name = "Second", Quantity = 2 }
            }, sheet => sheet.Protection(protection)));
        using var output = new MemoryStream();

        new NpoiExcelExporter().Export(request, output);

        Assert.True(output.CanWrite);
        output.Position = 0;
        using var result = WorkbookFactory.Create(output);
        var sheet = result.GetSheet("Orders");
        Assert.True(sheet.Protect);
        Assert.True(sheet.GetRow(0).GetCell(0).CellStyle.IsLocked);
        Assert.False(sheet.GetRow(1).GetCell(0).CellStyle.IsLocked);
        Assert.False(sheet.GetRow(2).GetCell(0).CellStyle.IsLocked);
        Assert.True(sheet.GetRow(1).GetCell(1).CellStyle.IsLocked);
        Assert.Equal(IndexedColors.Yellow.Index, sheet.GetRow(1).GetCell(0).CellStyle.FillForegroundColor);
        Assert.Equal(FillPattern.SolidForeground, sheet.GetRow(1).GetCell(0).CellStyle.FillPattern);
    }

    /// <summary>
    /// 验证可填写区域越过格式列上限时在写入目标流前失败。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="column">超出格式上限的列索引。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, 256)]
    [InlineData(ExcelFormat.Xlsx, 16384)]
    public void Export_ProtectionOutsideBounds_ShouldPreserveDestination(ExcelFormat format, int column)
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .Format(format)
            .AddSheet("Orders", new[] { new ProtectionRow { Name = "First", Quantity = 1 } },
                sheet => sheet.Protection(new ExcelSheetProtectionDefinition(new[]
                {
                    new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = column, EndColumn = column }
                }))));
        var original = new byte[] { 0x42, 0x49, 0x4E, 0x47 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() => new NpoiExcelExporter().Export(request, output));

        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证实体布局在 XLS 模板、默认 XLSX 和 XLSX 模板中应用保护及可填写区域。
    /// </summary>
    /// <param name="format">目标 Excel 文件格式。</param>
    /// <param name="useTemplate">是否使用已有工作簿模板。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls, true)]
    [InlineData(ExcelFormat.Xlsx, false)]
    [InlineData(ExcelFormat.Xlsx, true)]
    public void EntityLayout_Protection_ShouldApplyAcrossFormatsAndTemplates(ExcelFormat format, bool useTemplate)
    {
        using var template = useTemplate ? CreateEntityTemplate(format, protectSheet: false) : null;
        var definition = new ExcelSheetProtectionDefinition(new[]
        {
            new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1 }
        }, "entity-password");
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title)
            .Protection("单据", definition));
        var entity = new Invoice { Title = "Entity protected" };
        using var output = new MemoryStream();

        if (useTemplate)
            new NpoiExcelExporter().ExportForTemplate(entity, layout,
                new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        else
            new NpoiExcelExporter().ExportEntity(entity, layout, output);

        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var sheet = workbook.GetSheet("单据");
        Assert.True(sheet.Protect);
        Assert.Equal("Entity protected", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.True(sheet.GetRow(0).GetCell(0).CellStyle.IsLocked);
        Assert.False(sheet.GetRow(1).GetCell(1).CellStyle.IsLocked);
        if (useTemplate)
        {
            Assert.Equal(IndexedColors.Yellow.Index,
                sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);
            Assert.Equal(FillPattern.SolidForeground, sheet.GetRow(1).GetCell(1).CellStyle.FillPattern);
        }
    }

    /// <summary>
    /// 验证实体模板未配置保护时保留原工作表保护。
    /// </summary>
    /// <param name="format">模板 Excel 文件格式。</param>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void EntityLayout_WithoutProtection_ShouldPreserveTemplateProtection(ExcelFormat format)
    {
        using var template = CreateEntityTemplate(format, protectSheet: true);
        var layout = ExcelEntity.Layout<Invoice>(builder => builder
            .Cell("单据", "A1", item => item.Title));
        using var output = new MemoryStream();

        new NpoiExcelExporter().ExportForTemplate(new Invoice { Title = "Keep protection" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using var workbook = WorkbookFactory.Create(output);
        var sheet = workbook.GetSheet("单据");
        Assert.True(sheet.Protect);
        Assert.Equal("Keep protection", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(IndexedColors.Yellow.Index, sheet.GetRow(1).GetCell(1).CellStyle.FillForegroundColor);
    }

    /// <summary>
    /// 创建包含可编辑模板单元格的实体模板。
    /// </summary>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="protectSheet">是否预先保护工作表。</param>
    /// <returns>模板输入流。</returns>
    private static MemoryStream CreateEntityTemplate(ExcelFormat format, bool protectSheet)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("单据");
        var editableCell = sheet.CreateRow(1).CreateCell(1);
        var style = workbook.CreateCellStyle();
        style.FillForegroundColor = IndexedColors.Yellow.Index;
        style.FillPattern = FillPattern.SolidForeground;
        editableCell.CellStyle = style;
        if (protectSheet)
            sheet.ProtectSheet("template-password");
        var template = new MemoryStream();
        workbook.Write(template, leaveOpen: true);
        template.Position = 0;
        return template;
    }

    /// <summary>
    /// 实体模板测试单据。
    /// </summary>
    public sealed class Invoice
    {
        /// <summary>
        /// 获取或设置标题。
        /// </summary>
        public string Title { get; set; }
    }

    /// <summary>
    /// 工作表保护测试行。
    /// </summary>
    private sealed class ProtectionRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
