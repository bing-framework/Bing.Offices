using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// ClosedXML Workbook 工作表保护的职责级测试。
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
        var exporter = new ClosedXmlExcelExporter();
        using var output = new MemoryStream();
        await exporter.ExportAsync(request, output);
        using (var workbook = new XLWorkbook(new MemoryStream(output.ToArray())))
        {
            var sheet = workbook.Worksheet("Orders");
            Assert.True(sheet.Protection.IsProtected);
            Assert.False(sheet.Cell("A2").Style.Protection.Locked);
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
    /// 验证 XLSX 模板样式保留、工作表保护及可填写区域。
    /// </summary>
    [Fact]
    public void Export_Protection_ShouldLockSheetAndPreserveEditableTemplateStyle()
    {
        using var template = new MemoryStream();
        using (var sourceWorkbook = new XLWorkbook())
        {
            var sourceSheet = sourceWorkbook.Worksheets.Add("Orders");
            sourceSheet.Cell("A2").Value = "Template";
            sourceSheet.Cell("A2").Style.Fill.BackgroundColor = XLColor.Yellow;
            sourceWorkbook.SaveAs(template);
        }
        template.Position = 0;

        var protection = new ExcelSheetProtectionDefinition(new[]
        {
            new ExcelRangeDefinition { StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 0 }
        }, "sheet-password");
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("Orders", new[]
            {
                new ProtectionRow { Name = "First", Quantity = 1 },
                new ProtectionRow { Name = "Second", Quantity = 2 }
            }, sheet => sheet.Protection(protection)));
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().Export(request, output);

        Assert.True(output.CanWrite);
        output.Position = 0;
        using var result = new XLWorkbook(output);
        var sheet = result.Worksheet("Orders");
        Assert.True(sheet.Protection.IsProtected);
        Assert.True(sheet.Protection.IsPasswordProtected);
        Assert.True(sheet.Cell("A1").Style.Protection.Locked);
        Assert.False(sheet.Cell("A2").Style.Protection.Locked);
        Assert.False(sheet.Cell("A3").Style.Protection.Locked);
        Assert.True(sheet.Cell("B2").Style.Protection.Locked);
        Assert.Equal(XLColor.Yellow, sheet.Cell("A2").Style.Fill.BackgroundColor);
    }

    /// <summary>
    /// 验证 XLSX 的可填写区域越过行上限时在写入目标流前失败。
    /// </summary>
    [Fact]
    public void Export_ProtectionOutsideXlsxBounds_ShouldPreserveDestination()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Orders", new[] { new ProtectionRow { Name = "First", Quantity = 1 } },
                sheet => sheet.Protection(new ExcelSheetProtectionDefinition(new[]
                {
                    new ExcelRangeDefinition
                    {
                        StartRow = 1048576, EndRow = 1048576, StartColumn = 0, EndColumn = 0
                    }
                }))));
        var original = new byte[] { 0x42, 0x49, 0x4E, 0x47 };
        using var output = new MemoryStream();
        output.Write(original, 0, original.Length);
        output.Position = 0;

        Assert.Throws<BingOfficesConfigurationException>(() =>
            new ClosedXmlExcelExporter().Export(request, output));

        Assert.Equal(original, output.ToArray());
        Assert.True(output.CanWrite);
    }

    /// <summary>
    /// 验证实体布局在 XLSX 模板中应用保护、保留可填写单元格样式。
    /// </summary>
    [Fact]
    public void EntityLayout_Protection_ShouldApplyToTemplateAndPreserveEditableStyle()
    {
        using var template = CreateEntityTemplate(protectSheet: false);
        var definition = new ExcelSheetProtectionDefinition(new[]
        {
            new ExcelRangeDefinition { StartRow = 1, EndRow = 1, StartColumn = 1, EndColumn = 1 }
        }, "entity-password");
        var layout = ExcelEntity.Layout<ProtectionEntity>(builder => builder
            .Cell("单据", "A1", item => item.Title)
            .Protection("单据", definition));
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().ExportForTemplate(new ProtectionEntity { Title = "Entity protected" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("单据");
        Assert.True(sheet.Protection.IsProtected);
        Assert.True(sheet.Protection.IsPasswordProtected);
        Assert.Equal("Entity protected", sheet.Cell("A1").GetString());
        Assert.True(sheet.Cell("A1").Style.Protection.Locked);
        Assert.False(sheet.Cell("B2").Style.Protection.Locked);
        Assert.Equal(XLColor.Yellow, sheet.Cell("B2").Style.Fill.BackgroundColor);
    }

    /// <summary>
    /// 验证实体模板未配置保护时保留原工作表保护和单元格样式。
    /// </summary>
    [Fact]
    public void EntityLayout_WithoutProtection_ShouldPreserveTemplateProtection()
    {
        using var template = CreateEntityTemplate(protectSheet: true);
        var layout = ExcelEntity.Layout<ProtectionEntity>(builder => builder
            .Cell("单据", "A1", item => item.Title));
        using var output = new MemoryStream();

        new ClosedXmlExcelExporter().ExportForTemplate(new ProtectionEntity { Title = "Keep protection" }, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheet("单据");
        Assert.True(sheet.Protection.IsProtected);
        Assert.True(sheet.Cell("A1").Style.Protection.Locked);
        Assert.True(sheet.Cell("B2").Style.Protection.Locked);
        Assert.Equal(XLColor.Yellow, sheet.Cell("B2").Style.Fill.BackgroundColor);
    }

    /// <summary>
    /// 创建包含可编辑模板单元格的实体模板。
    /// </summary>
    /// <param name="protectSheet">是否预先保护工作表。</param>
    /// <returns>模板输入流。</returns>
    private static MemoryStream CreateEntityTemplate(bool protectSheet)
    {
        var template = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("单据");
            sheet.Cell("B2").Value = "Editable template";
            sheet.Cell("B2").Style.Fill.BackgroundColor = XLColor.Yellow;
            if (protectSheet)
                sheet.Protect("template-password");
            workbook.SaveAs(template);
        }
        template.Position = 0;
        return template;
    }

    /// <summary>
    /// 工作表保护实体模板的根模型。
    /// </summary>
    private sealed class ProtectionEntity
    {
        /// <summary>
        /// 获取或设置标题。
        /// </summary>
        public string Title { get; set; }
    }

    /// <summary>
    /// 工作表保护实体模板的明细模型。
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
