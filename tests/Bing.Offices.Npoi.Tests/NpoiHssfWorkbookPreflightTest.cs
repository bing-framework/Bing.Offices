using System.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 HSSF 工作簿在 DOM 创建前执行物理结构资源限制。
/// </summary>
public sealed class NpoiHssfWorkbookPreflightTest
{
    /// <summary>
    /// 验证同步与异步实体导入在 HSSF 工作簿超限时于预检阶段失败。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="limitName">超限的结构预算名称。</param>
    [Theory]
    [InlineData(false, "MaxCells")]
    [InlineData(true, "MaxCells")]
    [InlineData(false, "MaxCellsWithBlank")]
    [InlineData(true, "MaxCellsWithBlank")]
    [InlineData(false, "MaxColumnsPerSheet")]
    [InlineData(true, "MaxColumnsPerSheet")]
    [InlineData(false, "MaxSheets")]
    [InlineData(true, "MaxSheets")]
    public async Task ImportEntity_ShouldRejectHssfLimitsBeforeDom(bool asynchronous, string limitName)
    {
        using var source = new MemoryStream(CreateWorkbook(limitName));
        var options = new ExcelEntityImportOptions(CreateLimits(limitName));
        var importer = new NpoiExcelImporter();
        var layout = CreateLayout();

        BingOfficesResourceLimitException exception;
        if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityAsync(source, layout, options));
        }
        else
        {
            exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntity(source, layout, options));
        }

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证模板导入在打开超限的 HSSF 模板前执行资源预检。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportForTemplate_ShouldRejectHssfTemplateBeforeDom(bool asynchronous)
    {
        using var template = new MemoryStream(CreateWorkbook("MaxSheets"));
        using var source = new MemoryStream(CreateWorkbook(null));
        var importer = new NpoiExcelImporter();
        var layout = CreateLayout();
        var templateOptions = new ExcelEntityTemplateOptions(template, leaveOpen: true);
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxSheets = 1 });

        BingOfficesResourceLimitException exception;
        if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportForTemplateAsync(source, layout, templateOptions, importOptions));
        }
        else
        {
            exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportForTemplate(source, layout, templateOptions, importOptions));
        }

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.True(template.CanRead);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 结构限制允许恰好位于预算边界的工作簿。
    /// </summary>
    [Fact]
    public void ImportEntity_ShouldAllowHssfLimitsAtBoundary()
    {
        using var source = new MemoryStream(CreateWorkbook(null));
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxSheets = 1,
            MaxColumnsPerSheet = 1,
            MaxCells = 2
        });

        var result = new NpoiExcelImporter().ImportEntity(source, CreateLayout(), options);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A" }, result.Entity.Lines.Select(line => line.Code));
    }

    /// <summary>
    /// 根据指定结构预算创建资源限制。
    /// </summary>
    private static ExcelResourceLimits CreateLimits(string limitName) => limitName switch
    {
        "MaxCells" => new ExcelResourceLimits { MaxCells = 1 },
        "MaxCellsWithBlank" => new ExcelResourceLimits { MaxCells = 2 },
        "MaxColumnsPerSheet" => new ExcelResourceLimits { MaxColumnsPerSheet = 1 },
        "MaxSheets" => new ExcelResourceLimits { MaxSheets = 1 },
        _ => throw new ArgumentOutOfRangeException(nameof(limitName))
    };

    /// <summary>
    /// 创建单明细区域实体布局。
    /// </summary>
    private static ExcelEntityLayout<Document> CreateLayout() =>
        ExcelEntity.Layout<Document>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建用于结构资源预检的 HSSF 工作簿。
    /// </summary>
    /// <param name="limitName">需要触发的资源限制；为空时创建单 Sheet 两单元格工作簿。</param>
    private static byte[] CreateWorkbook(string limitName)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Code");
        var dataRow = sheet.CreateRow(1);
        dataRow.CreateCell(0).SetCellValue("A");
        if (limitName == "MaxColumnsPerSheet")
            dataRow.CreateCell(1).SetCellValue("额外列");
        if (limitName == "MaxSheets")
            workbook.CreateSheet("额外工作表");
        if (limitName == "MaxCellsWithBlank")
            sheet.CreateRow(2).CreateCell(0, CellType.Blank);

        using var output = new MemoryStream();
        workbook.Write(output);
        return output.ToArray();
    }

    /// <summary>
    /// 保存资源预检测试根模型。
    /// </summary>
    public sealed class Document
    {
        /// <summary>
        /// 获取或设置明细集合。
        /// </summary>
        public List<Line> Lines { get; set; } = new();
    }

    /// <summary>
    /// 保存资源预检测试明细。
    /// </summary>
    public sealed class Line
    {
        /// <summary>
        /// 获取或设置明细编码。
        /// </summary>
        public string Code { get; set; }
    }
}
