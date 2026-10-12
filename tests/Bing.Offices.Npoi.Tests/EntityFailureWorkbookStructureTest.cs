using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证实体失败工作簿对动态字典和关系列表的结构保留。
/// </summary>
public sealed class EntityFailureWorkbookStructureTest
{
    /// <summary>
    /// 验证隐式动态字典转换失败时保留动态列表头和值边界。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ErrorRowsOnly_ShouldRetainImplicitDynamicDictionary(bool xls)
    {
        using var input = new MemoryStream(CreateDynamicWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var result = new NpoiExcelImporter().ImportEntity(input, CreateDynamicLayout(),
            CreateOptions(failure));

        Assert.False(result.IsSuccess);
        Assert.NotEmpty(result.Errors);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.GetSheet("动态");
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Zone", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("Score", sheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("A", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("North", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("wrong", sheet.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 验证关系列表中的子项错误时保留父子列表的结构行。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ErrorRowsOnly_ShouldRetainRelatedListStructure(bool xls)
    {
        using var input = new MemoryStream(CreateRelationWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var result = new NpoiExcelImporter().ImportEntity(input, CreateRelationLayout(),
            CreateOptions(failure));

        Assert.False(result.IsSuccess);
        Assert.NotEmpty(result.Errors);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        var parents = output.GetSheet("父项");
        var children = output.GetSheet("子项");
        Assert.Equal("Id", parents.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("ParentId", children.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Name", children.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("Quantity", children.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("P-1", children.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("detail", children.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("wrong", children.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 创建仅输出到调用方流的失败工作簿选项。
    /// </summary>
    /// <param name="destination">接收失败工作簿的流。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions CreateOptions(Stream destination) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            });

    /// <summary>
    /// 创建隐式动态字典布局。
    /// </summary>
    /// <returns>动态字典实体布局。</returns>
    private static ExcelEntityLayout<DynamicDocument> CreateDynamicLayout()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(DynamicLine.Code), Title = "Code" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string", Order = 0, ColumnIndex = 1 },
                new() { Key = "score", Title = "Score", DataTypeName = "decimal", Order = 1, ColumnIndex = 2 }
            }
        };
        return ExcelEntity.Layout<DynamicDocument>(builder => builder
            .ListRegion("动态", "A1", document => document.Lines,
                region => region.Mapping(mapping)));
    }

    /// <summary>
    /// 创建带父子关系的两个列表区域布局。
    /// </summary>
    /// <returns>关系实体布局。</returns>
    private static ExcelEntityLayout<RelationDocument> CreateRelationLayout() =>
        ExcelEntity.Layout<RelationDocument>(builder => builder
            .ListRegion("父项", "A1", document => document.Parents)
            .ListRegion("子项", "A1", document => document.Children)
            .HasMany(document => document.Parents, document => document.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 创建隐式动态字典转换失败的工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 工作簿。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateDynamicWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("动态");
        WriteHeader(sheet, "Code", "Zone", "Score");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A");
        row.CreateCell(1).SetCellValue("North");
        row.CreateCell(2).SetCellValue("wrong");
        return WriteWorkbook(workbook);
    }

    /// <summary>
    /// 创建子项数量转换失败的父子工作簿。
    /// </summary>
    /// <param name="xls">是否生成 XLS 工作簿。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateRelationWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        WriteHeader(parents, "Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        var children = workbook.CreateSheet("子项");
        WriteHeader(children, "ParentId", "Name", "Quantity");
        var row = children.CreateRow(1);
        row.CreateCell(0).SetCellValue("P-1");
        row.CreateCell(1).SetCellValue("detail");
        row.CreateCell(2).SetCellValue("wrong");
        return WriteWorkbook(workbook);
    }

    /// <summary>
    /// 写入工作表表头。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="headers">表头文本。</param>
    private static void WriteHeader(ISheet sheet, params string[] headers)
    {
        var row = sheet.CreateRow(0);
        for (var index = 0; index < headers.Length; index++)
            row.CreateCell(index).SetCellValue(headers[index]);
    }

    /// <summary>
    /// 将工作簿序列化为字节。
    /// </summary>
    /// <param name="workbook">待序列化的工作簿。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] WriteWorkbook(IWorkbook workbook)
    {
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    private sealed class DynamicDocument
    {
        /// <summary>
        /// 获取动态明细集合。
        /// </summary>
        public List<DynamicLine> Lines { get; } = new();
    }

    private sealed class DynamicLine
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段字典。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Fields { get; set; } = new();
    }

    private sealed class RelationDocument
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关联子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
