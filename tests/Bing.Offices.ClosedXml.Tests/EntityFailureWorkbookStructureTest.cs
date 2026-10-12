using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Configurations;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 ClosedXML Entity 失败工作簿保留动态列和列表关系结构。
/// </summary>
public sealed class EntityFailureWorkbookStructureTest
{
    /// <summary>
    /// 验证单个隐式动态字典的转换失败保留动态表头和原始错误值。
    /// </summary>
    [Fact]
    public void ErrorRowsOnly_DynamicDictionaryConversionFailure_ShouldRetainHeaderAndValue()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(DynamicRow.Code), Title = "Code" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "score", Title = "Score", DataTypeName = "int32", ColumnIndex = 1 }
            }
        };
        var layout = ExcelEntity.Layout<DynamicRoot>(builder => builder
            .ListRegion("Data", "A1", root => root.Rows, region => region.Mapping(mapping)));
        using var input = new MemoryStream(CreateDynamicFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = FailureOptions(failure);

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("score", error.PropertyName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        using var workbook = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        var data = workbook.Worksheet("Data");
        Assert.Equal(new[] { "Code", "Score" }, new[]
        {
            data.Cell("A1").GetString(), data.Cell("B1").GetString()
        });
        Assert.Equal("C-1", data.Cell("A2").GetString());
        Assert.Equal("wrong", data.Cell("B2").GetString());
        Assert.True(workbook.TryGetWorksheet("_ImportErrors", out var summary));
        Assert.Contains(error.Message, summary.CellsUsed().Select(cell => cell.GetString()));
    }

    /// <summary>
    /// 验证子列表转换失败时保留父列表布局、两张工作表的表头和失败原值。
    /// </summary>
    /// <remarks>
    /// ErrorRowsOnly 会保留无错误父工作表的表头，但不会保留父列表中的成功业务行。
    /// </remarks>
    [Fact]
    public void ErrorRowsOnly_RelatedChildConversionFailure_ShouldRetainBothLists()
    {
        var layout = ExcelEntity.Layout<RelationRoot>(builder => builder
            .ListRegion("Parents", "A1", root => root.Parents)
            .ListRegion("Children", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, System.StringComparer.OrdinalIgnoreCase));
        using var input = new MemoryStream(CreateRelatedFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, FailureOptions(failure));

        Assert.False(result.IsSuccess);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("Children", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        using var workbook = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        Assert.Equal(new[] { "Children", "Parents", "_ImportErrors" }.OrderBy(name => name,
                System.StringComparer.Ordinal),
            workbook.Worksheets.Select(sheet => sheet.Name).OrderBy(name => name,
                System.StringComparer.Ordinal));
        var parents = workbook.Worksheet("Parents");
        Assert.Equal("Id", parents.Cell("A1").GetString());
        Assert.Equal(1, parents.LastRowUsed().RowNumber());
        var children = workbook.Worksheet("Children");
        Assert.Equal(new[] { "ParentId", "Quantity" }, new[]
        {
            children.Cell("A1").GetString(), children.Cell("B1").GetString()
        });
        Assert.Equal("P-1", children.Cell("A2").GetString());
        Assert.Equal("wrong", children.Cell("B2").GetString());
        Assert.True(workbook.TryGetWorksheet("_ImportErrors", out var summary));
        Assert.Contains(error.Message, summary.CellsUsed().Select(cell => cell.GetString()));
    }

    /// <summary>
    /// 创建仅输出失败行和错误摘要的导入选项。
    /// </summary>
    /// <param name="destination">接收失败工作簿的流。</param>
    /// <returns>用于继续处理校验错误并输出失败工作簿的选项。</returns>
    private static ExcelEntityImportOptions FailureOptions(Stream destination) =>
        new(null, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            });

    /// <summary>
    /// 创建动态 Score 列包含非数值原文的真实工作簿。
    /// </summary>
    /// <returns>失败工作簿测试输入字节。</returns>
    private static byte[] CreateDynamicFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("A1").Value = "Code";
        sheet.Cell("B1").Value = "Score";
        sheet.Cell("A2").Value = "C-1";
        sheet.Cell("B2").Value = "wrong";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含有效父行和数量转换错误子行的真实工作簿。
    /// </summary>
    /// <returns>关系失败工作簿测试输入字节。</returns>
    private static byte[] CreateRelatedFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.Worksheets.Add("Parents");
        parents.Cell("A1").Value = "Id";
        parents.Cell("A2").Value = "P-1";
        var children = workbook.Worksheets.Add("Children");
        children.Cell("A1").Value = "ParentId";
        children.Cell("B1").Value = "Quantity";
        children.Cell("A2").Value = "P-1";
        children.Cell("B2").Value = "wrong";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 保存包含隐式动态列字典的实体根对象。
    /// </summary>
    private sealed class DynamicRoot
    {
        /// <summary>
        /// 获取动态列测试行集合。
        /// </summary>
        public List<DynamicRow> Rows { get; } = new();
    }

    /// <summary>
    /// 保存动态列转换测试行。
    /// </summary>
    private sealed class DynamicRow
    {
        /// <summary>
        /// 获取或设置固定业务编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取动态列值集合。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Fields { get; set; } = new();
    }

    /// <summary>
    /// 保存父列表和子列表的关联导入结果。
    /// </summary>
    private sealed class RelationRoot
    {
        /// <summary>
        /// 获取父列表。
        /// </summary>
        public List<ParentRow> Parents { get; set; } = new();

        /// <summary>
        /// 获取子列表。
        /// </summary>
        public List<ChildRow> Children { get; set; } = new();
    }

    /// <summary>
    /// 保存关系导入中的父行。
    /// </summary>
    private sealed class ParentRow
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关联到当前父项的子行。
        /// </summary>
        public List<ChildRow> Children { get; } = new();
    }

    /// <summary>
    /// 保存关系导入中的子行。
    /// </summary>
    private sealed class ChildRow
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置需要转换为整数的数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
