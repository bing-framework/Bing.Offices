using System;
using System.Collections.Generic;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 新增公共 Excel 请求契约的直接行为测试。
/// </summary>
public sealed class AbstractionsContractTest
{
    /// <summary>
    /// 测试 - 资源限制新增字段应允许关闭或使用正值，并拒绝非正值。
    /// </summary>
    [Fact]
    public void ExcelResourceLimits_NewCountLimits_ShouldValidate()
    {
        new ExcelResourceLimits
        {
            MaxSheets = null,
            MaxColumnsPerSheet = null,
            MaxCells = null
        }.Validate();

        new ExcelResourceLimits
        {
            MaxSheets = 2,
            MaxColumnsPerSheet = 10,
            MaxCells = 20
        }.Validate();

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxSheets = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxColumnsPerSheet = 0 }.Validate());
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelResourceLimits { MaxCells = 0 }.Validate());
    }

    /// <summary>
    /// 测试 - 行高允许空值和 Excel 最大值，拒绝非法浮点数或超出上限的值。
    /// </summary>
    [Fact]
    public void ExcelRowHeightOptions_ShouldValidatePointRange()
    {
        new ExcelRowHeightOptions
        {
            HeaderHeight = null,
            BodyHeight = 409.5
        }.Validate();

        foreach (var value in new[] { 0d, -1d, 409.5001d, double.NaN, double.PositiveInfinity })
        {
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                new ExcelRowHeightOptions { HeaderHeight = value }.Validate());
        }
    }

    /// <summary>
    /// 测试 - Sheet Builder 应将行高配置保存到不可变导出请求。
    /// </summary>
    [Fact]
    public void ExcelSheetExportBuilder_RowHeight_ShouldBuildRequest()
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet(
            "Data",
            new[] { new ExportRow { Code = "A" } },
            sheet => sheet.RowHeight(new ExcelRowHeightOptions
            {
                HeaderHeight = 24,
                BodyHeight = 18
            })));

        Assert.Equal(24, request.Sheets[0].RowHeight.HeaderHeight);
        Assert.Equal(18, request.Sheets[0].RowHeight.BodyHeight);
    }

    /// <summary>
    /// 测试 - Entity Layout 应保存 HasMany 关系快照。
    /// </summary>
    [Fact]
    public void ExcelEntityLayout_HasMany_ShouldExposeRelationSnapshot()
    {
        var layout = ExcelEntity.Layout<RelationWorkbook>(builder => builder
            .Cell("Data", "A1", workbook => workbook.Title)
            .HasMany(workbook => workbook.Parents, workbook => workbook.Children,
                parent => parent.Id,
                child => child.ParentId,
                parent => parent.Children));

        var relation = Assert.Single(layout.Relations);
        Assert.Equal(typeof(RelationParent), relation.ParentType);
        Assert.Equal(typeof(RelationChild), relation.ChildType);
        Assert.Throws<NotSupportedException>(() =>
            ((IList<ExcelRelationRequest>)layout.Relations).Add(relation));
    }

    /// <summary>
    /// 导出请求测试行。
    /// </summary>
    private sealed class ExportRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 关系测试工作簿。
    /// </summary>
    private sealed class RelationWorkbook
    {
        /// <summary>
        /// 获取标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取父实体集合。
        /// </summary>
        public ICollection<RelationParent> Parents { get; } = new List<RelationParent>();

        /// <summary>
        /// 获取子实体集合。
        /// </summary>
        public ICollection<RelationChild> Children { get; } = new List<RelationChild>();
    }

    /// <summary>
    /// 关系测试父实体。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取标识。
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 获取子实体导航集合。
        /// </summary>
        public ICollection<RelationChild> Children { get; } = new List<RelationChild>();
    }

    /// <summary>
    /// 关系测试子实体。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取父实体标识。
        /// </summary>
        public int ParentId { get; set; }
    }
}
