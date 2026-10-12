using System;
using System.Collections.Generic;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 字典与名称列表的公共模型测试。
/// </summary>
public sealed class LookupListModelTest
{
    /// <summary>
    /// 验证字典构建时去重并隔离调用方列表。
    /// </summary>
    [Fact]
    public void LookupList_ShouldNormalizeAndSnapshotValues()
    {
        var values = new List<string> { " 中文 ", "中文", "=A1", "A,B", "A\nB" };
        var request = ExcelExport.Workbook(builder => builder
            .LookupList("Names", values)
            .AddSheet("Data", new[] { new Row() }));
        values[0] = "Changed";

        Assert.Equal(new[] { "中文", "=A1", "A,B", "A\nB" }, request.LookupLists[0].Values);
    }

    /// <summary>
    /// 验证空值、重名和非法名称在布局构建时被拒绝。
    /// </summary>
    [Fact]
    public void LookupList_ShouldRejectInvalidDefinitions()
    {
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder
            .LookupList("Names", new[] { " " }).AddSheet("Data", new[] { new Row() })));
        Assert.Throws<ArgumentException>(() => ExcelExport.Workbook(builder => builder
            .LookupList("Names", new[] { "A" })
            .LookupList("names", new[] { "B" }).AddSheet("Data", new[] { new Row() })));
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Row>(builder => builder
            .Cell("Data", "A1", item => item.Status)
            .LookupList("A1", new[] { "A" })));
    }

    /// <summary>
    /// 验证列表只能选择内联值或名称来源之一。
    /// </summary>
    [Fact]
    public void NamedValidation_ShouldRequireExactlyOneSource()
    {
        var range = new ExcelRangeDefinition
        {
            StartRow = 1, EndRow = 1, StartColumn = 0, EndColumn = 0
        };
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition
        {
            Range = range, Type = ExcelDataValidationType.List
        }.Validate());
        Assert.Throws<ArgumentException>(() => new ExcelDataValidationDefinition
        {
            Range = range, Type = ExcelDataValidationType.List,
            Values = new[] { "A" }, ListSourceName = "Names"
        }.Validate());
        new ExcelDataValidationDefinition
        {
            Range = range, Type = ExcelDataValidationType.List, ListSourceName = "Names"
        }.Validate();
    }

    /// <summary>
    /// 接收模型绑定的测试行。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置状态。
        /// </summary>
        public string Status { get; set; }
    }
}
