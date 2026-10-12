using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Entities;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证公开实体布局入口统一校验工作表名称。
/// </summary>
public sealed class EntitySheetNameValidationTest
{
    /// <summary>
    /// 获取 Excel 禁止的工作表名称样例。
    /// </summary>
    public static IEnumerable<object[]> InvalidSheetNames
    {
        get
        {
            yield return new object[] { "Bad:Name" };
            yield return new object[] { "Bad\\Name" };
            yield return new object[] { "Bad/Name" };
            yield return new object[] { "Bad?Name" };
            yield return new object[] { "Bad*Name" };
            yield return new object[] { "Bad[Name" };
            yield return new object[] { "Bad]Name" };
            yield return new object[] { "'Leading" };
            yield return new object[] { "Trailing'" };
            yield return new object[] { new string('S', 32) };
            yield return new object[] { "\0" };
            yield return new object[] { "\u0003" };
        }
    }

    /// <summary>
    /// 获取合法中文名称和 Excel 长度上限名称。
    /// </summary>
    public static IEnumerable<object[]> ValidSheetNames
    {
        get
        {
            yield return new object[] { "采购明细" };
            yield return new object[] { new string('S', 31) };
        }
    }

    /// <summary>
    /// 验证固定单元格配置拒绝非法工作表名称。
    /// </summary>
    /// <param name="sheetName">待验证的非法工作表名称。</param>
    [Theory]
    [MemberData(nameof(InvalidSheetNames))]
    public void Cell_ShouldRejectInvalidSheetName(string sheetName)
    {
        var exception = Assert.Throws<System.ArgumentException>(() => ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.Cell(sheetName, "A1", document => document.Title)));

        Assert.Contains("工作表名称无效", exception.Message);
    }

    /// <summary>
    /// 验证列表区域配置拒绝非法工作表名称。
    /// </summary>
    /// <param name="sheetName">待验证的非法工作表名称。</param>
    [Theory]
    [MemberData(nameof(InvalidSheetNames))]
    public void ListRegion_ShouldRejectInvalidSheetName(string sheetName)
    {
        var exception = Assert.Throws<System.ArgumentException>(() => ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.ListRegion(sheetName, "A1", document => document.Lines)));

        Assert.Contains("工作表名称无效", exception.Message);
    }

    /// <summary>
    /// 验证合并区域配置拒绝非法工作表名称。
    /// </summary>
    /// <param name="sheetName">待验证的非法工作表名称。</param>
    [Theory]
    [MemberData(nameof(InvalidSheetNames))]
    public void Merge_ShouldRejectInvalidSheetName(string sheetName)
    {
        var exception = Assert.Throws<System.ArgumentException>(() => ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.Cell("基础", "A1", document => document.Title)
                .Merge(sheetName, "B2:C3")));

        Assert.Contains("工作表名称无效", exception.Message);
    }

    /// <summary>
    /// 验证固定单元格保留合法工作表名称和布局描述。
    /// </summary>
    /// <param name="sheetName">合法的工作表名称。</param>
    [Theory]
    [MemberData(nameof(ValidSheetNames))]
    public void Cell_ShouldBuildWithValidSheetName(string sheetName)
    {
        var layout = ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.Cell(sheetName, "A1", document => document.Title));

        var cell = Assert.Single(layout.Cells);
        Assert.Equal(sheetName, cell.SheetName);
        Assert.Equal("A1", cell.Reference.Address);
        Assert.Equal(nameof(SheetNameDocument.Title), cell.PropertyName);
        Assert.Empty(layout.ListRegions);
        Assert.Empty(layout.Merges);
    }

    /// <summary>
    /// 验证列表区域保留合法工作表名称和布局描述。
    /// </summary>
    /// <param name="sheetName">合法的工作表名称。</param>
    [Theory]
    [MemberData(nameof(ValidSheetNames))]
    public void ListRegion_ShouldBuildWithValidSheetName(string sheetName)
    {
        var layout = ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.ListRegion(sheetName, "A1", document => document.Lines));

        var region = Assert.Single(layout.ListRegions);
        Assert.Equal(sheetName, region.SheetName);
        Assert.Equal("A1", region.Start.Address);
        Assert.Equal(nameof(SheetNameDocument.Lines), region.PropertyName);
        Assert.Empty(layout.Cells);
        Assert.Empty(layout.Merges);
    }

    /// <summary>
    /// 验证合并区域保留合法工作表名称和完整布局描述。
    /// </summary>
    /// <param name="sheetName">合法的工作表名称。</param>
    [Theory]
    [MemberData(nameof(ValidSheetNames))]
    public void Merge_ShouldBuildWithValidSheetName(string sheetName)
    {
        var layout = ExcelEntity.Layout<SheetNameDocument>(builder =>
            builder.Cell(sheetName, "A1", document => document.Title)
                .Merge(sheetName, "B2:C3"));

        var cell = Assert.Single(layout.Cells);
        var merge = Assert.Single(layout.Merges);
        Assert.Equal(sheetName, cell.SheetName);
        Assert.Equal("A1", cell.Reference.Address);
        Assert.Equal(sheetName, merge.SheetName);
        Assert.Equal("B2:C3", merge.Range.Address);
        Assert.Empty(layout.ListRegions);
    }

    /// <summary>
    /// 表示工作表名称校验测试使用的最小实体。
    /// </summary>
    private sealed class SheetNameDocument
    {
        /// <summary>
        /// 初始化一个 <see cref="SheetNameDocument" /> 类型的实例。
        /// </summary>
        public SheetNameDocument()
        {
        }

        /// <summary>
        /// 获取或设置固定单元格测试值。
        /// </summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>
        /// 获取或设置列表区域测试行。
        /// </summary>
        public List<SheetNameLine> Lines { get; set; } = new List<SheetNameLine>();
    }

    /// <summary>
    /// 表示工作表名称校验测试使用的最小列表项。
    /// </summary>
    private sealed class SheetNameLine
    {
        /// <summary>
        /// 初始化一个 <see cref="SheetNameLine" /> 类型的实例。
        /// </summary>
        public SheetNameLine()
        {
        }

        /// <summary>
        /// 获取或设置列表项测试值。
        /// </summary>
        public string Value { get; set; } = string.Empty;
    }
}
