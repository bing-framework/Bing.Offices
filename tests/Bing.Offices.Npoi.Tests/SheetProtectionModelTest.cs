using Bing.Offices.Entities;
using Bing.Offices.Exports;
using System;
using System.Collections.Generic;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 工作表保护公共模型测试。
/// </summary>
public sealed class SheetProtectionModelTest
{
    /// <summary>
    /// 验证保护区域在创建和布局构建时隔离调用方数据。
    /// </summary>
    [Fact]
    public void Protection_ShouldSnapshotEditableRanges()
    {
        var ranges = new List<ExcelRangeDefinition>
        {
            new() { StartRow = 1, EndRow = 2, StartColumn = 0, EndColumn = 0 }
        };
        var definition = new ExcelSheetProtectionDefinition(ranges, "secret");
        ranges[0] = new ExcelRangeDefinition { StartRow = 9, EndRow = 9 };
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new Row() }, sheet => sheet.Protection(definition)));
        var layout = ExcelEntity.Layout<Row>(builder => builder
            .Cell("Data", "A1", row => row.Name)
            .Protection("Data", definition));

        Assert.Equal(1, definition.EditableRanges[0].StartRow);
        Assert.Equal(1, request.Sheets[0].Protection.EditableRanges[0].StartRow);
        Assert.Equal(1, layout.Protections["Data"].EditableRanges[0].StartRow);
        Assert.Equal("secret", request.Sheets[0].Protection.Password);
    }

    /// <summary>
    /// 验证非法区域和重复工作表保护被拒绝。
    /// </summary>
    [Fact]
    public void Protection_ShouldRejectInvalidDefinitions()
    {
        Assert.Throws<ArgumentNullException>(() => new ExcelSheetProtectionDefinition(null));
        Assert.Throws<ArgumentException>(() => new ExcelSheetProtectionDefinition(new[]
        {
            new ExcelRangeDefinition { StartRow = 3, EndRow = 1 }
        }));
        var definition = new ExcelSheetProtectionDefinition(Array.Empty<ExcelRangeDefinition>());
        Assert.Throws<ArgumentException>(() => ExcelEntity.Layout<Row>(builder => builder
            .Cell("Data", "A1", row => row.Name)
            .Protection("Data", definition)
            .Protection("data", definition)));
    }

    /// <summary>
    /// 测试行。
    /// </summary>
    public sealed class Row
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }
}
