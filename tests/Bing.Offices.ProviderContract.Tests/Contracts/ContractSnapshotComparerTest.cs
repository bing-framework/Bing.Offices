using System;
using System.Collections.Generic;
using System.Globalization;
using Bing.Offices.Imports;
using Bing.Offices.Testing.Models;
using Bing.Offices.Testing.Snapshots;
using Xunit;

namespace Bing.Offices.ProviderContract.Tests.Contracts;

/// <summary>
/// 验证 Provider 无关快照的排序、文化和空值边界。
/// </summary>
public sealed class ContractSnapshotComparerTest
{
    /// <summary>
    /// 比较器应区分顺序差异和序列长度差异。
    /// </summary>
    [Fact]
    public void Difference_ShouldDescribeOrderAndCountChanges()
    {
        var first = new[]
        {
            new MappingRowSnapshot("A", "Alice", 1),
            new MappingRowSnapshot("B", "Bob", 2)
        };
        var reordered = new[] { first[1], first[0] };

        Assert.False(ContractSnapshotComparer.Equal(first, reordered));
        Assert.Equal("count expected=2, actual=1",
            ContractSnapshotComparer.Difference(first, new[] { first[0] }));
        Assert.Equal("index=0, expected=MappingRowSnapshot { Code = A, Name = Alice, Quantity = 1 }, actual=MappingRowSnapshot { Code = B, Name = Bob, Quantity = 2 }",
            ContractSnapshotComparer.Difference(first, reordered));
    }

    /// <summary>
    /// 快照规范化应在 zh-CN 和 en-US 下保持不变。
    /// </summary>
    /// <param name="cultureName">用于验证快照区域性隔离的区域性名称。</param>
    [Theory]
    [InlineData("zh-CN")]
    [InlineData("en-US")]
    public void ScalarSnapshot_ShouldUseInvariantValues(string cultureName)
    {
        var previous = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
            var workbook = new ScalarContractWorkbook();
            workbook.Rows.Add(new ScalarContractRow
            {
                Code = "A",
                Count = 1,
                Amount = 12.5m,
                Enabled = true,
                Kind = ContractKind.Active,
                Date = new DateTime(2026, 9, 4),
                Optional = null,
                Values = new Dictionary<string, object>(StringComparer.Ordinal)
            });

            var actual = ContractSnapshots.ScalarRows(workbook);
            Assert.Equal("12.50", actual[0].Amount);
            Assert.Equal("2026-09-04", actual[0].Date);
            Assert.Equal("<null>", actual[0].Optional);
            Assert.Equal("<missing>", actual[0].Region);
        }
        finally
        {
            CultureInfo.CurrentCulture = previous;
        }
    }

    /// <summary>
    /// 错误快照应区分 null、空字符串和空白原始值。
    /// </summary>
    [Fact]
    public void ErrorSnapshot_ShouldPreserveNullEmptyAndBlankValues()
    {
        var nullValue = new ImportErrorSnapshot(ExcelImportErrorCode.Validation,
            "Data", 2, 1, null, null, "", null, null);
        var emptyValue = new ImportErrorSnapshot(ExcelImportErrorCode.Validation,
            "Data", 2, 1, "", "", "", null, "");
        var blankValue = new ImportErrorSnapshot(ExcelImportErrorCode.Validation,
            "Data", 2, 1, " ", " ", " ", null, " ");

        Assert.False(ContractSnapshotComparer.Equal(new[] { nullValue }, new[] { emptyValue }));
        Assert.False(ContractSnapshotComparer.Equal(new[] { emptyValue }, new[] { blankValue }));
        Assert.True(ContractSnapshotComparer.Equal(new[] { nullValue }, new[] { nullValue }));
    }
}
