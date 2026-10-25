using System.Collections.Generic;
using System.IO;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>验证 CSV 表头查找顺序及每次导入之间的状态隔离。</summary>
public sealed class CsvHeaderLookupTest
{
    /// <summary>跨属性标题、别名和属性名冲突时，较早映射仍应获胜。</summary>
    [Theory]
    [InlineData("Second")]
    [InlineData("FirstTitle")]
    [InlineData("ThirdTitle")]
    public void Import_ConflictingHeaderTokens_ShouldUseMappingOrder(string header)
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                Column(nameof(TokenConflictRow.First), "FirstTitle", 0, "Second", "ThirdTitle"),
                Column(nameof(TokenConflictRow.Second), "SecondTitle", 1, "FirstTitle"),
                Column(nameof(TokenConflictRow.Third), "ThirdTitle", 2)
            }
        };
        var result = Import<TokenConflictRow>(new CsvEntityImporter(),
            header + "\r\nselected\r\n", new CsvImportOptions<TokenConflictRow>
            {
                MappingConfiguration = mapping,
                RequireExpectedHeaders = false
            });

        AssertSuccess(result);
        var item = Assert.Single(result.Items);
        Assert.Equal("selected", item.First);
        Assert.Null(item.Second);
        Assert.Null(item.Third);
    }

    /// <summary>固定列与动态列使用相同表头时，应把值绑定到固定属性。</summary>
    [Fact]
    public void Import_FixedAndDynamicTokenConflict_ShouldPreferFixedProperty()
    {
        var options = new CsvImportOptions<FixedAndDynamicRow>
        {
            MappingConfiguration = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    Column(nameof(FixedAndDynamicRow.Name), "Shared", 0)
                },
                DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
                {
                    new() { Key = "shared", Title = "Shared", DataTypeName = "string" }
                }
            }
        };

        var result = Import<FixedAndDynamicRow>(new CsvEntityImporter(), "Shared\r\nfixed\r\n", options);

        AssertSuccess(result);
        var item = Assert.Single(result.Items);
        Assert.Equal("fixed", item.Name);
        Assert.Empty(item.Values);
    }

    /// <summary>同一导入器连续使用不同映射时，不应沿用前一请求的表头索引。</summary>
    [Fact]
    public void Import_SequentialMappings_ShouldNotReusePreviousHeaderTokens()
    {
        var importer = new CsvEntityImporter();
        var firstMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                Column(nameof(TwoTargetRow.First), "Token", 0),
                Column(nameof(TwoTargetRow.Second), "Other", 1)
            }
        };
        var secondMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                Column(nameof(TwoTargetRow.First), "Other", 0),
                Column(nameof(TwoTargetRow.Second), "Token", 1)
            }
        };

        var first = Import<TwoTargetRow>(importer, "Token\r\nfirst\r\n", new CsvImportOptions<TwoTargetRow>
        {
            MappingConfiguration = firstMapping,
            RequireExpectedHeaders = false
        });
        var second = Import<TwoTargetRow>(importer, "Token\r\nsecond\r\n", new CsvImportOptions<TwoTargetRow>
        {
            MappingConfiguration = secondMapping,
            RequireExpectedHeaders = false
        });

        AssertSuccess(first);
        Assert.Equal("first", Assert.Single(first.Items).First);
        Assert.Null(Assert.Single(first.Items).Second);
        AssertSuccess(second);
        Assert.Null(Assert.Single(second.Items).First);
        Assert.Equal("second", Assert.Single(second.Items).Second);
    }

    /// <summary>必需列缺失返回结构化错误后，同一映射仍可成功处理后续输入。</summary>
    [Fact]
    public void Import_MissingRequiredHeader_ShouldReturnInvalidHeader_AndAllowReuse()
    {
        var importer = new CsvEntityImporter();
        var options = new CsvImportOptions<RequiredHeadersRow>
        {
            MappingConfiguration = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    Column(nameof(RequiredHeadersRow.Name), "Name", 0),
                    Column(nameof(RequiredHeadersRow.Code), "Code", 1)
                }
            }
        };

        var failed = Import<RequiredHeadersRow>(importer, "Name\r\nfirst\r\n", options);
        AssertMissingHeader(failed, "CSV 不存在列: Code");

        var recovered = Import<RequiredHeadersRow>(importer, "Name,Code\r\nfirst,ready\r\n", options);
        AssertSuccess(recovered);
        var item = Assert.Single(recovered.Items);
        Assert.Equal("first", item.Name);
        Assert.Equal("ready", item.Code);
    }

    private static ExcelColumnConfiguration Column(string property, string title, int index,
        params string[] aliases) => new()
    {
        PropertyName = property,
        Title = title,
        ColumnIndex = index,
        Aliases = new List<string>(aliases)
    };

    private static CsvImportResult<T> Import<T>(CsvEntityImporter importer, string text,
        CsvImportOptions<T> options) where T : class, new()
    {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(text));
        var result = importer.Import(source, options);
        Assert.True(source.CanRead);
        return result;
    }

    private static void AssertSuccess<T>(CsvImportResult<T> result)
    {
        Assert.Empty(result.Errors);
        Assert.False(result.IsTruncated);
        Assert.Null(result.MaxErrors);
    }

    private static void AssertMissingHeader<T>(CsvImportResult<T> result, string message)
    {
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(CsvImportErrorCode.InvalidHeader, error.Code);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        Assert.Equal(message, error.Message);
        Assert.Null(error.FirstRowNumber);
        Assert.False(result.IsTruncated);
        Assert.Null(result.MaxErrors);
    }

    private sealed class TokenConflictRow
    {
        public string First { get; set; }
        public string Second { get; set; }
        public string Third { get; set; }
    }

    private sealed class FixedAndDynamicRow
    {
        public string Name { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
    }

    private sealed class TwoTargetRow
    {
        public string First { get; set; }
        public string Second { get; set; }
    }

    private sealed class RequiredHeadersRow
    {
        public string Name { get; set; }
        public string Code { get; set; }
    }
}
