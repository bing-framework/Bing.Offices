using System;
using System.Linq.Expressions;
using System.Text.Json;
using Bing.Offices.Configurations;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 导入和导出映射构建器的直接属性合同及失败状态测试。
/// </summary>
public class MappingBuilderPropertyTest
{
    /// <summary>
    /// 测试 - 导入配置必须拒绝其他对象和静态属性，且失败后保留完整配置。
    /// </summary>
    [Theory]
    [InlineData("nested", false)]
    [InlineData("nested", true)]
    [InlineData("captured", false)]
    [InlineData("captured", true)]
    [InlineData("new", false)]
    [InlineData("new", true)]
    [InlineData("static", false)]
    [InlineData("static", true)]
    public void ImportProperty_InvalidReceiver_ShouldRejectWithoutChangingConfiguration(string kind, bool configured)
    {
        var builder = new ImportMappingBuilder<Row>();
        if (configured)
            builder.Property(row => row.Code).HasHeader("原名称").HasAlias("原别名");
        var before = JsonSerializer.Serialize(builder.Build());

        var exception = Assert.Throws<ArgumentException>(() =>
            builder.Property(InvalidExpression(kind)).HasHeader("错误覆盖"));

        Assert.Equal("expression", exception.ParamName);
        Assert.Equal(before, JsonSerializer.Serialize(builder.Build()));
    }

    /// <summary>
    /// 测试 - 导出配置必须拒绝其他对象和静态属性，且失败后保留完整配置。
    /// </summary>
    [Theory]
    [InlineData("nested", false)]
    [InlineData("nested", true)]
    [InlineData("captured", false)]
    [InlineData("captured", true)]
    [InlineData("new", false)]
    [InlineData("new", true)]
    [InlineData("static", false)]
    [InlineData("static", true)]
    public void ExportProperty_InvalidReceiver_ShouldRejectWithoutChangingConfiguration(string kind, bool configured)
    {
        var builder = new ExportMappingBuilder<Row>();
        if (configured)
            builder.Property(row => row.Code).HasHeader("原名称").HasFormatter("@");
        var before = JsonSerializer.Serialize(builder.Build());

        var exception = Assert.Throws<ArgumentException>(() =>
            builder.Property(InvalidExpression(kind)).HasHeader("错误覆盖"));

        Assert.Equal("expression", exception.ParamName);
        Assert.Equal(before, JsonSerializer.Serialize(builder.Build()));
    }

    /// <summary>
    /// 测试 - 导入直接属性可以重复配置，构建结果与构建器保持快照隔离。
    /// </summary>
    [Fact]
    public void ImportProperty_DirectReceiver_ShouldConfigureAndBuildIndependentSnapshots()
    {
        var builder = new ImportMappingBuilder<Row>();
        builder.Property(row => row.Code).HasHeader("原名称");
        var first = builder.Build();

        builder.Property(row => row.Code).HasHeader("新名称");
        var second = builder.Build();

        var firstColumn = Assert.Single(first.Columns);
        var secondColumn = Assert.Single(second.Columns);
        Assert.Equal(nameof(Row.Code), firstColumn.PropertyName);
        Assert.Equal(nameof(Row.Code), secondColumn.PropertyName);
        Assert.Equal("原名称", firstColumn.Title);
        Assert.Equal("新名称", secondColumn.Title);
        secondColumn.Title = "外部修改";
        secondColumn.Aliases.Add("外部别名");
        var thirdColumn = Assert.Single(builder.Build().Columns);
        Assert.Equal("新名称", thirdColumn.Title);
        Assert.Empty(thirdColumn.Aliases);
        Assert.Equal("原名称", firstColumn.Title);
    }

    /// <summary>
    /// 测试 - 导出直接属性可以重复配置，构建结果与构建器保持快照隔离。
    /// </summary>
    [Fact]
    public void ExportProperty_DirectReceiver_ShouldConfigureAndBuildIndependentSnapshots()
    {
        var builder = new ExportMappingBuilder<Row>();
        builder.Property(row => row.Code).HasHeader("原名称");
        var first = builder.Build();

        builder.Property(row => row.Code).HasHeader("新名称");
        var second = builder.Build();

        var firstColumn = Assert.Single(first.Columns);
        var secondColumn = Assert.Single(second.Columns);
        Assert.Equal(nameof(Row.Code), firstColumn.PropertyName);
        Assert.Equal(nameof(Row.Code), secondColumn.PropertyName);
        Assert.Equal("原名称", firstColumn.Title);
        Assert.Equal("新名称", secondColumn.Title);
        secondColumn.Title = "外部修改";
        secondColumn.Formatter = "外部格式";
        var thirdColumn = Assert.Single(builder.Build().Columns);
        Assert.Equal("新名称", thirdColumn.Title);
        Assert.Null(thirdColumn.Formatter);
        Assert.Equal("原名称", firstColumn.Title);
    }

    /// <summary>测试 - 导入空表达式应立即拒绝且不生成列配置。</summary>
    [Fact]
    public void ImportProperty_NullExpression_ShouldRejectWithoutAddingColumn()
    {
        var builder = new ImportMappingBuilder<Row>();

        var exception = Assert.Throws<ArgumentNullException>(() => builder.Property<string>(null));

        Assert.Equal("expression", exception.ParamName);
        Assert.Empty(builder.Build().Columns);
    }

    /// <summary>测试 - 导出空表达式应立即拒绝且不生成列配置。</summary>
    [Fact]
    public void ExportProperty_NullExpression_ShouldRejectWithoutAddingColumn()
    {
        var builder = new ExportMappingBuilder<Row>();

        var exception = Assert.Throws<ArgumentNullException>(() => builder.Property<string>(null));

        Assert.Equal("expression", exception.ParamName);
        Assert.Empty(builder.Build().Columns);
    }

    /// <summary>创建末端属性声明类型相同、但接收者不是 lambda 参数的表达式。</summary>
    private static Expression<Func<Row, string>> InvalidExpression(string kind)
    {
        var captured = new Row();
        return kind switch
        {
            "nested" => row => row.Child.Code,
            "captured" => row => captured.Code,
            "new" => row => new Row().Code,
            "static" => row => Row.StaticCode,
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
    }

    private sealed class Row
    {
        public string Code { get; set; }

        public Row Child { get; set; }

        public static string StaticCode { get; set; }
    }
}
