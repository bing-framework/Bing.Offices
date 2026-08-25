using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 配置 Patch 操作语义测试。
/// </summary>
public sealed class MappingConfigurationPatchTest
{
    /// <summary>
    /// 测试 - 显式 reset 应清除低优先级标量而普通 null 应保持原值。
    /// </summary>
    [Fact]
    public void Merge_ExplicitReset_ShouldClearLowerScalar()
    {
        // Arrange
        var lower = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new()
                {
                    PropertyName = "Name",
                    Title = "名称",
                    ColumnIndex = 2,
                    Formatter = "@",
                    Ignored = true,
                    DecimalScale = 2,
                    ConverterName = "text",
                    ImportWhitespace = ExcelWhitespacePolicy.Trim,
                    ImageMultiplicity = ExcelImageMultiplicityPolicy.First
                }
            }
        };
        var higher = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new()
                {
                    PropertyName = "Name",
                    ResetColumnIndex = true,
                    ClearFormatter = true,
                    ResetIgnored = true,
                    ResetDecimalScale = true,
                    ClearConverterName = true,
                    ResetImportWhitespace = true,
                    ResetImageMultiplicity = true
                }
            }
        };

        // Act
        var merged = MappingConfigurationMerger.Merge(lower, higher, MappingSourceKind.Request);

        // Assert
        var column = Assert.Single(merged.Columns);
        Assert.Null(column.ColumnIndex);
        Assert.Null(column.Formatter);
        Assert.Null(column.Ignored);
        Assert.Null(column.DecimalScale);
        Assert.Null(column.ConverterName);
        Assert.Null(column.ImportWhitespace);
        Assert.Null(column.ImageMultiplicity);
        Assert.Equal("名称", column.Title);
    }

    /// <summary>
    /// 测试 - clear 操作应清空低优先级别名、值映射和动态列。
    /// </summary>
    [Fact]
    public void Merge_ExplicitClear_ShouldRemoveLowerCollections()
    {
        // Arrange
        var lower = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "region", Title = "区域" }
            },
            Columns = new List<ExcelColumnConfiguration>
            {
                new()
                {
                    PropertyName = "Name",
                    Aliases = new List<string> { "旧名称" },
                    ValueMappings = new List<ExcelValueMappingConfiguration>
                    {
                        new() { Text = "有效", Value = "1" }
                    }
                }
            }
        };
        var higher = new ExcelMappingConfiguration
        {
            ClearDynamicColumns = true,
            Columns = new List<ExcelColumnConfiguration>
            {
                new()
                {
                    PropertyName = "Name",
                    ClearAliases = true,
                    ClearValueMappings = true
                }
            }
        };

        // Act
        var merged = MappingConfigurationMerger.Merge(lower, higher, MappingSourceKind.Request);

        // Assert
        Assert.Empty(merged.DynamicColumns);
        var column = Assert.Single(merged.Columns);
        Assert.Empty(column.Aliases);
        Assert.Empty(column.ValueMappings);
    }

    /// <summary>
    /// 测试 - reset style/layout 应移除低优先级对象，而新的高优先级对象仍可覆盖。
    /// </summary>
    [Fact]
    public void Merge_ResetStyleAndLayout_ShouldClearLowerObjects()
    {
        // Arrange
        var lower = new ExcelMappingConfiguration
        {
            Style = new ExcelMappingStyleConfiguration { HeaderStyleKey = "header" },
            Layout = new ExcelMappingLayoutConfiguration { PlacementKey = "after-code" }
        };
        var higher = new ExcelMappingConfiguration { ResetStyle = true, ResetLayout = true };

        // Act
        var merged = MappingConfigurationMerger.Merge(lower, higher, MappingSourceKind.Document);

        // Assert
        Assert.Null(merged.Style);
        Assert.Null(merged.Layout);
    }

    /// <summary>
    /// 测试 - 动态列 append 应按稳定 Key 更新已有项并追加新项，remove 应只删除指定项。
    /// </summary>
    [Fact]
    public void Merge_DynamicColumnsAppendAndRemove_ShouldUseStableKeys()
    {
        // Arrange
        var lower = new ExcelMappingConfiguration
        {
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "region", Title = "区域" },
                new() { Key = "legacy", Title = "旧列" }
            }
        };
        var higher = new ExcelMappingConfiguration
        {
            DynamicColumnMergeMode = ExcelDynamicColumnMergeMode.Append,
            DynamicColumnKeysToRemove = new List<string> { "legacy" },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "REGION", Title = "新区域" },
                new() { Key = "amount", Title = "金额" }
            }
        };

        // Act
        var merged = MappingConfigurationMerger.Merge(lower, higher, MappingSourceKind.Document);

        // Assert
        Assert.Equal(new[] { "新区域", "金额" },
            merged.DynamicColumns.ConvertAll(column => column.Title));
        Assert.Null(merged.DynamicColumnMergeMode);
        Assert.Empty(merged.DynamicColumnKeysToRemove);
    }

    /// <summary>
    /// 测试 - 样式字段清除和布局字段 reset 不应影响未被操作的字段。
    /// </summary>
    [Fact]
    public void Merge_StyleAndLayoutFieldPatch_ShouldPreserveUntouchedValues()
    {
        // Arrange
        var lower = new ExcelMappingConfiguration
        {
            Style = new ExcelMappingStyleConfiguration
            {
                HeaderStyleKey = "header",
                BodyStyleKey = "body",
                NumberFormat = "0.00"
            },
            Layout = new ExcelMappingLayoutConfiguration
            {
                ColumnIndex = 3,
                PlacementKey = "after-code"
            }
        };
        var higher = new ExcelMappingConfiguration
        {
            Style = new ExcelMappingStyleConfiguration
            {
                ClearHeaderStyleKey = true,
                NumberFormat = "0"
            },
            Layout = new ExcelMappingLayoutConfiguration
            {
                ResetColumnIndex = true
            }
        };

        // Act
        var merged = MappingConfigurationMerger.Merge(lower, higher, MappingSourceKind.Request);

        // Assert
        Assert.Null(merged.Style.HeaderStyleKey);
        Assert.Equal("body", merged.Style.BodyStyleKey);
        Assert.Equal("0", merged.Style.NumberFormat);
        Assert.Null(merged.Layout.ColumnIndex);
        Assert.Equal("after-code", merged.Layout.PlacementKey);
    }

    /// <summary>
    /// 测试 - v2 JSON/XML 应保留动态列、样式和布局 Patch 字段。
    /// </summary>
    [Fact]
    public void Loader_V2PatchFields_ShouldRoundTripAcrossJsonAndXml()
    {
        // Arrange
        const string json = "{\"version\":2,\"import\":{\"dynamicColumnKeysToRemove\":[\"legacy\"],\"dynamicColumnMergeMode\":1,\"style\":{\"clearHeaderStyleKey\":true},\"layout\":{\"resetColumnIndex\":true}}}";

        // Act
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var xml = ExcelMappingConfigurationLoader.ToXml(document);
        var roundTrip = ExcelMappingConfigurationLoader.FromXmlDocument(xml);

        // Assert
        Assert.Equal(new[] { "legacy" }, roundTrip.Import.DynamicColumnKeysToRemove);
        Assert.Equal(ExcelDynamicColumnMergeMode.Append, roundTrip.Import.DynamicColumnMergeMode);
        Assert.True(roundTrip.Import.Style.ClearHeaderStyleKey);
        Assert.True(roundTrip.Import.Layout.ResetColumnIndex);
        Assert.Null(roundTrip.Export);
    }

    /// <summary>
    /// 测试 - 文档级 clear/reset 应在 Profile 之后进入最终 Plan，而不是退化为未设置。
    /// </summary>
    [Fact]
    public void Plan_DocumentPatch_ShouldClearProfileColumnValues()
    {
        // Arrange
        var registry = new MappingProfileRegistry();
        registry.Register(new ProfileDescriptor("patch-profile", MappingDirection.Import, typeof(PatchRow),
            new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new()
                    {
                        PropertyName = nameof(PatchRow.Name),
                        Title = "Profile 标题",
                        Aliases = new List<string> { "旧名称" },
                        Formatter = "@",
                        Ignored = true,
                        ValueMappings = new List<ExcelValueMappingConfiguration>
                        {
                            new() { Text = "有效", Value = "yes" }
                        },
                        ImageMultiplicity = ExcelImageMultiplicityPolicy.All
                    }
                }
            }));
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Profile = "patch-profile",
                Columns = new List<ExcelColumnConfiguration>
                {
                    new()
                    {
                        PropertyName = nameof(PatchRow.Name),
                        ClearTitle = true,
                        ClearAliases = true,
                        ClearFormatter = true,
                        ResetIgnored = true,
                        ClearValueMappings = true,
                        ResetImageMultiplicity = true
                    }
                }
            }
        };
        var factory = new ExcelMappingPlanFactory(profileRegistry: registry);

        // Act
        var column = Assert.Single(factory.Create<PatchRow>(document, MappingDirection.Import).Columns);

        // Assert
        Assert.Null(column.Title);
        Assert.Empty(column.Aliases);
        Assert.Null(column.Formatter);
        Assert.False(column.Ignored);
        Assert.Empty(column.ValueMap);
        Assert.Equal(ExcelImageMultiplicityPolicy.First, column.ImageMultiplicity);
    }

    /// <summary>
    /// 测试 - 显式文档缺少目标方向时默认失败，避免静默回退到约定映射。
    /// </summary>
    [Fact]
    public void Plan_MissingDirection_ShouldFailWithoutExplicitFallback()
    {
        // Arrange
        var factory = new ExcelMappingPlanFactory();
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration()
        };

        // Act
        var exception = Assert.Throws<InvalidOperationException>(() =>
            factory.Create<PatchRow>(document, MappingDirection.Export));

        // Assert
        Assert.Contains("UseConventionFallback", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Export", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 显式启用 Convention fallback 后，缺失方向应使用实体约定映射。
    /// </summary>
    [Fact]
    public void Plan_MissingDirection_WithExplicitFallback_ShouldUseConvention()
    {
        // Arrange
        var factory = new ExcelMappingPlanFactory();
        var document = new ExcelMappingDocument
        {
            UseConventionFallback = true,
            Import = new ExcelMappingConfiguration()
        };

        // Act
        var plan = factory.Create<PatchRow>(document, MappingDirection.Export);

        // Assert
        Assert.NotNull(plan.Columns.Single(column => column.Name == nameof(PatchRow.Name)));
    }

    /// <summary>
    /// 测试 - Attribute 层的标题、格式、忽略和值映射也应支持文档级 clear/reset。
    /// </summary>
    [Fact]
    public void Plan_DocumentPatch_ShouldClearAttributeColumnValues()
    {
        // Arrange
        var document = new ExcelMappingDocument
        {
            Import = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new()
                    {
                        PropertyName = nameof(AttributePatchRow.Name),
                        ClearTitle = true,
                        ClearFormatter = true,
                        ResetIgnored = true,
                        ClearValueMappings = true
                    }
                }
            }
        };
        var factory = new ExcelMappingPlanFactory();

        // Act
        var column = Assert.Single(factory.Create<AttributePatchRow>(document, MappingDirection.Import).Columns);

        // Assert
        Assert.Null(column.Title);
        Assert.Null(column.Formatter);
        Assert.False(column.Ignored);
        Assert.Empty(column.ValueMap);
    }

    /// <summary>
    /// 测试 - JSON 中未定义的校验和值映射合并枚举值应被拒绝。
    /// </summary>
    [Fact]
    public void Loader_InvalidColumnMergeEnums_ShouldRejectUndefinedValues()
    {
        // Arrange
        const string validation = "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Name\",\"validationRuleMergeMode\":99}]}}";
        const string valueMapping = "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Name\",\"valueMappingMergeMode\":99}]}}";

        // Act / Assert
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(validation));
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(valueMapping));
    }

    private sealed class PatchRow
    {
        public string Name { get; set; }
    }

    private sealed class AttributePatchRow
    {
        [ColumnName("属性名称")]
        [DataFormat("@")]
        [ExcelIgnore]
        [ValueMapping("有效", "yes")]
        public string Name { get; set; }
    }
}
