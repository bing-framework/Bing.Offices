using System.Linq;
using Bing.Offices.Configurations;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Mapping Profile v2 契约与不可变快照测试。
/// </summary>
public class MappingProfileV2Test
{
    /// <summary>
    /// 测试 - 双模型 Profile 应分别编译导入和导出模型的映射配置。
    /// </summary>
    [Fact]
    public void DirectionalProfile_DifferentModels_ShouldBuildIndependentPlans()
    {
        // Arrange
        var import = new ImportMappingBuilder<ImportModel>();
        import.Property(model => model.Name)
            .HasHeader("导入名称")
            .HasAlias("旧名称")
            .HasImageMultiplicity(ExcelImageMultiplicityPolicy.All);
        var export = new ExportMappingBuilder<ExportModel>();
        export.Property(model => model.Label)
            .HasHeader("导出标签")
            .HasFormatter("@");
        var document = new ExcelMappingDocument { Import = import.Build(), Export = export.Build() };

        // Act
        var importMap = ExcelTypeMapFactory.Get<ImportModel>(document, MappingDirection.Import);
        var exportMap = ExcelTypeMapFactory.Get<ExportModel>(document, MappingDirection.Export);

        // Assert
        var importProperty = Assert.Single(importMap.Properties.Where(property => property.Name == nameof(ImportModel.Name)));
        var exportProperty = Assert.Single(exportMap.Properties.Where(property => property.Name == nameof(ExportModel.Label)));
        Assert.Equal("导入名称", importProperty.Title);
        Assert.Contains("旧名称", importProperty.Aliases);
        Assert.Equal(ExcelImageMultiplicityPolicy.All, importProperty.ImageMultiplicity);
        Assert.Equal("导出标签", exportProperty.Title);
        Assert.Equal("@", exportProperty.Formatter);
    }

    /// <summary>
    /// 测试 - 同模型 Profile 应通过方向接口复用同一模型并保持独立配置。
    /// </summary>
    [Fact]
    public void SameModelProfile_ShouldSupportImportAndExportDirections()
    {
        // Arrange
        var import = new ImportMappingBuilder<ImportModel>();
        import.Property(model => model.Name).HasHeader("导入名称");
        var export = new ExportMappingBuilder<ImportModel>();
        export.Property(model => model.Name).HasHeader("导出名称");
        var document = new ExcelMappingDocument { Import = import.Build(), Export = export.Build() };

        // Act
        var importMap = ExcelTypeMapFactory.Get<ImportModel>(document, MappingDirection.Import);
        var exportMap = ExcelTypeMapFactory.Get<ImportModel>(document, MappingDirection.Export);

        // Assert
        Assert.Equal("导入名称", importMap.Properties.Single(property => property.Name == nameof(ImportModel.Name)).Title);
        Assert.Equal("导出名称", exportMap.Properties.Single(property => property.Name == nameof(ImportModel.Name)).Title);
    }

    /// <summary>
    /// 测试 - Profile 构建完成后，读取配置或修改输入副本不应改变已构建结果。
    /// </summary>
    [Fact]
    public void DirectionalProfile_Snapshot_ShouldBeImmutable()
    {
        // Arrange
        var builder = new ImportMappingBuilder<ImportModel>();
        builder.Property(model => model.Name).HasHeader("稳定名称");
        var first = builder.Build();

        // Act
        first.Columns[0].Title = "外部修改";
        first.Columns[0].Aliases.Add("外部别名");
        var second = builder.Build();

        // Assert
        Assert.Equal("稳定名称", second.Columns[0].Title);
        Assert.Empty(second.Columns[0].Aliases);
        var document = new ExcelMappingDocument { Import = second };
        Assert.Equal("稳定名称", ExcelTypeMapFactory.Get<ImportModel>(document,
            MappingDirection.Import).Properties.Single(property => property.Name == nameof(ImportModel.Name)).Title);
    }

    /// <summary>
    /// 测试 - 方向专用构建器不得暴露另一方向的设置方法。
    /// </summary>
    [Fact]
    public void DirectionBuilders_ShouldKeepImportAndExportApiSeparate()
    {
        // Arrange
        var importMethods = typeof(ImportColumnMappingBuilder<ImportModel, string>).GetMethods()
            .Select(method => method.Name).ToArray();
        var exportMethods = typeof(ExportColumnMappingBuilder<ExportModel, string>).GetMethods()
            .Select(method => method.Name).ToArray();

        // Assert
        Assert.Contains(nameof(ImportColumnMappingBuilder<ImportModel, string>.HasWhitespace), importMethods);
        Assert.Contains(nameof(ImportColumnMappingBuilder<ImportModel, string>.HasValidationRule), importMethods);
        Assert.DoesNotContain(nameof(ExportColumnMappingBuilder<ExportModel, string>.HasFormatter), importMethods);
        Assert.Contains(nameof(ExportColumnMappingBuilder<ExportModel, string>.HasFormatter), exportMethods);
        Assert.DoesNotContain(nameof(ImportColumnMappingBuilder<ImportModel, string>.HasValidationRule), exportMethods);
    }

    /// <summary>
    /// 测试 - Remove/Clear 与 Append/Replace 应在配置编译时保留明确语义。
    /// </summary>
    [Fact]
    public void MappingConfiguration_MergeOperations_ShouldApplyExplicitly()
    {
        // Arrange
        var configuration = ExcelMapping.For<ImportModel>()
            .Property(model => model.Name)
            .HasValidationRule("legacy")
            .And()
            .Build();
        configuration.Columns[0].ValidationRuleNamesToRemove.Add("legacy");
        configuration.Columns[0].ValidationRuleNames.Add("current");
        configuration.Columns[0].ValidationRuleMergeMode = ExcelValidationRuleMergeMode.Append;

        var appendConfiguration = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(ImportModel.Name),
                    ValueMappings = { new ExcelValueMappingConfiguration { Text = "是", Value = "true" } },
                    ValueMappingMergeMode = ExcelValueMappingMergeMode.Append
                }
            }
        };

        // Act
        var map = ExcelTypeMapFactory.Get<ImportModel>(configuration);
        var appendMap = ExcelTypeMapFactory.Get<ImportModel>(appendConfiguration);

        // Assert
        var property = map.Properties.Single(item => item.Name == nameof(ImportModel.Name));
        Assert.DoesNotContain("legacy", property.ValidationRuleNames);
        Assert.Contains("current", property.ValidationRuleNames);
        Assert.Contains("是", appendMap.Properties.Single(item => item.Name == nameof(ImportModel.Name)).ValueMap.Keys);
    }

    private sealed class ImportModel
    {
        public string Name { get; set; }
    }

    private sealed class ExportModel
    {
        public string Label { get; set; }
    }
}
