using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 映射计划缓存键测试。
/// </summary>
public sealed class ExcelMappingPlanCacheKeyTest
{
    /// <summary>
    /// 测试 - 复用序列化选项后，缓存键仍与原 UTF-8 序列化合同完全等价。
    /// </summary>
    [Fact]
    public void Create_ShouldMatchPreviousUtf8SerializationContract()
    {
        // Arrange
        var document = CreateDocument("tenant-a", "v1");
        var configuration = CreateConfiguration("编码");

        // Act
        var actual = ExcelMappingPlanCacheKey.Create<CacheKeyRow>(document,
            MappingDirection.Import, configuration);
        var expectedPayload = JsonSerializer.SerializeToUtf8Bytes(new
        {
            document.TenantId,
            ModelType = typeof(CacheKeyRow).AssemblyQualifiedName,
            Direction = MappingDirection.Import,
            document.ConfigurationVersion,
            Configuration = configuration
        }, new JsonSerializerOptions { DefaultIgnoreCondition = JsonIgnoreCondition.Never });
        using var sha256 = SHA256.Create();
        var expected = Convert.ToBase64String(sha256.ComputeHash(expectedPayload));

        // Assert
        Assert.Equal(expected, actual);
        Assert.Equal(actual, ExcelMappingPlanCacheKey.Create<CacheKeyRow>(document,
            MappingDirection.Import, configuration));
    }

    /// <summary>
    /// 测试 - 租户、模型类型、方向、配置版本和配置内容均保持缓存隔离。
    /// </summary>
    [Fact]
    public void Create_ShouldIsolateEveryCacheIdentityField()
    {
        // Arrange
        var document = CreateDocument("tenant-a", "v1");
        var configuration = CreateConfiguration("编码");
        var baseline = ExcelMappingPlanCacheKey.Create<CacheKeyRow>(document,
            MappingDirection.Import, configuration);

        // Act / Assert
        Assert.NotEqual(baseline, ExcelMappingPlanCacheKey.Create<CacheKeyRow>(
            CreateDocument("tenant-b", "v1"), MappingDirection.Import, configuration));
        Assert.NotEqual(baseline, ExcelMappingPlanCacheKey.Create<OtherCacheKeyRow>(
            document, MappingDirection.Import, configuration));
        Assert.NotEqual(baseline, ExcelMappingPlanCacheKey.Create<CacheKeyRow>(
            document, MappingDirection.Export, configuration));
        Assert.NotEqual(baseline, ExcelMappingPlanCacheKey.Create<CacheKeyRow>(
            CreateDocument("tenant-a", "v2"), MappingDirection.Import, configuration));
        Assert.NotEqual(baseline, ExcelMappingPlanCacheKey.Create<CacheKeyRow>(
            document, MappingDirection.Import, CreateConfiguration("金额")));
    }

    /// <summary>
    /// 创建测试映射文档。
    /// </summary>
    /// <param name="tenantId">租户标识。</param>
    /// <param name="configurationVersion">配置版本。</param>
    /// <returns>包含指定租户、配置版本和导入列映射的文档。</returns>
    private static ExcelMappingDocument CreateDocument(string tenantId, string configurationVersion) =>
        new ExcelMappingDocument
        {
            TenantId = tenantId,
            ConfigurationVersion = configurationVersion,
            Import = CreateConfiguration("编码")
        };

    /// <summary>
    /// 创建测试映射配置。
    /// </summary>
    /// <param name="title">配置标题。</param>
    /// <returns>将 Code 属性映射到指定列标题的订单配置。</returns>
    private static ExcelMappingConfiguration CreateConfiguration(string title) =>
        new ExcelMappingConfiguration
        {
            Profile = "orders",
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(CacheKeyRow.Code),
                    Title = title
                }
            }
        };

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class CacheKeyRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class OtherCacheKeyRow
    {
        /// <summary>
        /// 获取或设置编码。
        /// </summary>
        public string Code { get; set; } = string.Empty;
    }
}
