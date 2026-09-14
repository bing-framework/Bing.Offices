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

    private static ExcelMappingDocument CreateDocument(string tenantId, string configurationVersion) =>
        new ExcelMappingDocument
        {
            TenantId = tenantId,
            ConfigurationVersion = configurationVersion,
            Import = CreateConfiguration("编码")
        };

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

    private sealed class CacheKeyRow
    {
        public string Code { get; set; } = string.Empty;
    }

    private sealed class OtherCacheKeyRow
    {
        public string Code { get; set; } = string.Empty;
    }
}
