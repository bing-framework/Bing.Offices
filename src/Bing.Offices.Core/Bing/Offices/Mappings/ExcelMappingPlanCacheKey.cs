using System.Security.Cryptography;
using System.Text.Json;
using Bing.Offices.Configurations;

namespace Bing.Offices.Mappings;

/// <summary>创建隔离映射计划缓存使用的稳定键。</summary>
internal static class ExcelMappingPlanCacheKey
{
    /// <summary>根据模型、方向、租户和规范化配置创建 SHA-256 Base64 缓存键。</summary>
    internal static string Create<T>(ExcelMappingDocument document, MappingDirection direction,
        ExcelMappingConfiguration configuration) where T : class, new()
    {
        var payload = JsonSerializer.SerializeToUtf8Bytes(new
        {
            document.TenantId,
            ModelType = typeof(T).AssemblyQualifiedName,
            Direction = direction,
            document.ConfigurationVersion,
            Configuration = configuration
        }, new JsonSerializerOptions { IgnoreNullValues = false });
        using var sha256 = SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(payload));
    }
}
