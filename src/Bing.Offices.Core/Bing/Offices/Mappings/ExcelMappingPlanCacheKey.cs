using System.Security.Cryptography;
using System.Text.Json;
using Bing.Offices.Configurations;

namespace Bing.Offices.Mappings;

/// <summary>创建隔离映射计划缓存使用的稳定键。</summary>
internal static class ExcelMappingPlanCacheKey
{
    /// <summary>缓存键序列化选项，保留 null 字段以维持键隔离；创建后只读并可在线程间复用。</summary>
    private static readonly JsonSerializerOptions SerializerOptions = new JsonSerializerOptions
    {
        IgnoreNullValues = false
    };

    /// <summary>根据模型、方向、租户和规范化配置创建 SHA-256 Base64 缓存键。</summary>
    /// <typeparam name="T">参与缓存隔离的实体类型。</typeparam>
    /// <param name="document">包含租户和版本信息的规范化映射文档。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="configuration">请求级映射配置。</param>
    /// <returns>由映射上下文稳定计算出的 Base64 编码 SHA-256 键。</returns>
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
        }, SerializerOptions);
        using var sha256 = SHA256.Create();
        return Convert.ToBase64String(sha256.ComputeHash(payload));
    }
}
