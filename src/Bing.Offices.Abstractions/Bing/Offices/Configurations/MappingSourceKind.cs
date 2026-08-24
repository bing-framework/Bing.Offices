namespace Bing.Offices.Configurations;

/// <summary>
/// 映射配置来源。
/// </summary>
public enum MappingSourceKind
{
    /// <summary>
    /// 约定或默认值。
    /// </summary>
    Convention,

    /// <summary>
    /// 实体特性。
    /// </summary>
    Attribute,

    /// <summary>
    /// Profile Fluent 配置。
    /// </summary>
    Profile,

    /// <summary>
    /// JSON/XML 文档配置。
    /// </summary>
    Document,

    /// <summary>
    /// 请求级 Fluent 配置。
    /// </summary>
    Request
}
