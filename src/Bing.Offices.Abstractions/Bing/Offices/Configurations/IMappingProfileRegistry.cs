namespace Bing.Offices.Configurations;

/// <summary>
/// 注册并解析按模型类型、映射方向和名称隔离的 Profile 描述。
/// </summary>
public interface IMappingProfileRegistry : IMappingProfileResolver
{
    /// <summary>
    /// 注册一个单方向 Profile 描述。
    /// </summary>
    /// <param name="descriptor">待注册的单方向 Profile 描述；名称、方向和模型类型组合必须唯一。</param>
    void Register(ProfileDescriptor descriptor);
}
