namespace Bing.Offices.Configurations;

/// <summary>
/// Mapping Profile 注册表。
/// </summary>
public interface IMappingProfileRegistry
{
    /// <summary>
    /// 注册一个单方向 Profile 描述。
    /// </summary>
    /// <param name="descriptor">Profile 描述。</param>
    void Register(ProfileDescriptor descriptor);

    /// <summary>
    /// 尝试获取指定名称、方向和模型的单方向 Profile 描述。
    /// </summary>
    /// <param name="profileName">Profile 名称。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="modelType">模型类型。</param>
    /// <param name="descriptor">匹配的描述。</param>
    bool TryGetDescriptor(string profileName, MappingDirection direction, Type modelType,
        out ProfileDescriptor descriptor);

}
