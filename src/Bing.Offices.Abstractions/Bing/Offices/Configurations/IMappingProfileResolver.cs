namespace Bing.Offices.Configurations;

/// <summary>
/// 按 Profile 名称、映射方向和模型类型解析单方向描述。
/// </summary>
public interface IMappingProfileResolver
{
    /// <summary>
    /// 尝试获取指定名称、方向和模型的单方向 Profile 描述。
    /// </summary>
    /// <param name="profileName">Profile 名称或稳定别名；不可为空。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="modelType">待匹配的模型类型；不可为空。</param>
    /// <param name="descriptor">匹配的 Profile 描述；未找到时为 <see langword="null" />。</param>
    /// <returns>找到匹配描述时返回 <see langword="true" />，否则返回 <see langword="false" />。</returns>
    bool TryGetDescriptor(string profileName, MappingDirection direction, Type modelType,
        out ProfileDescriptor descriptor);
}
