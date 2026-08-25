namespace Bing.Offices.Configurations;

/// <summary>
/// 只读的 Mapping Profile 解析契约，供计划编译和业务消费方使用。
/// </summary>
public interface IMappingProfileResolver
{
    /// <summary>
    /// 尝试获取指定名称、方向和模型的单方向 Profile 描述。
    /// </summary>
    /// <param name="profileName">Profile 名称或稳定别名。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="modelType">模型类型。</param>
    /// <param name="descriptor">匹配的描述。</param>
    /// <returns>找到匹配描述时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    bool TryGetDescriptor(string profileName, MappingDirection direction, Type modelType,
        out ProfileDescriptor descriptor);
}
