using System;

namespace Bing.Offices.Configurations;

/// <summary>
/// 单一方向的规范化 Profile 描述。
/// </summary>
public sealed class ProfileDescriptor
{
    /// <summary>
    /// 初始化 Profile 描述。
    /// </summary>
    /// <param name="name">Profile 名称。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="modelType">方向对应的模型类型。</param>
    /// <param name="configuration">已构建的方向配置。</param>
    /// <param name="profileType">Profile 实现类型。</param>
    public ProfileDescriptor(string name, MappingDirection direction, Type modelType,
        ExcelMappingConfiguration configuration, Type profileType = null)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Profile 名称不能为空。", nameof(name));
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));
        ModelType = modelType ?? throw new ArgumentNullException(nameof(modelType));
        ConfigurationSnapshot = MappingConfigurationCloner.Clone(configuration ??
            throw new ArgumentNullException(nameof(configuration)), MappingSourceKind.Profile);
        Name = name;
        Direction = direction;
        ProfileType = profileType;
    }

    /// <summary>
    /// 获取 Profile 名称。
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// 获取映射方向。
    /// </summary>
    public MappingDirection Direction { get; }

    /// <summary>
    /// 获取方向对应的模型类型。
    /// </summary>
    public Type ModelType { get; }

    /// <summary>
    /// 获取配置快照。
    /// </summary>
    public ExcelMappingConfiguration Configuration =>
        MappingConfigurationCloner.Clone(ConfigurationSnapshot, MappingSourceKind.Profile);

    /// <summary>
    /// 获取 Profile 实现类型。
    /// </summary>
    public Type ProfileType { get; }

    private ExcelMappingConfiguration ConfigurationSnapshot { get; }
}
