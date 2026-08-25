namespace Bing.Offices.Configurations;

/// <summary>
/// Mapping Profile 注册表。
/// </summary>
public interface IMappingProfileRegistry : IMappingProfileResolver
{
    /// <summary>
    /// 注册一个单方向 Profile 描述。
    /// </summary>
    /// <param name="descriptor">Profile 描述。</param>
    void Register(ProfileDescriptor descriptor);
}
