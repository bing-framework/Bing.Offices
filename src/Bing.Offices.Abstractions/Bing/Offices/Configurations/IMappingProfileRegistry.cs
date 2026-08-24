namespace Bing.Offices.Configurations;

/// <summary>
/// Mapping Profile 注册表。
/// </summary>
public interface IMappingProfileRegistry
{
    /// <summary>
    /// 获取指定名称和模型类型的方向 Profile。
    /// </summary>
    ExcelMappingProfile<TImport, TExport> Get<TImport, TExport>(string profileName)
        where TImport : class, new()
        where TExport : class, new();

    /// <summary>
    /// 尝试获取指定名称和方向的 Profile 快照。
    /// </summary>
    bool TryGet(string profileName, MappingDirection direction, Type modelType,
        out IMappingProfileSnapshot snapshot);
}
