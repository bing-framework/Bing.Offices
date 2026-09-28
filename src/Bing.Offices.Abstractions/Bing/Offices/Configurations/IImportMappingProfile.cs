namespace Bing.Offices.Configurations;

/// <summary>
/// 仅配置导入方向映射的 Profile。
/// </summary>
/// <typeparam name="TImport">导入模型类型。</typeparam>
public interface IImportMappingProfile<TImport>
    where TImport : class, new()
{
    /// <summary>
    /// 配置导入方向映射。
    /// </summary>
    /// <param name="setting">供 Profile 配置导入映射的 Fluent 设置。</param>
    void Configure(ImportMappingBuilder<TImport> setting);
}
