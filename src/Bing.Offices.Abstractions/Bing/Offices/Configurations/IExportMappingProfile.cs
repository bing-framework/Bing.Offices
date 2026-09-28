namespace Bing.Offices.Configurations;

/// <summary>
/// 仅配置导出方向映射的 Profile。
/// </summary>
/// <typeparam name="TExport">导出模型类型。</typeparam>
public interface IExportMappingProfile<TExport>
    where TExport : class, new()
{
    /// <summary>
    /// 配置导出方向映射。
    /// </summary>
    /// <param name="setting">供 Profile 配置导出映射的 Fluent 设置。</param>
    void Configure(ExportMappingBuilder<TExport> setting);
}
