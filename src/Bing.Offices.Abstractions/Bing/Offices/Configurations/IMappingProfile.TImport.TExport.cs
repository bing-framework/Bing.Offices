namespace Bing.Offices.Configurations;

/// <summary>
/// 描述导入和导出模型的映射 Profile。
/// </summary>
/// <typeparam name="TImport">导入模型类型。</typeparam>
/// <typeparam name="TExport">导出模型类型。</typeparam>
public interface IMappingProfile<TImport, TExport>
    where TImport : class, new()
    where TExport : class, new()
{
    /// <summary>
    /// 配置导入和导出方向的映射。
    /// </summary>
    /// <param name="setting">供 Profile 配置双向映射的 Fluent 设置。</param>
    void Configure(FluentSetting<TImport, TExport> setting);
}
