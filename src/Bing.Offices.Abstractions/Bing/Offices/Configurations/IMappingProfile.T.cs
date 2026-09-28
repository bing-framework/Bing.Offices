namespace Bing.Offices.Configurations;

/// <summary>
/// 同一模型同时用于导入和导出的映射 Profile。
/// </summary>
/// <typeparam name="T">模型类型。</typeparam>
public interface IMappingProfile<T>
    where T : class, new()
{
    /// <summary>
    /// 配置同一模型的导入和导出方向映射。
    /// </summary>
    /// <param name="setting">供 Profile 配置同一模型双向映射的 Fluent 设置。</param>
    void Configure(FluentSetting<T, T> setting);
}
