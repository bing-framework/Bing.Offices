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
