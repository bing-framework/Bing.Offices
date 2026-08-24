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
    /// <param name="setting">方向隔离的 Fluent 设置。</param>
    void Configure(FluentSetting<TImport, TExport> setting);
}

/// <summary>
/// 同一模型同时用于导入和导出的映射 Profile。
/// </summary>
/// <typeparam name="T">模型类型。</typeparam>
public interface IMappingProfile<T> : IMappingProfile<T, T>
    where T : class, new()
{
}

/// <summary>
/// 已构建的方向映射 Profile 快照。
/// </summary>
public interface IMappingProfileSnapshot
{
    /// <summary>
    /// 获取导入模型类型。
    /// </summary>
    Type ImportType { get; }

    /// <summary>
    /// 获取导出模型类型。
    /// </summary>
    Type ExportType { get; }

    /// <summary>
    /// 获取导入配置快照。
    /// </summary>
    ExcelMappingConfiguration ImportConfiguration { get; }

    /// <summary>
    /// 获取导出配置快照。
    /// </summary>
    ExcelMappingConfiguration ExportConfiguration { get; }
}
