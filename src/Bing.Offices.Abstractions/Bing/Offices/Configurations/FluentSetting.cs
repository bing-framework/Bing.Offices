namespace Bing.Offices.Configurations;

/// <summary>
/// 导入和导出方向隔离的 Profile Fluent 设置。
/// </summary>
/// <typeparam name="TImport">导入模型类型。</typeparam>
/// <typeparam name="TExport">导出模型类型。</typeparam>
public sealed class FluentSetting<TImport, TExport>
    where TImport : class, new()
    where TExport : class, new()
{
    private readonly ImportMappingBuilder<TImport> _import = new();
    private readonly ExportMappingBuilder<TExport> _export = new();

    /// <summary>
    /// 获取导入方向设置。
    /// </summary>
    public ImportMappingBuilder<TImport> Import => _import;

    /// <summary>
    /// 获取导出方向设置。
    /// </summary>
    public ExportMappingBuilder<TExport> Export => _export;

    /// <summary>
    /// 创建导入配置快照。
    /// </summary>
    public ExcelMappingConfiguration BuildImportConfiguration() => _import.Build(MappingSourceKind.Profile);

    /// <summary>
    /// 创建导出配置快照。
    /// </summary>
    public ExcelMappingConfiguration BuildExportConfiguration() => _export.Build(MappingSourceKind.Profile);
}
