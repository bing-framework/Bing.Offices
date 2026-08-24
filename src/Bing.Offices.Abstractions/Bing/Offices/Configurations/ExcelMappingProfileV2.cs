namespace Bing.Offices.Configurations;

/// <summary>
/// 导入和导出模型分离的不可变 Mapping Profile。
/// </summary>
/// <typeparam name="TImport">导入模型类型。</typeparam>
/// <typeparam name="TExport">导出模型类型。</typeparam>
public sealed class ExcelMappingProfile<TImport, TExport> : IMappingProfileSnapshot
    where TImport : class, new()
    where TExport : class, new()
{
    private readonly ExcelMappingConfiguration _importConfiguration;
    private readonly ExcelMappingConfiguration _exportConfiguration;

    /// <summary>
    /// 使用方向 Profile 构建不可变快照。
    /// </summary>
    public ExcelMappingProfile(IMappingProfile<TImport, TExport> profile)
    {
        if (profile == null)
            throw new ArgumentNullException(nameof(profile));
        var setting = new FluentSetting<TImport, TExport>();
        profile.Configure(setting);
        _importConfiguration = MappingConfigurationCloner.Clone(setting.BuildImportConfiguration(),
            MappingSourceKind.Profile);
        _exportConfiguration = MappingConfigurationCloner.Clone(setting.BuildExportConfiguration(),
            MappingSourceKind.Profile);
    }

    /// <summary>
    /// 使用 Fluent 配置委托构建不可变快照。
    /// </summary>
    public ExcelMappingProfile(Action<FluentSetting<TImport, TExport>> configure)
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var setting = new FluentSetting<TImport, TExport>();
        configure(setting);
        _importConfiguration = MappingConfigurationCloner.Clone(setting.BuildImportConfiguration(),
            MappingSourceKind.Profile);
        _exportConfiguration = MappingConfigurationCloner.Clone(setting.BuildExportConfiguration(),
            MappingSourceKind.Profile);
    }

    internal ExcelMappingProfile(ExcelMappingConfiguration importConfiguration,
        ExcelMappingConfiguration exportConfiguration)
    {
        _importConfiguration = MappingConfigurationCloner.Clone(importConfiguration, MappingSourceKind.Profile);
        _exportConfiguration = MappingConfigurationCloner.Clone(exportConfiguration, MappingSourceKind.Profile);
    }

    /// <inheritdoc />
    public Type ImportType => typeof(TImport);

    /// <inheritdoc />
    public Type ExportType => typeof(TExport);

    /// <inheritdoc />
    public ExcelMappingConfiguration ImportConfiguration =>
        MappingConfigurationCloner.Clone(_importConfiguration, MappingSourceKind.Profile);

    /// <inheritdoc />
    public ExcelMappingConfiguration ExportConfiguration =>
        MappingConfigurationCloner.Clone(_exportConfiguration, MappingSourceKind.Profile);
}
