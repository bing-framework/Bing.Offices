using System;

namespace Bing.Offices.Configurations;

/// <summary>
/// 将旧 Profile、Document 和请求配置归一化为单一不可变文档。
/// </summary>
public static class ExcelMappingDocumentFactory
{
    /// <summary>
    /// 创建指定方向的规范化文档。
    /// </summary>
    [Obsolete("请改用强类型双模型 Profile 或 ExcelMappingDocument。", false)]
    public static ExcelMappingDocument Create<T>(object profile, ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
        where T : class, new()
    {
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));

        var profileConfiguration = ResolveProfileConfiguration<T>(profile, direction);
        var documentConfiguration = document == null
            ? null
            : direction == MappingDirection.Import ? document.Import : document.Export;
        var merged = MappingConfigurationMerger.Merge(profileConfiguration, documentConfiguration,
            MappingSourceKind.Document);
        merged = MappingConfigurationMerger.Merge(merged, requestConfiguration, MappingSourceKind.Request)
            ?? new ExcelMappingConfiguration { SourceKind = MappingSourceKind.Request };

        return new ExcelMappingDocument
        {
            Version = document?.Version ?? 2,
            Profile = document?.Profile,
            ModelAlias = document?.ModelAlias,
            TenantId = document?.TenantId,
            ConfigurationVersion = document?.ConfigurationVersion,
            Import = direction == MappingDirection.Import ? merged : CloneOrEmpty(document?.Import),
            Export = direction == MappingDirection.Export ? merged : CloneOrEmpty(document?.Export)
        };
    }

    /// <summary>
    /// 创建未附加 Profile 和 Document 的规范化文档。
    /// </summary>
    public static ExcelMappingDocument Create<T>(ExcelMappingConfiguration requestConfiguration,
        MappingDirection direction) where T : class, new() =>
        Create<T>(null, null, requestConfiguration, direction);

    private static ExcelMappingConfiguration ResolveProfileConfiguration<T>(object profile,
        MappingDirection direction) where T : class, new()
    {
        if (profile == null)
            return null;
        if (profile is ExcelMappingProfile<T> legacy)
            return legacy.Configuration;
        if (!(profile is IMappingProfileSnapshot snapshot))
            throw new ArgumentException("映射 Profile 类型不受支持。", nameof(profile));

        var expectedType = direction == MappingDirection.Import ? snapshot.ImportType : snapshot.ExportType;
        if (expectedType != typeof(T))
            throw new ArgumentException(
                $"映射 Profile 的{(direction == MappingDirection.Import ? "导入" : "导出")}模型类型不匹配: {typeof(T).FullName}",
                nameof(profile));
        return direction == MappingDirection.Import
            ? snapshot.ImportConfiguration
            : snapshot.ExportConfiguration;
    }

    private static ExcelMappingConfiguration CloneOrEmpty(ExcelMappingConfiguration configuration) =>
        configuration == null
            ? new ExcelMappingConfiguration { SourceKind = MappingSourceKind.Convention }
            : MappingConfigurationCloner.Clone(configuration, configuration.SourceKind);
}
