using System;

namespace Bing.Offices.Configurations;

/// <summary>
/// 创建方向化、不可变的规范化映射文档。
/// </summary>
public static class ExcelMappingDocumentFactory
{
    /// <summary>
    /// 创建未附加文档的请求级配置文档。
    /// </summary>
    /// <typeparam name="T">方向模型类型。</typeparam>
    /// <param name="requestConfiguration">请求级覆盖配置。</param>
    /// <param name="direction">目标映射方向。</param>
    public static ExcelMappingDocument Create<T>(ExcelMappingConfiguration requestConfiguration,
        MappingDirection direction) where T : class, new() =>
        Create<T>(null, requestConfiguration, direction);

    /// <summary>
    /// 将文档方向配置与请求级覆盖配置合并为独立快照。
    /// </summary>
    /// <typeparam name="T">方向模型类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="requestConfiguration">请求级覆盖配置。</param>
    /// <param name="direction">目标映射方向。</param>
    public static ExcelMappingDocument Create<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
        where T : class, new()
    {
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));

        var directionConfiguration = document == null
            ? null
            : direction == MappingDirection.Import ? document.Import : document.Export;
        var merged = MappingConfigurationMerger.Merge(directionConfiguration, requestConfiguration,
            MappingSourceKind.Request);
        return new ExcelMappingDocument
        {
            Version = document?.Version ?? 2,
            TenantId = document?.TenantId,
            ConfigurationVersion = document?.ConfigurationVersion,
            Import = direction == MappingDirection.Import
                ? CloneOrEmpty(merged)
                : CloneOrEmpty(document?.Import),
            Export = direction == MappingDirection.Export
                ? CloneOrEmpty(merged)
                : CloneOrEmpty(document?.Export)
        };
    }

    private static ExcelMappingConfiguration CloneOrEmpty(ExcelMappingConfiguration configuration) =>
        configuration == null
            ? new ExcelMappingConfiguration { SourceKind = MappingSourceKind.Convention }
            : MappingConfigurationCloner.Clone(configuration, configuration.SourceKind);
}
