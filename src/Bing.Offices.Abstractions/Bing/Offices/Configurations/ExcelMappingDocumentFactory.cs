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
    /// <returns>包含请求级配置独立副本的规范化映射文档。</returns>
    public static ExcelMappingDocument Create<T>(ExcelMappingConfiguration requestConfiguration,
        MappingDirection direction) where T : class, new() =>
        Create<T>(null, requestConfiguration, direction);

    /// <summary>
    /// 选择文档方向并创建独立快照；请求级配置由 Plan Factory 在最终编译阶段合并。
    /// </summary>
    /// <typeparam name="T">方向模型类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="requestConfiguration">请求级覆盖配置。</param>
    /// <param name="direction">目标映射方向。</param>
    /// <returns>选择指定方向并合并请求级配置后的独立映射文档。</returns>
    public static ExcelMappingDocument Create<T>(ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
        where T : class, new()
    {
        if (!Enum.IsDefined(typeof(MappingDirection), direction))
            throw new ArgumentOutOfRangeException(nameof(direction));

        var directionConfiguration = document == null ? null :
            direction == MappingDirection.Import ? document.Import : document.Export;
        var normalizedDirectionConfiguration = MappingConfigurationMerger.Merge(directionConfiguration,
            requestConfiguration, MappingSourceKind.Request);
        return new ExcelMappingDocument
        {
            Version = document?.Version ?? 2,
            TenantId = document?.TenantId,
            ConfigurationVersion = document?.ConfigurationVersion,
            UseConventionFallback = document?.UseConventionFallback ?? false,
            Import = direction == MappingDirection.Import
                ? CloneOrNull(normalizedDirectionConfiguration)
                : CloneOrNull(document?.Import),
            Export = direction == MappingDirection.Export
                ? CloneOrNull(normalizedDirectionConfiguration)
                : CloneOrNull(document?.Export)
        };
    }

    /// <summary>复制可空映射配置。</summary>
    /// <param name="configuration">待复制的映射配置；为 <see langword="null" /> 时返回 <see langword="null" />。</param>
    /// <returns>映射配置的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    private static ExcelMappingConfiguration CloneOrNull(ExcelMappingConfiguration configuration) =>
        configuration == null ? null : MappingConfigurationCloner.Clone(configuration, configuration.SourceKind);
}
