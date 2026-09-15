namespace Bing.Offices.Configurations;

/// <summary>复制映射文档及其方向配置的内部辅助类。</summary>
internal static class MappingDocumentCloner
{
    /// <summary>复制映射文档及其方向配置。</summary>
    /// <param name="document">待复制的映射文档。</param>
    /// <returns>映射文档的独立副本；输入为 <see langword="null" /> 时返回 <see langword="null" />。</returns>
    internal static ExcelMappingDocument Clone(ExcelMappingDocument document)
    {
        if (document == null)
            return null;
        return new ExcelMappingDocument
        {
            Version = document.Version,
            TenantId = document.TenantId,
            ConfigurationVersion = document.ConfigurationVersion,
            UseConventionFallback = document.UseConventionFallback,
            Import = document.Import == null ? null :
                MappingConfigurationCloner.Clone(document.Import, MappingSourceKind.Document),
            Export = document.Export == null ? null :
                MappingConfigurationCloner.Clone(document.Export, MappingSourceKind.Document)
        };
    }
}
