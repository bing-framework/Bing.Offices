using System.Collections.Generic;

namespace Bing.Offices.Configurations;

internal static class MappingDocumentCloner
{
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
