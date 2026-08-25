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
            Import = document.Import == null
                ? new ExcelMappingConfiguration { Columns = new List<ExcelColumnConfiguration>() }
                : MappingConfigurationCloner.Clone(document.Import, MappingSourceKind.Document),
            Export = document.Export == null
                ? new ExcelMappingConfiguration { Columns = new List<ExcelColumnConfiguration>() }
                : MappingConfigurationCloner.Clone(document.Export, MappingSourceKind.Document)
        };
    }
}
