using System;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;

namespace Bing.Offices.Npoi.Internals;

internal static class NpoiWorkbookPlanKeyBuilder
{
    public static string Create(Type itemType, ExcelMappingDocument document,
        ExcelMappingConfiguration requestConfiguration, MappingDirection direction)
    {
        var directionDocument = direction == MappingDirection.Import
            ? new ExcelMappingDocument { Import = requestConfiguration }
            : new ExcelMappingDocument { Export = requestConfiguration };
        return string.Join("|", itemType.AssemblyQualifiedName,
            ExcelMappingConfigurationLoader.ToJson(document ?? new ExcelMappingDocument
            {
                UseConventionFallback = true
            }), ExcelMappingConfigurationLoader.ToJson(directionDocument));
    }
}
