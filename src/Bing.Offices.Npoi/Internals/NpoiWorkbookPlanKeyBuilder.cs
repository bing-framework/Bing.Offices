using System;
using Bing.Offices.Configurations;
using Bing.Offices.Mappings;

namespace Bing.Offices.Npoi.Internals;

internal static class NpoiWorkbookPlanKeyBuilder
{
    /// <summary>根据实体类型、映射文档、请求配置和方向生成稳定计划键。</summary>
    /// <param name="itemType">工作表数据项类型。</param>
    /// <param name="document">规范化映射文档；为空时使用约定回退文档。</param>
    /// <param name="requestConfiguration">请求级方向配置。</param>
    /// <param name="direction">导入或导出方向。</param>
    /// <returns>包含类型和序列化配置内容的计划缓存键。</returns>
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
