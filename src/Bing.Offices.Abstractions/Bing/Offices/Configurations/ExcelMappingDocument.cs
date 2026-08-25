using System.Collections.Generic;

namespace Bing.Offices.Configurations;

/// <summary>
/// v2 规范化映射文档；Import 和 Export 配置相互独立。
/// </summary>
public sealed class ExcelMappingDocument
{
    /// <summary>
    /// 获取或设置文档版本；v2 文档固定为 2。
    /// </summary>
    public int Version { get; set; } = 2;

    /// <summary>
    /// 获取或设置租户缓存隔离键。
    /// </summary>
    public string TenantId { get; set; }

    /// <summary>
    /// 获取或设置配置版本，用于计划缓存失效和并发隔离。
    /// </summary>
    public string ConfigurationVersion { get; set; }

    /// <summary>
    /// 获取或设置导入方向配置。
    /// </summary>
    public ExcelMappingConfiguration Import { get; set; } = new ExcelMappingConfiguration
    {
        Columns = new List<ExcelColumnConfiguration>()
    };

    /// <summary>
    /// 获取或设置导出方向配置。
    /// </summary>
    public ExcelMappingConfiguration Export { get; set; } = new ExcelMappingConfiguration
    {
        Columns = new List<ExcelColumnConfiguration>()
    };
}
