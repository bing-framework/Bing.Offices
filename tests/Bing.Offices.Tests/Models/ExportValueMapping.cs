using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models;

/// <summary>
/// 导出值映射
/// </summary>
public class ExportValueMapping
{
    /// <summary>
    /// 索引
    /// </summary>
    public int Index { get; set; }

    [ValueMapping("A 系统", "100")]
    [ValueMapping("B 系统", "200")]
    [ValueMapping("C 系统", "300")]
    [ValueMapping("D 系统", "400")]
    public string Code { get; set; }
}