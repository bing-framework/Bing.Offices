using System.Collections.Generic;

namespace Bing.Offices.Testing.Models;

/// <summary>
/// 跨 Provider 实体基线合同的根聚合模型。
/// </summary>
public sealed class EntityContractRoot
{
    /// <summary>
    /// 获取或设置单据编号。
    /// </summary>
    public int Number { get; set; }

    /// <summary>
    /// 获取或设置客户名称。
    /// </summary>
    public string Customer { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置单据标题。
    /// </summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置明细行集合。
    /// </summary>
    public List<EntityContractLine> Lines { get; set; } = new();
}

/// <summary>
/// 跨 Provider 实体基线合同的明细行模型。
/// </summary>
public sealed class EntityContractLine
{
    /// <summary>
    /// 获取或设置明细编码。
    /// </summary>
    public string Code { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置明细数量。
    /// </summary>
    public int Quantity { get; set; }
}

/// <summary>
/// 跨 Provider 实体动态列合同的根聚合模型。
/// </summary>
public sealed class EntityDynamicContractRoot
{
    /// <summary>
    /// 获取或设置动态明细行集合。
    /// </summary>
    public List<EntityDynamicContractLine> Lines { get; set; } = new();
}

/// <summary>
/// 跨 Provider 实体动态列合同的明细行模型。
/// </summary>
public sealed class EntityDynamicContractLine
{
    /// <summary>
    /// 获取或设置明细名称。
    /// </summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// 获取或设置明细数量。
    /// </summary>
    public int Quantity { get; set; }

    /// <summary>
    /// 获取或设置动态列字典。
    /// </summary>
    [Bing.Offices.Attributes.DynamicColumn]
    public IDictionary<string, object> CustomFields { get; set; } =
        new Dictionary<string, object>(System.StringComparer.OrdinalIgnoreCase);
}
