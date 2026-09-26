using System;
using System.Collections.Generic;
using Bing.Offices.Attributes;

namespace Bing.Offices.Testing.Models;

/// <summary>
/// 跨 Provider 标量合同使用的一行数据。
/// </summary>
public sealed class ScalarContractRow
{
    /// <summary>
    /// 获取或设置业务编码。
    /// </summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置整数数量。
    /// </summary>
    public int Count { get; set; }
    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public decimal Amount { get; set; }
    /// <summary>
    /// 获取或设置布尔状态。
    /// </summary>
    public bool Enabled { get; set; }
    /// <summary>
    /// 获取或设置枚举类型。
    /// </summary>
    public ContractKind Kind { get; set; }
    /// <summary>
    /// 获取或设置日期。
    /// </summary>
    public DateTime Date { get; set; }
    /// <summary>
    /// 获取或设置可空整数。
    /// </summary>
    public int? Optional { get; set; }
    /// <summary>
    /// 获取或设置动态列值。
    /// </summary>
    [DynamicColumn]
    public Dictionary<string, object> Values { get; set; } = new(StringComparer.Ordinal);
}

/// <summary>
/// 标量合同的根 Workbook。
/// </summary>
public sealed class ScalarContractWorkbook
{
    /// <summary>
    /// 获取导入行集合。
    /// </summary>
    public List<ScalarContractRow> Rows { get; } = new();
}

/// <summary>
/// 标量合同中的枚举值。
/// </summary>
public enum ContractKind
{
    /// <summary>
    /// 活动状态。
    /// </summary>
    Active,
    /// <summary>
    /// 等待状态。
    /// </summary>
    Pending
}

/// <summary>
/// 动态列合同使用的一行数据。
/// </summary>
public sealed class DynamicContractRow
{
    /// <summary>
    /// 获取或设置业务编码。
    /// </summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置金额。
    /// </summary>
    public decimal Amount { get; set; }
    /// <summary>
    /// 获取或设置动态列值。
    /// </summary>
    [DynamicColumn]
    public Dictionary<string, object> Values { get; set; } = new(StringComparer.Ordinal);
}

/// <summary>
/// 动态列合同的根 Workbook。
/// </summary>
public sealed class DynamicContractWorkbook
{
    /// <summary>
    /// 获取导入行集合。
    /// </summary>
    public List<DynamicContractRow> Rows { get; } = new();
}

/// <summary>
/// 显示值映射合同使用的一行数据。
/// </summary>
public sealed class MappingContractRow
{
    /// <summary>
    /// 获取或设置业务编码。
    /// </summary>
    public string Code { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置显示名称。
    /// </summary>
    public string Name { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置数量。
    /// </summary>
    public int Quantity { get; set; }
}

/// <summary>
/// 映射合同的根 Workbook。
/// </summary>
public sealed class MappingContractWorkbook
{
    /// <summary>
    /// 获取导入行集合。
    /// </summary>
    public List<MappingContractRow> Rows { get; } = new();
}

/// <summary>
/// 校验合同使用的一行数据。
/// </summary>
public sealed class ValidationContractRow
{
    /// <summary>
    /// 获取或设置必填编码。
    /// </summary>
    [ExcelRequired]
    public string Code { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置数量。
    /// </summary>
    public int Quantity { get; set; }
}

/// <summary>
/// 校验合同的根 Workbook。
/// </summary>
public sealed class ValidationContractWorkbook
{
    /// <summary>
    /// 获取导入行集合。
    /// </summary>
    public List<ValidationContractRow> Rows { get; } = new();
}

/// <summary>
/// 关系合同中的父项。
/// </summary>
public sealed class RelationContractParent
{
    /// <summary>
    /// 获取或设置父项订单号。
    /// </summary>
    public string OrderNo { get; set; } = string.Empty;
    /// <summary>
    /// 获取子项导航集合。
    /// </summary>
    [ExcelIgnore]
    public List<RelationContractChild> Items { get; } = new();
}

/// <summary>
/// 关系合同中的子项。
/// </summary>
public sealed class RelationContractChild
{
    /// <summary>
    /// 获取或设置关联订单号。
    /// </summary>
    public string OrderNo { get; set; } = string.Empty;
    /// <summary>
    /// 获取或设置子项名称。
    /// </summary>
    public string Name { get; set; } = string.Empty;
}

/// <summary>
/// 关系合同的根 Workbook。
/// </summary>
public sealed class RelationContractWorkbook
{
    /// <summary>
    /// 获取父项集合。
    /// </summary>
    public List<RelationContractParent> Parents { get; } = new();
    /// <summary>
    /// 获取子项集合。
    /// </summary>
    public List<RelationContractChild> Children { get; } = new();
}

/// <summary>
/// 资源限制合同使用的唯一值数据行。
/// </summary>
public sealed class UniqueContractRow
{
    /// <summary>
    /// 获取或设置唯一编码。
    /// </summary>
    [Bing.Offices.Attributes.ExcelUnique]
    public string Code { get; set; } = string.Empty;
}

/// <summary>
/// 资源限制合同使用的唯一值 Workbook。
/// </summary>
public sealed class UniqueContractWorkbook
{
    /// <summary>
    /// 获取唯一编码行集合。
    /// </summary>
    public List<UniqueContractRow> Rows { get; } = new();
}
