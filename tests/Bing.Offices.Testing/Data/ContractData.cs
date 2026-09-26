using System;
using System.Collections.Generic;
using Bing.Offices.Exports;
using Bing.Offices.Testing.Models;

namespace Bing.Offices.Testing.Data;

/// <summary>
/// 提供跨 Provider 合同使用的确定性输入数据。
/// </summary>
public static class ContractData
{
    /// <summary>
    /// 创建标量合同数据。
    /// </summary>
    /// <returns>标量合同数据集合。</returns>
    public static IReadOnlyList<ScalarContractRow> ScalarRows() => new[]
    {
        new ScalarContractRow
        {
            Code = "A",
            Count = 7,
            Amount = 12.5m,
            Enabled = true,
            Kind = ContractKind.Active,
            Date = new DateTime(2026, 9, 4),
            Optional = null,
            Values = new Dictionary<string, object>(StringComparer.Ordinal) { ["region"] = "east" }
        },
        new ScalarContractRow
        {
            Code = "B",
            Count = 8,
            Amount = 3.25m,
            Enabled = false,
            Kind = ContractKind.Pending,
            Date = new DateTime(2026, 9, 5),
            Optional = 11,
            Values = new Dictionary<string, object>(StringComparer.Ordinal) { ["region"] = "west" }
        }
    };

    /// <summary>
    /// 创建动态列定义。
    /// </summary>
    /// <returns>用于区域动态列的定义集合。</returns>
    public static IReadOnlyList<ExcelDynamicColumnDefinition> DynamicDefinitions() => new[]
    {
        new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        }
    };

    /// <summary>
    /// 创建动态列合同数据。
    /// </summary>
    /// <returns>动态列合同数据集合。</returns>
    public static IReadOnlyList<DynamicContractRow> DynamicRows() => new[]
    {
        new DynamicContractRow
        {
            Code = "A",
            Amount = 1.25m,
            Values = new Dictionary<string, object>(StringComparer.Ordinal) { ["region"] = "east" }
        },
        new DynamicContractRow
        {
            Code = "B",
            Amount = 9.5m,
            Values = new Dictionary<string, object>(StringComparer.Ordinal) { ["region"] = "west" }
        }
    };

    /// <summary>
    /// 创建映射合同数据。
    /// </summary>
    /// <returns>显示值映射合同数据集合。</returns>
    public static IReadOnlyList<MappingContractRow> MappingRows() => new[]
    {
        new MappingContractRow { Code = "A", Name = "Alice", Quantity = 2 },
        new MappingContractRow { Code = "B", Name = "Bob", Quantity = 5 }
    };

    /// <summary>
    /// 创建包含校验错误的合同数据。
    /// </summary>
    /// <returns>包含必填编码错误的数据集合。</returns>
    public static IReadOnlyList<ValidationContractRow> InvalidValidationRows() => new[]
    {
        new ValidationContractRow { Code = string.Empty, Quantity = 0 }
    };

    /// <summary>
    /// 创建关系合同父项数据。
    /// </summary>
    /// <returns>关系合同的父项集合。</returns>
    public static IReadOnlyList<RelationContractParent> RelationParents() => new[]
    {
        new RelationContractParent { OrderNo = "A-1" }
    };

    /// <summary>
    /// 创建关系合同子项数据。
    /// </summary>
    /// <returns>具有大小写差异关联键的子项集合。</returns>
    public static IReadOnlyList<RelationContractChild> RelationChildren() => new[]
    {
        new RelationContractChild { OrderNo = "a-1", Name = "Item-1" },
        new RelationContractChild { OrderNo = "A-1", Name = "Item-2" }
    };
}
