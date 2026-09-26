using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Bing.Offices.Imports;
using Bing.Offices.Testing.Models;

namespace Bing.Offices.Testing.Snapshots;

/// <summary>
/// 稳定的标量行快照。
/// </summary>
/// <param name="Code">业务编码。</param>
/// <param name="Count">整数数量。</param>
/// <param name="Amount">按固定小数位规范化的金额。</param>
/// <param name="Enabled">布尔状态。</param>
/// <param name="Kind">枚举状态。</param>
/// <param name="Date">按固定日期格式规范化的文本。</param>
/// <param name="Optional">可空整数的规范化文本。</param>
/// <param name="Region">动态区域值或缺失标记。</param>
public sealed record ScalarRowSnapshot(
    string Code,
    int Count,
    string Amount,
    bool Enabled,
    ContractKind Kind,
    string Date,
    string Optional,
    string Region);

/// <summary>
/// 稳定的动态列行快照。
/// </summary>
/// <param name="Code">业务编码。</param>
/// <param name="Amount">按固定小数位规范化的金额。</param>
/// <param name="Region">动态区域值或缺失标记。</param>
public sealed record DynamicRowSnapshot(string Code, string Amount, string Region);

/// <summary>
/// 稳定的映射行快照。
/// </summary>
/// <param name="Code">业务编码。</param>
/// <param name="Name">显示名称。</param>
/// <param name="Quantity">数量。</param>
public sealed record MappingRowSnapshot(string Code, string Name, int Quantity);

/// <summary>
/// 稳定的关系快照。
/// </summary>
public sealed class RelationSnapshot : IEquatable<RelationSnapshot>
{
    /// <summary>
    /// 初始化一个 <see cref="RelationSnapshot"/> 类型的实例。
    /// </summary>
    /// <param name="orderNo">父项订单号。</param>
    /// <param name="itemNames">按输入顺序排列的子项名称。</param>
    public RelationSnapshot(string orderNo, IReadOnlyList<string> itemNames)
    {
        OrderNo = orderNo;
        ItemNames = itemNames;
    }

    /// <summary>
    /// 获取父项订单号。
    /// </summary>
    public string OrderNo { get; }
    /// <summary>
    /// 获取按输入顺序排列的子项名称。
    /// </summary>
    public IReadOnlyList<string> ItemNames { get; }

    /// <inheritdoc />
    public bool Equals(RelationSnapshot other) => other != null
        && string.Equals(OrderNo, other.OrderNo, StringComparison.Ordinal)
        && ItemNames.SequenceEqual(other.ItemNames, StringComparer.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object obj) => Equals(obj as RelationSnapshot);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        var hash = new HashCode();
        hash.Add(OrderNo, StringComparer.Ordinal);
        foreach (var item in ItemNames)
            hash.Add(item, StringComparer.Ordinal);
        return hash.ToHashCode();
    }

    /// <inheritdoc />
    public override string ToString() => $"RelationSnapshot {{ OrderNo = {OrderNo}, ItemNames = [{string.Join(",", ItemNames)}] }}";
}

/// <summary>
/// 稳定的导入错误快照。
/// </summary>
/// <param name="Code">导入错误分类。</param>
/// <param name="SheetName">发生错误的工作表名称。</param>
/// <param name="RowIndex">错误对应的行索引。</param>
/// <param name="ColumnIndex">错误对应的列索引。</param>
/// <param name="PropertyName">目标实体属性名称。</param>
/// <param name="ColumnKey">错误对应的列键。</param>
/// <param name="Header">错误对应的表头文本。</param>
/// <param name="FirstRowNumber">重复值首次出现的行号；不适用时为 null。</param>
/// <param name="RawValue">原始输入值的文本；原始值为空时为 null。</param>
public sealed record ImportErrorSnapshot(
    ExcelImportErrorCode Code,
    string SheetName,
    int RowIndex,
    int ColumnIndex,
    string PropertyName,
    string ColumnKey,
    string Header,
    int? FirstRowNumber,
    string RawValue);

/// <summary>
/// 创建和比较 Provider 无关的预期快照。
/// </summary>
public static class ContractSnapshots
{
    /// <summary>
    /// 获取独立编写的标量预期快照。
    /// </summary>
    /// <returns>独立定义的预期标量行快照集合。</returns>
    public static IReadOnlyList<ScalarRowSnapshot> ExpectedScalarRows() => new[]
    {
        new ScalarRowSnapshot("A", 7, "12.50", true, ContractKind.Active, "2026-09-04", "<null>", "<missing>"),
        new ScalarRowSnapshot("B", 8, "3.25", false, ContractKind.Pending, "2026-09-05", "11", "<missing>")
    };

    /// <summary>
    /// 获取独立编写的动态列预期快照。
    /// </summary>
    /// <returns>独立定义的预期动态行快照集合。</returns>
    public static IReadOnlyList<DynamicRowSnapshot> ExpectedDynamicRows() => new[]
    {
        new DynamicRowSnapshot("A", "1.25", "east"),
        new DynamicRowSnapshot("B", "9.50", "west")
    };

    /// <summary>
    /// 获取独立编写的映射预期快照。
    /// </summary>
    /// <returns>独立定义的预期映射行快照集合。</returns>
    public static IReadOnlyList<MappingRowSnapshot> ExpectedMappingRows() => new[]
    {
        new MappingRowSnapshot("A", "Alice", 2),
        new MappingRowSnapshot("B", "Bob", 5)
    };

    /// <summary>
    /// 将标量结果规范化为稳定快照。
    /// </summary>
    /// <param name="workbook">待规范化的导入工作簿结果。</param>
    /// <returns>按输入行顺序排列的规范化快照集合。</returns>
    public static IReadOnlyList<ScalarRowSnapshot> ScalarRows(ScalarContractWorkbook workbook) =>
        workbook.Rows.Select(row => new ScalarRowSnapshot(
            row.Code,
            row.Count,
            row.Amount.ToString("0.00", CultureInfo.InvariantCulture),
            row.Enabled,
            row.Kind,
            row.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            row.Optional?.ToString(CultureInfo.InvariantCulture) ?? "<null>",
            row.Values.TryGetValue("region", out var value) ? value?.ToString() ?? "<null>" : "<missing>"))
        .ToArray();

    /// <summary>
    /// 将动态列结果规范化为稳定快照。
    /// </summary>
    /// <param name="workbook">待规范化的导入工作簿结果。</param>
    /// <returns>按输入行顺序排列的规范化快照集合。</returns>
    public static IReadOnlyList<DynamicRowSnapshot> DynamicRows(DynamicContractWorkbook workbook) =>
        workbook.Rows.Select(row => new DynamicRowSnapshot(
            row.Code,
            row.Amount.ToString("0.00", CultureInfo.InvariantCulture),
            row.Values.TryGetValue("region", out var value) ? value?.ToString() ?? "<null>" : "<missing>"))
        .ToArray();

    /// <summary>
    /// 将映射结果规范化为稳定快照。
    /// </summary>
    /// <param name="workbook">待规范化的导入工作簿结果。</param>
    /// <returns>按输入行顺序排列的规范化快照集合。</returns>
    public static IReadOnlyList<MappingRowSnapshot> MappingRows(MappingContractWorkbook workbook) =>
        workbook.Rows.Select(row => new MappingRowSnapshot(row.Code, row.Name, row.Quantity)).ToArray();

    /// <summary>
    /// 将关系结果规范化为稳定快照。
    /// </summary>
    /// <param name="workbook">待规范化的导入工作簿结果。</param>
    /// <returns>按父项顺序排列的关系快照集合。</returns>
    public static IReadOnlyList<RelationSnapshot> Relations(RelationContractWorkbook workbook) =>
        workbook.Parents.Select(parent => new RelationSnapshot(parent.OrderNo,
            parent.Items.Select(item => item.Name).ToArray())).ToArray();

    /// <summary>
    /// 将结构化错误规范化为稳定快照。
    /// </summary>
    /// <param name="errors">待规范化的导入错误集合。</param>
    /// <returns>按原错误顺序排列的规范化快照集合。</returns>
    public static IReadOnlyList<ImportErrorSnapshot> Errors(IReadOnlyList<ExcelImportError> errors) =>
        errors.Select(error => new ImportErrorSnapshot(error.Code, error.SheetName, error.RowIndex,
            error.ColumnIndex, error.PropertyName, error.ColumnKey, error.Header, error.FirstRowNumber,
            error.RawValue?.ToString())).ToArray();
}

/// <summary>
/// 提供程序无关的快照比较器。
/// </summary>
/// <remarks>不依赖 xUnit 或提供程序实现。</remarks>
public static class ContractSnapshotComparer
{
    /// <summary>
    /// 比较两个序列并返回第一处差异描述。
    /// </summary>
    /// <typeparam name="T">快照序列的元素类型。</typeparam>
    /// <param name="expected">预期序列。</param>
    /// <param name="actual">实际序列。</param>
    /// <returns>第一处差异说明；两个序列完全相等时为空字符串。</returns>
    public static string Difference<T>(IReadOnlyList<T> expected, IReadOnlyList<T> actual)
    {
        if (expected.Count != actual.Count)
            return $"count expected={expected.Count}, actual={actual.Count}";
        for (var index = 0; index < expected.Count; index++)
        {
            if (!EqualityComparer<T>.Default.Equals(expected[index], actual[index]))
                return $"index={index}, expected={expected[index]}, actual={actual[index]}";
        }
        return string.Empty;
    }

    /// <summary>
    /// 判断两个序列是否完全相等。
    /// </summary>
    /// <typeparam name="T">快照序列的元素类型。</typeparam>
    /// <param name="expected">预期序列。</param>
    /// <param name="actual">实际序列。</param>
    /// <returns>序列长度和对应元素均相等时为 true，否则为 false。</returns>
    public static bool Equal<T>(IReadOnlyList<T> expected, IReadOnlyList<T> actual) =>
        Difference(expected, actual).Length == 0;
}
