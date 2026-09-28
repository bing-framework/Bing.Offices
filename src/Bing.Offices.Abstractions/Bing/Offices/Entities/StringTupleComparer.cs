using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 为包含工作表名称和整数坐标的元组提供忽略大小写的相等性。
/// </summary>
internal sealed class StringTupleComparer : IEqualityComparer<(string Sheet, int A, int B)>,
    IEqualityComparer<(string Sheet, int A, int B, int C, int D)>
{
    /// <summary>
    /// 获取全局共享的忽略大小写比较器。
    /// </summary>
    internal static readonly StringTupleComparer OrdinalIgnoreCase = new StringTupleComparer();
    /// <inheritdoc />
    public bool Equals((string Sheet, int A, int B) x, (string Sheet, int A, int B) y) =>
        string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase) && x.A == y.A && x.B == y.B;
    /// <inheritdoc />
    public int GetHashCode((string Sheet, int A, int B) value) =>
        Combine(StringComparer.OrdinalIgnoreCase.GetHashCode(value.Sheet ?? string.Empty), value.A, value.B);
    /// <inheritdoc />
    public bool Equals((string Sheet, int A, int B, int C, int D) x, (string Sheet, int A, int B, int C, int D) y) =>
        string.Equals(x.Sheet, y.Sheet, StringComparison.OrdinalIgnoreCase) && x.A == y.A && x.B == y.B
        && x.C == y.C && x.D == y.D;
    /// <inheritdoc />
    public int GetHashCode((string Sheet, int A, int B, int C, int D) value) => Combine(
        StringComparer.OrdinalIgnoreCase.GetHashCode(value.Sheet ?? string.Empty), value.A, value.B, value.C, value.D);

    /// <summary>
    /// 按顺序组合多个整数哈希值。
    /// </summary>
    /// <param name="values">待组合的整数哈希值。</param>
    /// <returns>合并后的哈希值。</returns>
    private static int Combine(params int[] values)
    {
        unchecked
        {
            var hash = 17;
            foreach (var value in values)
                hash = hash * 31 + value;
            return hash;
        }
    }
}
