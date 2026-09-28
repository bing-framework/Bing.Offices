namespace Bing.Offices.Imports;

/// <summary>
/// 按照对象引用而非对象值比较关联实体的比较器。
/// </summary>
internal sealed class ReferenceObjectComparer : IEqualityComparer<object>
{
    /// <summary>
    /// 保存可复用的引用比较器实例。
    /// </summary>
    internal static readonly ReferenceObjectComparer Instance = new();

    /// <inheritdoc />
    public new bool Equals(object x, object y) => ReferenceEquals(x, y);

    /// <inheritdoc />
    public int GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

#pragma warning restore CS0618
