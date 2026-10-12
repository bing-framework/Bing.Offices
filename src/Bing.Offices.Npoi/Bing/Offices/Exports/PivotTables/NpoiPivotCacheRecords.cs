namespace Bing.Offices.Npoi.Exports.PivotTables;

/// <summary>
/// 保存透视缓存记录的不可变共享项索引快照。
/// </summary>
internal sealed class NpoiPivotCacheRecords
{
    /// <summary>
    /// 初始化一个 <see cref="NpoiPivotCacheRecords" /> 类型的实例。
    /// </summary>
    /// <param name="partName">缓存记录在 XLSX 包内的相对部件路径。</param>
    /// <param name="sharedItemCounts">各来源字段的共享项数量。</param>
    /// <param name="rows">按来源行顺序排列的完整字段索引。</param>
    internal NpoiPivotCacheRecords(string partName, IReadOnlyList<int> sharedItemCounts,
        IReadOnlyList<IReadOnlyList<uint>> rows)
    {
        if (string.IsNullOrWhiteSpace(partName))
            throw new ArgumentException("缓存记录部件路径不能为空。", nameof(partName));
        if (sharedItemCounts == null)
            throw new ArgumentNullException(nameof(sharedItemCounts));
        if (rows == null)
            throw new ArgumentNullException(nameof(rows));
        var counts = sharedItemCounts.ToArray();
        if (counts.Length == 0 || counts.Any(count => count < 0))
            throw new ArgumentException("缓存字段数量和共享项数量无效。", nameof(sharedItemCounts));
        var snapshot = new IReadOnlyList<uint>[rows.Count];
        for (var row = 0; row < rows.Count; row++)
        {
            var values = rows[row]?.ToArray()
                ?? throw new ArgumentException("缓存记录不能为空。", nameof(rows));
            if (values.Length != counts.Length)
                throw new ArgumentException("缓存记录必须包含全部来源字段。", nameof(rows));
            for (var column = 0; column < values.Length; column++)
                if (values[column] >= counts[column])
                    throw new ArgumentException("缓存记录的共享项索引超出范围。", nameof(rows));
            snapshot[row] = Array.AsReadOnly(values);
        }
        PartName = partName.TrimStart('/');
        if (PartName.Length == 0)
            throw new ArgumentException("缓存记录部件路径不能为空。", nameof(partName));
        SharedItemCounts = Array.AsReadOnly(counts);
        Rows = Array.AsReadOnly(snapshot);
    }

    /// <summary>
    /// 获取缓存记录的包内部件路径。
    /// </summary>
    internal string PartName { get; }

    /// <summary>
    /// 获取来源字段数量。
    /// </summary>
    internal int FieldCount => SharedItemCounts.Count;

    /// <summary>
    /// 获取各字段的共享项数量。
    /// </summary>
    internal IReadOnlyList<int> SharedItemCounts { get; }

    /// <summary>
    /// 获取完整缓存记录的索引快照。
    /// </summary>
    internal IReadOnlyList<IReadOnlyList<uint>> Rows { get; }

}
