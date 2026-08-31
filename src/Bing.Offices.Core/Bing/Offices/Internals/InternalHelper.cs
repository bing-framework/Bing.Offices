namespace Bing.Offices.Internals;

/// <summary>
/// 生成和还原重复列标题内部标记的辅助方法。
/// </summary>
internal static class InternalHelper
{
    /// <summary>
    /// 为重复列标题追加内部唯一标记，以便在 CSV 或表格处理中保留列身份。
    /// </summary>
    /// <param name="columnName">列名</param>
    /// <returns>带内部重复列标记的列名。</returns>
    public static string GetEncodedColumnName(string columnName) => $"{columnName}{InternalConst.DuplicateColumnMark}{Guid.NewGuid():N}";

    /// <summary>
    /// 移除内部重复列标记，恢复调用方可见的原始列标题。
    /// </summary>
    /// <param name="columnName">列名</param>
    /// <returns>去除内部标记后的列名；没有标记时返回原值。</returns>
    public static string GetDecodeColumnName(string columnName)
    {
        var duplicateMarkIndex = columnName.IndexOf(InternalConst.DuplicateColumnMark, StringComparison.OrdinalIgnoreCase);
        if (duplicateMarkIndex > 0)
            return columnName.Substring(0, duplicateMarkIndex);
        return columnName;
    }
}
