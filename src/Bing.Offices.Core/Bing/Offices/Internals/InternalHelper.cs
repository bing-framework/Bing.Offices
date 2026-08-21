namespace Bing.Offices.Internals;

/// <summary>
/// 内部帮助类
/// </summary>
internal static class InternalHelper
{
    /// <summary>
    /// 获取编码列名
    /// </summary>
    /// <param name="columnName">列名</param>
    public static string GetEncodedColumnName(string columnName) => $"{columnName}{InternalConst.DuplicateColumnMark}{Guid.NewGuid():N}";

    /// <summary>
    /// 获取解码列名
    /// </summary>
    /// <param name="columnName">列名</param>
    public static string GetDecodeColumnName(string columnName)
    {
        var duplicateMarkIndex = columnName.IndexOf(InternalConst.DuplicateColumnMark, StringComparison.OrdinalIgnoreCase);
        if (duplicateMarkIndex > 0)
            return columnName.Substring(0, duplicateMarkIndex);
        return columnName;
    }
}
