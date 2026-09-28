namespace Bing.Offices.Imports;

/// <summary>
/// 工作表读取列范围，列索引从零开始。
/// </summary>
public sealed class ExcelReadColumnRange
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelReadColumnRange" /> 类型的实例。
    /// </summary>
    /// <param name="startIndex">起始列的零基索引。</param>
    /// <param name="count">要读取的列数。</param>
    private ExcelReadColumnRange(int startIndex, int count)
    {
        StartIndex = startIndex;
        Count = count;
    }

    /// <summary>
    /// 创建列读取范围。
    /// </summary>
    /// <param name="startIndex">起始列的零基索引。</param>
    /// <param name="count">要读取的列数。</param>
    /// <returns>验证通过的列读取范围。</returns>
    public static ExcelReadColumnRange Create(int startIndex, int count)
    {
        if (startIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(startIndex));
        if (count <= 0)
            throw new ArgumentOutOfRangeException(nameof(count));
        if ((long)startIndex + count > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(count));
        return new ExcelReadColumnRange(startIndex, count);
    }

    /// <summary>
    /// 获取起始列索引。
    /// </summary>
    public int StartIndex { get; }

    /// <summary>
    /// 获取读取列数。
    /// </summary>
    public int Count { get; }

    /// <summary>
    /// 判断指定列是否在范围内。
    /// </summary>
    /// <param name="columnIndex">待检查的零基列索引。</param>
    /// <returns>列索引位于当前范围内时为 <see langword="true" />，否则为 <see langword="false" />。</returns>
    public bool Contains(int columnIndex) => columnIndex >= StartIndex && columnIndex < StartIndex + Count;
}
