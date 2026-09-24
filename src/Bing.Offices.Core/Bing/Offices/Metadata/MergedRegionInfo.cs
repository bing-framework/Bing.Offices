namespace Bing.Offices.Metadata;

/// <summary>
/// 工作表合并区域的零基行列边界。
/// </summary>
public class MergedRegionInfo
{
    /// <summary>
    /// 获取或设置合并区域索引；未对应已登记的合并区域时可为 -1。
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 获取或设置左上角的零基行号（含）。
    /// </summary>
    public int FirstRow { get; set; }

    /// <summary>
    /// 获取或设置右下角的零基行号（含）。
    /// </summary>
    public int LastRow { get; set; }

    /// <summary>
    /// 获取或设置左上角的零基列号（含）。
    /// </summary>
    public int FirstCol { get; set; }

    /// <summary>
    /// 获取或设置右下角的零基列号（含）。
    /// </summary>
    public int LastCol { get; set; }

    /// <summary>
    /// 初始化一个 <see cref="MergedRegionInfo" /> 类型的实例。
    /// </summary>
    /// <param name="index">合并区域在工作表登记集合中的索引；未登记时为 -1。</param>
    /// <param name="firstRow">左上角的零基行号。</param>
    /// <param name="lastRow">右下角的零基行号。</param>
    /// <param name="firstCol">左上角的零基列号。</param>
    /// <param name="lastCol">右下角的零基列号。</param>
    public MergedRegionInfo(int index, int firstRow, int lastRow, int firstCol, int lastCol)
    {
        Index = index;
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstCol = firstCol;
        LastCol = lastCol;
    }
}
