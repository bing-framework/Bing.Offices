namespace Bing.Offices.Metadata;

/// <summary>
/// 合并区域信息
/// </summary>
public class MergedRegionInfo
{
    /// <summary>
    /// 索引
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// 左上角X坐标
    /// </summary>
    public int FirstRow { get; set; }

    /// <summary>
    /// 右下角X坐标
    /// </summary>
    public int LastRow { get; set; }

    /// <summary>
    /// 左上角Y坐标
    /// </summary>
    public int FirstCol { get; set; }

    /// <summary>
    /// 右下角Y坐标
    /// </summary>
    public int LastCol { get; set; }

    /// <summary>
    /// 初始化一个<see cref="MergedRegionInfo"/>类型的实例
    /// </summary>
    /// <param name="index">索引</param>
    /// <param name="firstRow">左上角X坐标</param>
    /// <param name="lastRow">右下角X坐标</param>
    /// <param name="firstCol">左上角Y坐标</param>
    /// <param name="lastCol">右下角Y坐标</param>
    public MergedRegionInfo(int index, int firstRow, int lastRow, int firstCol, int lastCol)
    {
        Index = index;
        FirstRow = firstRow;
        LastRow = lastRow;
        FirstCol = firstCol;
        LastCol = lastCol;
    }
}