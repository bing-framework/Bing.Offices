namespace Bing.Offices.Metadata;

/// <summary>
/// 图片数据及其覆盖的工作表零基行列范围。
/// </summary>
public class PictureInfo
{
    /// <summary>
    /// 获取或设置图片覆盖范围的最小零基行号。
    /// </summary>
    public int MinRow { get; set; }

    /// <summary>
    /// 获取或设置图片覆盖范围的最大零基行号。
    /// </summary>
    public int MaxRow { get; set; }

    /// <summary>
    /// 获取或设置图片覆盖范围的最小零基列号。
    /// </summary>
    public int MinCol { get; set; }

    /// <summary>
    /// 获取或设置图片覆盖范围的最大零基列号。
    /// </summary>
    public int MaxCol { get; set; }

    /// <summary>
    /// 获取或设置图片原始字节数据。
    /// </summary>
    public byte[] PictureData { get; set; }

    /// <summary>
    /// 获取或设置图片锚点和外观样式；添加图片时不能为 null。
    /// </summary>
    public PictureStyle PictureStyle { get; set; }

    /// <summary>
    /// 初始化一个 <see cref="PictureInfo" /> 类型的实例。
    /// </summary>
    /// <param name="minRow">图片覆盖范围的最小零基行号。</param>
    /// <param name="maxRow">图片覆盖范围的最大零基行号。</param>
    /// <param name="minCol">图片覆盖范围的最小零基列号。</param>
    /// <param name="maxCol">图片覆盖范围的最大零基列号。</param>
    /// <param name="pictureData">图片原始字节数据。</param>
    /// <param name="pictureStyle">图片锚点和外观样式。</param>
    public PictureInfo(int minRow, int maxRow, int minCol, int maxCol, byte[] pictureData,
        PictureStyle pictureStyle)
    {
        MinRow = minRow;
        MaxRow = maxRow;
        MinCol = minCol;
        MaxCol = maxCol;
        PictureData = pictureData;
        PictureStyle = pictureStyle;
    }
}
