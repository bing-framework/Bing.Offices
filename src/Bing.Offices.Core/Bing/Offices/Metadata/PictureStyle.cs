namespace Bing.Offices.Metadata;

/// <summary>
/// 图片在工作表中的锚点、填充和线条样式配置。
/// </summary>
public class PictureStyle
{
    /// <summary>获取或设置图片左上角锚点的水平偏移量。</summary>
    public int AnchorDx1 { get; set; }
    /// <summary>获取或设置图片右下角锚点的水平偏移量。</summary>
    public int AnchorDx2 { get; set; }
    /// <summary>获取或设置图片左上角锚点的垂直偏移量。</summary>
    public int AnchorDy1 { get; set; }
    /// <summary>获取或设置图片右下角锚点的垂直偏移量。</summary>
    public int AnchorDy2 { get; set; }
    /// <summary>获取或设置图片填充颜色的颜色值。</summary>
    public int FillColor { get; set; }
    /// <summary>获取或设置是否不绘制图片填充。</summary>
    public bool IsNoFill { get; set; }
    /// <summary>获取或设置图片边框的线型值。</summary>
    public int LineStyle { get; set; }
    /// <summary>获取或设置图片边框线条的颜色值。</summary>
    public int LineStyleColor { get; set; }
    /// <summary>获取或设置图片边框线条的宽度。</summary>
    public double LineWidth { get; set; }
}
