namespace Bing.Offices.Metadata.Property;

/// <summary>
/// 列宽属性
/// </summary>
public class ColumnWidthProperty
{
    /// <summary>
    /// 宽度
    /// </summary>
    public int Width { get; set; }

    /// <summary>
    /// 初始化一个<see cref="ColumnWidthProperty"/>类型的实例
    /// </summary>
    /// <param name="width">宽度</param>
    public ColumnWidthProperty(int width) => Width = width;
}
