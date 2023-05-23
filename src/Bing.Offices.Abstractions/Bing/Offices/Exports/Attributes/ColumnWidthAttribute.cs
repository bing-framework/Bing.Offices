namespace Bing.Offices.Exports.Attributes;

/// <summary>
/// 列宽 特性
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field | AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
public class ColumnWidthAttribute : Attribute
{
    /// <summary>
    /// 最小宽度
    /// </summary>
    public int MinWidth { get; set; }

    /// <summary>
    /// 最大宽度
    /// </summary>
    public int MaxWidth { get; set; }

    /// <summary>
    /// 初始化一个<see cref="ColumnWidthAttribute"/>类型的实例
    /// </summary>
    /// <param name="minWidth">最小宽度</param>
    /// <param name="maxWidth">最大宽度</param>
    public ColumnWidthAttribute(int minWidth, int maxWidth)
    {
        MinWidth = minWidth;
        MaxWidth = maxWidth;
    }
}
