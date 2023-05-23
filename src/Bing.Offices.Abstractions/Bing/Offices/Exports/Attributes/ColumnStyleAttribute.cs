namespace Bing.Offices.Exports.Attributes;

/// <summary>
/// 内容样式 特性
/// </summary>
[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, Inherited = false, AllowMultiple = false)]
public class ColumnStyleAttribute : StyleAttribute
{
    /// <summary>
    /// 初始化一个<see cref="ColumnStyleAttribute"/>类型的实例
    /// </summary>
    /// <param name="isBold">加粗</param>
    public ColumnStyleAttribute(bool isBold) => Style.IsBold = isBold;

    /// <summary>
    /// 初始化一个<see cref="ColumnStyleAttribute"/>类型的实例
    /// </summary>
    public ColumnStyleAttribute() { }
}
