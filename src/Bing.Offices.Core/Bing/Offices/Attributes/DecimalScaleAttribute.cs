namespace Bing.Offices.Attributes;

/// <summary>
/// 小数位数
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public class DecimalScaleAttribute : Attribute
{
    /// <summary>
    /// 保留小数位数
    /// </summary>
    public byte Scale { get; set; }

    /// <summary>
    /// 初始化一个<see cref="DecimalScaleAttribute"/>类型的实例
    /// </summary>
    /// <param name="scale">保留小数位数，默认：2</param>
    public DecimalScaleAttribute(byte scale = 2) => Scale = scale;
}