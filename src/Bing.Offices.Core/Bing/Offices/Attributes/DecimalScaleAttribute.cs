namespace Bing.Offices.Attributes;

/// <summary>
/// 指定数值导出时保留的小数位数。
/// </summary>
[AttributeUsage(AttributeTargets.Property, Inherited = true, AllowMultiple = false)]
public class DecimalScaleAttribute : Attribute
{
    /// <summary>
    /// 获取或设置数值导出时保留的小数位数。
    /// </summary>
    public byte Scale { get; set; }

    /// <summary>
    /// 初始化一个 <see cref="DecimalScaleAttribute" /> 类型的实例。
    /// </summary>
    /// <param name="scale">要保留的小数位数；默认值为 2。</param>
    public DecimalScaleAttribute(byte scale = 2) => Scale = scale;
}
