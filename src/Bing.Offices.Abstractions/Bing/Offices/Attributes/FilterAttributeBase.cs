namespace Bing.Offices.Attributes;

/// <summary>
/// 表示应用于导入字段的校验筛选特性基类。
/// </summary>
public abstract class FilterAttributeBase : Attribute
{
    /// <summary>
    /// 获取或设置校验失败且规则未提供专用消息时使用的默认错误消息。
    /// </summary>
    public virtual string ErrorMsg { get; set; } = "非法";
}
