namespace Bing.Offices.Attributes;

/// <summary>
/// 将自定义校验特性绑定到 Stream-first 校验规则类型。
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public sealed class BindFilterAttribute : Attribute
{
    /// <summary>
    /// 初始化一个<see cref="BindFilterAttribute"/>类型的实例。
    /// </summary>
    /// <param name="ruleType">校验规则类型。</param>
    public BindFilterAttribute(Type ruleType) => RuleType = ruleType ?? throw new ArgumentNullException(nameof(ruleType));

    /// <summary>
    /// 获取校验规则类型。
    /// </summary>
    public Type RuleType { get; }
}
