using Bing.Offices.Attributes;

namespace Bing.Offices.Decorators;

/// <summary>
/// 装饰器特性比较器
/// </summary>
internal class DecoratorAttributeComparer : IEqualityComparer<DecoratorAttributeBase>
{
    /// <summary>
    /// 是否相等
    /// </summary>
    /// <param name="x">对象A</param>
    /// <param name="y">对象B</param>
    public bool Equals(DecoratorAttributeBase x, DecoratorAttributeBase y) => x.GetType() == y.GetType();

    /// <summary>
    /// 获取哈希编码
    /// </summary>
    /// <param name="obj">对象</param>
    public int GetHashCode(DecoratorAttributeBase obj) => obj.GetType().GetHashCode();
}
