using System.Collections.Generic;
using Bing.Offices.Abstractions.Attributes;

namespace Bing.Offices.Decorators
{
    /// <summary>
    /// 装饰器特性比较器
    /// </summary>
    internal class DecoratorAttributeComparer:IEqualityComparer<DecorateAttributeBase>
    {
        /// <summary>
        /// 是否相等
        /// </summary>
        /// <param name="x">对象A</param>
        /// <param name="y">对象B</param>
        public bool Equals(DecorateAttributeBase x, DecorateAttributeBase y) => x.GetType() == y.GetType();

        /// <summary>
        /// 获取哈希编码
        /// </summary>
        /// <param name="obj">对象</param>
        public int GetHashCode(DecorateAttributeBase obj) => obj.GetType().GetHashCode();
    }
}
