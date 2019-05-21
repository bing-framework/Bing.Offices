using System.Collections.Generic;
using Bing.Offices.Attributes.Decorators;

namespace Bing.Offices.Decorators
{
    /// <summary>
    /// 装饰器特性比较器
    /// </summary>
    public class DecoratorAttributeComparer:IEqualityComparer<DecorateBaseAttribute>
    {
        /// <summary>
        /// 是否相等
        /// </summary>
        public bool Equals(DecorateBaseAttribute x, DecorateBaseAttribute y)
        {
            return x.GetType() == y.GetType();
        }

        /// <summary>
        /// 获取哈希值
        /// </summary>
        public int GetHashCode(DecorateBaseAttribute obj)
        {
            return obj.GetType().GetHashCode();
        }
    }
}
