using System.Collections.Generic;
using Bing.Offices.Attributes.Decorators;

namespace Bing.Offices.Decorators
{
    /// <summary>
    /// 类型装饰器信息
    /// </summary>
    public class TypeDecoratorInfo
    {
        /// <summary>
        /// 类型装饰器列表
        /// </summary>
        public List<DecorateBaseAttribute> TypeDecorates { get; set; }

        /// <summary>
        /// 属性装饰器列表
        /// </summary>
        public List<PropertyDecoratorInfo> PropertyDecorators { get; set; }
    }
}
