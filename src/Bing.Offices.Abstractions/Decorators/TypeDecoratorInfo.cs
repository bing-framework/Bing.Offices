using System.Collections.Generic;
using Bing.Offices.Abstractions.Attributes;

namespace Bing.Offices.Abstractions.Decorators
{
    /// <summary>
    /// 类型装饰器信息
    /// </summary>
    public class TypeDecoratorInfo
    {
        /// <summary>
        /// 类型装饰器列表
        /// </summary>
        public IList<DecorateAttributeBase> TypeDecorators { get; set; }

        /// <summary>
        /// 属性装饰器信息列表
        /// </summary>
        public IList<PropertyDecoratorInfo> PropertyDecoratorInfos { get; set; }
    }
}
