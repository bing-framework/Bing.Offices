using System.Collections.Generic;
using Bing.Offices.Attributes;

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
        public IList<DecoratorAttributeBase> TypeDecorators { get; set; } = new List<DecoratorAttributeBase>();

        /// <summary>
        /// 属性装饰器信息列表
        /// </summary>
        public IList<PropertyDecoratorInfo> PropertyDecoratorInfos { get; set; } = new List<PropertyDecoratorInfo>();
    }
}
