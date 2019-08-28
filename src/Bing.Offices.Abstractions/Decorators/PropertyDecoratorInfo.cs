using System.Collections.Generic;
using Bing.Offices.Abstractions.Attributes;

namespace Bing.Offices.Abstractions.Decorators
{
    /// <summary>
    /// 属性装饰器信息
    /// </summary>
    public class PropertyDecoratorInfo
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 装饰器列表
        /// </summary>
        public IList<DecoratorAttributeBase> Decorators { get; set; } = new List<DecoratorAttributeBase>();
    }
}
