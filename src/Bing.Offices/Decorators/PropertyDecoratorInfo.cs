using System.Collections.Generic;
using Bing.Offices.Attributes.Decorators;

namespace Bing.Offices.Decorators
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
        public List<DecorateBaseAttribute> Decorators { get; set; }
    }
}
