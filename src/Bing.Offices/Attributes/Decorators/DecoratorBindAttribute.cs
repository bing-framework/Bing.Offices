using System;

namespace Bing.Offices.Attributes.Decorators
{
    /// <summary>
    /// 装饰器绑定
    /// </summary>
    public class DecoratorBindAttribute : Attribute
    {
        /// <summary>
        /// 装饰器类型
        /// </summary>
        public Type DecoratorType { get; set; }

        /// <summary>
        /// 初始化一个<see cref="DecoratorBindAttribute"/>类型的实例
        /// </summary>
        /// <param name="decoratorType">装饰器类型</param>
        public DecoratorBindAttribute(Type decoratorType)
        {
            DecoratorType = decoratorType;
        }
    }
}
