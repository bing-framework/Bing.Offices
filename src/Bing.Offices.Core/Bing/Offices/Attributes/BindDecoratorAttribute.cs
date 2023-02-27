using System;

namespace Bing.Offices.Attributes;

/// <summary>
/// 绑定装饰器特性
/// </summary>
public class BindDecoratorAttribute : Attribute
{
    /// <summary>
    /// 装饰器类型
    /// </summary>
    public Type DecoratorType { get; set; }

    /// <summary>
    /// 初始化一个<see cref="BindDecoratorAttribute"/>类型的实例
    /// </summary>
    /// <param name="decoratorType">装饰器类型</param>
    public BindDecoratorAttribute(Type decoratorType)
    {
        DecoratorType = decoratorType;
    }
}