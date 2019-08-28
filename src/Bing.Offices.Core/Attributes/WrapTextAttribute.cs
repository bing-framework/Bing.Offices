using System;
using Bing.Offices.Abstractions.Attributes;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 自动换行特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class WrapTextAttribute : DecoratorAttributeBase
    {
    }
}
