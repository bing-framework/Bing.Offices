using System;

namespace Bing.Offices.Attributes.Decorators
{
    /// <summary>
    /// 自动换行
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class WrapTextAttribute : DecorateBaseAttribute
    {
    }
}
