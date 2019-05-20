using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 自动换行
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class WrapTextAttribute : Attribute
    {
    }
}
