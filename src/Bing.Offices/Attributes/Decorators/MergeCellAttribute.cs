using System;

namespace Bing.Offices.Attributes.Decorators
{
    /// <summary>
    /// 合并单元格
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class MergeCellAttribute : DecorateBaseAttribute
    {
    }
}
