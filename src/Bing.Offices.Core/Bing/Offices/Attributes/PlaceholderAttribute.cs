using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 普通占位符特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class PlaceholderAttribute : Attribute
    {
        /// <summary>
        /// 占位符
        /// </summary>
        public string Placeholder { get; set; }

        /// <summary>
        /// 初始化一个<see cref="PlaceholderAttribute"/>类型的实例
        /// </summary>
        /// <param name="placeholder">占位符</param>
        public PlaceholderAttribute(string placeholder)
        {
            Placeholder = placeholder;
        }
    }
}
