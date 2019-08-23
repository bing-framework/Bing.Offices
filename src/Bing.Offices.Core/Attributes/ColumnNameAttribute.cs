using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 列名特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnNameAttribute : Attribute
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ColumnNameAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">名称</param>
        public ColumnNameAttribute(string name) => Name = name;
    }
}
