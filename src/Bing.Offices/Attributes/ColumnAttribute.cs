using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        /// <summary>
        /// 列名
        /// </summary>
        public string ColumnName { get; set; }
    }
}
