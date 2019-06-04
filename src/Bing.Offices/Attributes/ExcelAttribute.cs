using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// Excel 定义。用于识别Xml配置名称
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class ExcelAttribute : Attribute
    {
        /// <summary>
        /// 名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">名称</param>
        public ExcelAttribute(string name)
        {
            Name = name;
        }
    }
}
