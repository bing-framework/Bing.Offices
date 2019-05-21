using System;

namespace Bing.Offices.Attributes.Filters
{
    /// <summary>
    /// 正则表达式
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = false)]
    public class RegexAttribute : FilterBaseAttribute
    {
        /// <summary>
        /// 正则表达式
        /// </summary>
        public string Regex { get; set; }

        /// <summary>
        /// 初始化一个<see cref="RegexAttribute"/>类型的实例
        /// </summary>
        /// <param name="regex">正则表达式</param>
        public RegexAttribute(string regex)
        {
            Regex = regex;
            ErrorMessage = "无效内容";
        }
    }
}
