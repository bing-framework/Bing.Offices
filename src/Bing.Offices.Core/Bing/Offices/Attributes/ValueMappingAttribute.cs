using System;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 值映射 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ValueMappingAttribute : Attribute
    {
        /// <summary>
        /// 初始化一个<see cref="ValueMappingAttribute"/>类型的实例
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="value">值</param>
        public ValueMappingAttribute(string text, object value)
        {
            Text = text;
            Value = value;
        }

        /// <summary>
        /// 文本
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// 值
        /// </summary>
        public object Value { get; }
    }
}
