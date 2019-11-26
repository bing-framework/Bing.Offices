using System;

namespace Bing.Offices.Core
{
    /// <summary>
    /// 导出 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExportAttribute : Attribute
    {
        /// <summary>
        /// 名称。比如当前Sheet名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 头部字体大小
        /// </summary>
        public float? HeaderFontSize { get; set; }

        /// <summary>
        /// 正文字体大小
        /// </summary>
        public float? FontSize { get; set; }
    }
}
