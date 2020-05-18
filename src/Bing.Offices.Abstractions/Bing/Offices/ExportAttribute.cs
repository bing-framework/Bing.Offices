using System;
using Bing.Offices.Filters;

namespace Bing.Offices
{
    /// <summary>
    /// 导出 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExportAttribute : Attribute
    {
        /// <summary>
        /// 名称（比如当前Sheet名称）
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

        /// <summary>
        /// 一个Sheet最大允许的行数，设置了只会将输出多个Sheet
        /// </summary>
        public int MaxRowNumberOnASheet { get; set; } = 0;

        /// <summary>
        /// 表格样式风格
        /// </summary>
        public string TableStyle { get; set; } = "Medium10";

        /// <summary>
        /// 自适应所有列
        /// </summary>
        public bool AutoFitAllColumn { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// 导出列头筛选器 <br />
        /// 必须实现【<see cref="IExportHeaderFilter"/>】
        /// </summary>
        public Type ExportHeaderFilter { get; set; }
    }
}
