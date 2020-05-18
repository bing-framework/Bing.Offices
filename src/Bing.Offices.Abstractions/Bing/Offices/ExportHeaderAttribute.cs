using System;
using Bing.Offices.Exports;

namespace Bing.Offices
{
    /// <summary>
    /// 导出头部 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportHeaderAttribute : Attribute, IExportHeader
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { get; set; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 是否自适应
        /// </summary>
        public bool IsAutoFit { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExportHeaderAttribute"/>类型的实例
        /// </summary>
        /// <param name="name">显示名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="format">格式化</param>
        /// <param name="isBold">是否加粗</param>
        /// <param name="isAutoFit">是否自适应</param>
        public ExportHeaderAttribute(string name = null, float fontSize = 11, string format = null, bool isBold = true, bool isAutoFit = true)
        {
            Name = name;
            FontSize = fontSize;
            Format = format;
            IsBold = isBold;
            IsAutoFit = isAutoFit;
        }
    }
}
