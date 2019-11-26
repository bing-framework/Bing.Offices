using System;

namespace Bing.Offices.Core
{
    /// <summary>
    /// 导出头部 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportHeaderAttribute : Attribute
    {
        /// <summary>
        /// 初始化一个<see cref="ExportHeaderAttribute"/>类型的实例
        /// </summary>
        /// <param name="displayName">显示名称</param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="format">格式化</param>
        /// <param name="isBold">是否加粗</param>
        /// <param name="isAutoFit">是否自适应</param>
        public ExportHeaderAttribute(string displayName = null, float fontSize = 11, string format = null,
            bool isBold = true, bool isAutoFit = true)
        {
            DisplayName = displayName;
            FontSize = fontSize;
            Format = format;
            IsBold = isBold;
            IsAutoFit = isAutoFit;
        }

        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { get; set; }

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
    }
}
