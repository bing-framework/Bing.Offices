using System;
using Bing.Offices.Exports;

namespace Bing.Offices.Fields
{
    /// <summary>
    /// 导出图片字段 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportImageFieldAttribute : Attribute, IExportImageField
    {
        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// 图片不存在时替换文本
        /// </summary>
        public string Alt { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExportImageFieldAttribute"/>类型的实例
        /// </summary>
        /// <param name="height">高度</param>
        /// <param name="width">宽度</param>
        /// <param name="alt">图片不存在时替换文本</param>
        public ExportImageFieldAttribute(int height = 15, int width = 50, string alt = null)
        {
            Height = height;
            Width = width;
            Alt = alt;
        }
    }
}
