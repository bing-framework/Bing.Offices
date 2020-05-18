using System;
using System.IO;
using Bing.Offices.Imports;

namespace Bing.Offices.Fields
{
    /// <summary>
    /// 导入图片字段 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ImportImageFieldAttribute : Attribute, IImportImageField
    {
        /// <summary>
        /// 图片存储路径（默认存储到临时目录）
        /// </summary>
        public string ImageDirectory { get; set; } = Path.GetTempPath();

        /// <summary>
        /// 图片导入方式（默认Base64）
        /// </summary>
        public ImportImageTo ImportImageTo { get; set; } = ImportImageTo.Base64;

        /// <summary>
        /// 初始化一个<see cref="ImportImageFieldAttribute"/>类型的实例
        /// </summary>
        public ImportImageFieldAttribute() { }

        /// <summary>
        /// 初始化一个<see cref="ImportImageFieldAttribute"/>类型的实例
        /// </summary>
        /// <param name="imageDirectory">图片存储路径</param>
        public ImportImageFieldAttribute(string imageDirectory)
        {
            ImportImageTo = ImportImageTo.TempFolder;
            ImageDirectory = imageDirectory ?? Path.GetTempPath();
        }
    }
}
