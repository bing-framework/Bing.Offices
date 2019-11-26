namespace Bing.Offices.Core.Models
{
    /// <summary>
    /// 模板文件信息
    /// </summary>
    public class TemplateFileInfo
    {
        /// <summary>
        /// 初始化一个<see cref="TemplateFileInfo"/>类型的实例
        /// </summary>
        public TemplateFileInfo() { }

        /// <summary>
        /// 初始化一个<see cref="TemplateFileInfo"/>类型的实例
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="fileType">文件类型</param>
        public TemplateFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }

        /// <summary>
        /// 文件名
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件类型
        /// </summary>
        public string FileType { get; set; }
    }
}
