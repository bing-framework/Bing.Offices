namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出文件信息
    /// </summary>
    public class ExportFileInfo
    {
        /// <summary>
        /// 文件名（路径）
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件Mime类型
        /// </summary>
        public string FileType { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExportFileInfo"/>类型的实例
        /// </summary>
        public ExportFileInfo() { }

        /// <summary>
        /// 初始化一个<see cref="ExportFileInfo"/>类型的实例
        /// </summary>
        /// <param name="fileName">文件名</param>
        /// <param name="fileType">文件Mime类型</param>
        public ExportFileInfo(string fileName, string fileType)
        {
            FileName = fileName;
            FileType = fileType;
        }
    }
}
