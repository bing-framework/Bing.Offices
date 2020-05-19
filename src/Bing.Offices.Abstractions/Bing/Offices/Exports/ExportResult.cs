namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出结果
    /// </summary>
    public class ExportResult
    {
        /// <summary>
        /// 文件名。例如：xxx.xls
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// 文件路径。例如：/files/export/xxx.xls
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 扩展名。例如：.xls
        /// </summary>
        public string Extension { get; set; }

        /// <summary>
        /// 绝对文件路径
        /// </summary>
        public string AbsFilePath { get; set; }
    }
}
