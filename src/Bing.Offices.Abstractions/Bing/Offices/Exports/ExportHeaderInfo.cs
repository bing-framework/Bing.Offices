namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出列头信息
    /// </summary>
    public class ExportHeaderInfo
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列属性
        /// </summary>
        public IExportHeader Header { get; set; }

        /// <summary>
        /// 图片属性
        /// </summary>
        public IExportImageField ImageField { get; set; }

        /// <summary>
        /// C#数据类型
        /// </summary>
        public string CsTypeName { get; set; }

        /// <summary>
        /// 显示名称。最终显示的列名
        /// </summary>
        public string DisplayName { get; set; }
    }
}
