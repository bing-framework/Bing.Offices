namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 模板导出参数
    /// </summary>
    public class TemplateExportParam : ExportParamBase, ITemplateExportParam
    {
        /// <summary>
        /// 数据源名称
        /// </summary>
        public string DataSourceName { get; set; }

        /// <summary>
        /// 是否插入新行
        /// </summary>
        public bool InsertNewLine { get; set; }

        /// <summary>
        /// 数据填充方向。默认：DataDirection.None
        /// </summary>
        public DataDirection DataDirection { get; set; } = DataDirection.None;

        /// <summary>
        /// 是否复制单元格样式。默认：true
        /// </summary>
        public bool CopyCellStyle { get; set; } = true;

        /// <summary>
        /// 初始化一个<see cref="TemplateExportParam"/>类型的实例
        /// </summary>
        /// <param name="dataSourceName">数据源名称</param>
        /// <param name="dataSource">数据源</param>
        public TemplateExportParam(string dataSourceName, object dataSource)
        {
            DataSourceName = dataSourceName;
            DataSource = dataSource;
        }
    }
}
