namespace Bing.Offices.Parameters
{
    /// <summary>
    /// 模板导出参数
    /// </summary>
    public interface ITemplateExportParam : ITemplateParam, IExportParam
    {
        /// <summary>
        /// 是否插入新行
        /// </summary>
        bool InsertNewLine { get; set; }

        /// <summary>
        /// 数据填充方向
        /// </summary>
        DataDirection DataDirection { get; set; }

        /// <summary>
        /// 是否复制单元格样式。默认：true
        /// </summary>
        bool CopyCellStyle { get; set; }
    }
}
