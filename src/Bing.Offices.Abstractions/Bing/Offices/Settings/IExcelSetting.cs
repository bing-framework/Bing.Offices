namespace Bing.Offices.Settings
{
    /// <summary>
    /// Excel 文档属性设置
    /// </summary>
    public interface IExcelSetting
    {
        /// <summary>
        /// 作者
        /// </summary>
        string Author { get; set; }

        /// <summary>
        /// 公司
        /// </summary>
        string Company { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; set; }

        /// <summary>
        /// 描述
        /// </summary>
        string Description { get; set; }

        /// <summary>
        /// 主题
        /// </summary>
        string Subject { get; set; }

        /// <summary>
        /// 类别
        /// </summary>
        string Category { get; set; }
    }
}
