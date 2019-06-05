namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// Excel 文档属性设置
    /// </summary>
    public interface IExcelSetting
    {
        /// <summary>
        /// 作者
        /// </summary>
        string Author { get; }

        /// <summary>
        /// 公司
        /// </summary>
        string Company { get; }

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; }

        /// <summary>
        /// 描述
        /// </summary>
        string Description { get; }

        /// <summary>
        /// 主题
        /// </summary>
        string Subject { get; }

        /// <summary>
        /// 目录
        /// </summary>
        string Category { get; }

        /// <summary>
        /// 最大导出数量
        /// </summary>
        int MaxExportNum { get; }
    }
}
