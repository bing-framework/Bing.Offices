using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices.Settings
{
    /// <summary>
    /// Excel 文档属性设置
    /// </summary>
    public sealed class ExcelSetting : IExcelSetting
    {
        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; } = "jian玄冰";

        /// <summary>
        /// 公司
        /// </summary>
        public string Company { get; set; } = "jianxuanbing";

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; } = "Bing.Offices 生成";

        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; } = "Bing.Offices 生成";

        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; } = "Bing.Offices 生成";

        /// <summary>
        /// 目录
        /// </summary>
        public string Category { get; set; } = "Bing.Offices";

        /// <summary>
        /// 最大导出数量
        /// </summary>
        public int MaxExportNum { get; set; } = 10000;
    }
}
