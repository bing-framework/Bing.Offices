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
        public string Author { get; set; } = "简玄冰";

        /// <summary>
        /// 公司
        /// </summary>
        public string Company { get; set; } = "简玄冰";

        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; } = "Bing.Offices";

        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; } = "Bing.Offices";

        /// <summary>
        /// 主题
        /// </summary>
        public string Subject { get; set; } = "Bing.Offices";

        /// <summary>
        /// 类别
        /// </summary>
        public string Category { get; set; } = "Bing.Offices";

        /// <summary>
        /// 默认 Excel 文档属性设置
        /// </summary>
        public static ExcelSetting Default = new ExcelSetting();
    }
}
