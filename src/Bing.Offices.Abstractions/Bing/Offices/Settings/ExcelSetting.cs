namespace Bing.Offices.Settings
{
    /// <summary>
    /// Excel设置
    /// </summary>
    public sealed class ExcelSetting
    {
        /// <summary>
        /// 默认Excel设置
        /// </summary>
        private static ExcelSetting _defaultExcelSetting = new ExcelSetting();

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
        /// 主题
        /// </summary>
        public string Subject { get; set; } = "Bing.Offices";

        /// <summary>
        /// 类别
        /// </summary>
        public string Category { get; set; } = "Bing.Offices";

        /// <summary>
        /// 备注
        /// </summary>
        public string Description { get; set; } = "Bing.Offices 生成";

        /// <summary>
        /// 默认设置
        /// </summary>
        public static ExcelSetting Default
        {
            get => _defaultExcelSetting;
            set
            {
                if (value != null)
                    _defaultExcelSetting = value;
            }
        }
    }
}
