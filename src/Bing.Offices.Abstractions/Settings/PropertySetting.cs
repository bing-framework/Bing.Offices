namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 属性设置
    /// </summary>
    public sealed class PropertySetting
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 列标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 列格式化程序
        /// </summary>
        public string Formatter { get; set; }

        /// <summary>
        /// 是否忽略属性
        /// </summary>
        public bool Ignored { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; set; }
    }
}
