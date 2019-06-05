using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices.Settings
{
    /// <summary>
    /// 属性设置
    /// </summary>
    public sealed class PropertySetting : IPropertySetting
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; internal set; }

        /// <summary>
        /// 列标题
        /// </summary>
        public string Title { get; internal set; }

        /// <summary>
        /// 列格式化程序
        /// </summary>
        public string Formatter { get; internal set; }

        /// <summary>
        /// 是否忽略属性
        /// </summary>
        public bool Ignored { get; internal set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; internal set; }
    }
}
