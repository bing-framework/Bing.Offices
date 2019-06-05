namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 属性设置
    /// </summary>
    public interface IPropertySetting
    {
        /// <summary>
        /// 列索引
        /// </summary>
        int Index { get; }

        /// <summary>
        /// 列标题
        /// </summary>
        string Title { get; }

        /// <summary>
        /// 列格式化程序
        /// </summary>
        string Formatter { get; }

        /// <summary>
        /// 是否忽略属性
        /// </summary>
        bool Ignored { get; }

        /// <summary>
        /// 默认值
        /// </summary>
        object DefaultValue { get; }
    }
}
