namespace Bing.Offices.Abstractions.Configurations
{
    /// <summary>
    /// 属性配置
    /// </summary>
    public interface IPropertyConfiguration
    {
        /// <summary>
        /// 设置索引
        /// </summary>
        /// <param name="index">索引</param>
        IPropertyConfiguration HasColumnIndex(int index);

        /// <summary>
        /// 设置标题
        /// </summary>
        /// <param name="title">标题</param>
        IPropertyConfiguration HasColumnTitle(string title);

        /// <summary>
        /// 设置格式化程序
        /// </summary>
        /// <param name="formatter">格式化程序</param>
        IPropertyConfiguration HasColumnFormatter(string formatter);

        /// <summary>
        /// 设置忽略
        /// </summary>
        IPropertyConfiguration Ignored();

        /// <summary>
        /// 设置默认值
        /// </summary>
        /// <param name="value">默认值</param>
        IPropertyConfiguration DefaultValue(object value);
    }
}
