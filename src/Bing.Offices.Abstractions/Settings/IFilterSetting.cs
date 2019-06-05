namespace Bing.Offices.Abstractions.Settings
{
    /// <summary>
    /// 过滤器设置
    /// </summary>
    public interface IFilterSetting
    {
        /// <summary>
        /// 首列索引
        /// </summary>
        int FirstColumn { get; }

        /// <summary>
        /// 最后一列索引
        /// </summary>
        int? LastColumn { get; }
    }
}
