namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出头部
    /// </summary>
    public interface IExportHeader
    {
        /// <summary>
        /// 显示名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        float? FontSize { get; set; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        bool IsBold { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        string Format { get; set; }

        /// <summary>
        /// 是否自适应
        /// </summary>
        bool IsAutoFit { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        bool IsIgnore { get; set; }
    }
}
