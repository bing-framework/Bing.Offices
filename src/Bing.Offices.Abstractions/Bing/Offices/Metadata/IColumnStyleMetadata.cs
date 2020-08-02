namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义数据列样式元数据
    /// </summary>
    public interface IColumnStyleMetadata
    {
        /// <summary>
        /// 宽度
        /// </summary>
        int Width { get; set; }

        /// <summary>
        /// 高度
        /// </summary>
        int Height { get; set; }

        /// <summary>
        /// 是否自动宽度
        /// </summary>
        bool IsAutoWidth { get; set; }
    }
}
