namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义工作表元数据
    /// </summary>
    public interface ISheetMetadata
    {
        /// <summary>
        /// 索引
        /// </summary>
        int Index { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; set; }

        /// <summary>
        /// 标题
        /// </summary>
        string Title { get; set; }

        /// <summary>
        /// 是否拥有标题
        /// </summary>
        bool IsHasTitle { get; set; }

        /// <summary>
        /// 标题样式
        /// </summary>
        ICellStyleMetadata TitleStyle { get; set; }

        /// <summary>
        /// 表头样式
        /// </summary>
        ICellStyleMetadata HeaderStyle { get; set; }

        /// <summary>
        /// 正文样式
        /// </summary>
        ICellStyleMetadata BodyStyle { get; set; }

        /// <summary>
        /// 页脚样式
        /// </summary>
        ICellStyleMetadata FooterStyle { get; set; }
    }
}
