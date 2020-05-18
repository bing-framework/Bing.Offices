namespace Bing.Offices.Exports
{
    /// <summary>
    /// 导出图片字段
    /// </summary>
    public interface IExportImageField
    {
        /// <summary>
        /// 高度
        /// </summary>
        int Height { get; set; }

        /// <summary>
        /// 宽度
        /// </summary>
        int Width { get; set; }

        /// <summary>
        /// 图片不存在时替换文本
        /// </summary>
        string Alt { get; set; }
    }
}
