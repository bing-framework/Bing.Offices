namespace Bing.Offices.Imports
{
    /// <summary>
    /// 导入图片字段
    /// </summary>
    public interface IImportImageField
    {
        /// <summary>
        /// 图片存储路径（默认存储到临时目录）
        /// </summary>
        string ImageDirectory { get; set; }

        /// <summary>
        /// 图片导入方式（默认Base64）
        /// </summary>
        ImportImageTo ImportImageTo { get; set; }
    }
}
