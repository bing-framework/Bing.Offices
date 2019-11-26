namespace Bing.Offices.Core.Models
{
    /// <summary>
    /// Excel 头部样式
    /// </summary>
    public class ExcelHeadStyle
    {
        /// <summary>
        /// 字体大小
        /// </summary>
        public float FontSize { get; set; } = 11;

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 是否自适应
        /// </summary>
        public bool IsAutoFit { get; set; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }
    }
}
