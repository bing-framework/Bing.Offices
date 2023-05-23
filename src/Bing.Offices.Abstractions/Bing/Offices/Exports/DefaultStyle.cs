using Bing.Offices.Enums.Styles;

namespace Bing.Offices.Exports
{
    /// <summary>
    /// 默认样式
    /// </summary>
    public class DefaultStyle
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        public const string FontName = "宋体";

        /// <summary>
        /// 字体颜色
        /// </summary>
        public const short FontColor = 8;

        /// <summary>
        /// 字体大小
        /// </summary>
        public const short FontSize = 11;

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public const HorizontalAlignment HorizontalAlignment = Enums.Styles.HorizontalAlignment.Left;

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public const VerticalAlignment VerticalAlignment = Enums.Styles.VerticalAlignment.Center;
    }
}
